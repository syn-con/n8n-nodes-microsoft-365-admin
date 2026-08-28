import { describe, expect, it } from 'vitest';

import { operationContext } from '../operation.fixtures';
import { execute } from './create.operation';

const GROUP_ID = '4b8b8d7e-1a2b-4c3d-9e0f-5a6b7c8d9e0f';

const BASE = {
	displayName: 'Engineering',
	mailNickname: 'engineering',
};

async function run(
	parameters: Record<string, unknown>,
	responses?: Parameters<typeof operationContext>[0]['responses'],
) {
	const { ctx, requests } = operationContext({
		parameters: { ...BASE, ...parameters },
		responses: responses ?? [{ body: { id: GROUP_ID } }],
	});
	const output = await execute.call(ctx, 0);
	return { output, requests };
}

describe('group create', () => {
	it('creates a security group with the flags Graph requires but does not show', async () => {
		const { requests } = await run({ groupType: '' });

		expect(requests[0]).toMatchObject({
			method: 'POST',
			url: 'https://graph.microsoft.com/v1.0/groups',
		});
		expect(requests[0].body).toEqual({
			displayName: 'Engineering',
			mailNickname: 'engineering',
			mailEnabled: false,
			securityEnabled: true,
		});
	});

	it('creates a Microsoft 365 group from the visible flags', async () => {
		const { requests } = await run({
			groupType: 'Unified',
			mailEnabled: true,
			securityEnabled: false,
		});

		expect(requests[0].body).toMatchObject({
			groupTypes: ['Unified'],
			mailEnabled: true,
			securityEnabled: false,
		});
	});

	it('adds DynamicMembership to groupTypes alongside the group type', async () => {
		const { requests } = await run({ groupType: 'Unified', membershipType: 'DynamicMembership' });
		expect(requests[0].body.groupTypes).toEqual(['Unified', 'DynamicMembership']);
	});

	it('sends the create-time additional fields with the create call', async () => {
		const { requests } = await run({
			groupType: '',
			membershipType: 'DynamicMembership',
			additionalFields: {
				membershipRule: 'user.department -eq "R&D"',
				membershipRuleProcessingState: 'On',
			},
		});

		expect(requests).toHaveLength(1);
		expect(requests[0].body).toMatchObject({
			membershipRule: 'user.department -eq "R&D"',
			membershipRuleProcessingState: 'On',
		});
	});

	it('applies the remaining additional fields in a follow-up PATCH', async () => {
		const { requests, output } = await run({
			groupType: '',
			additionalFields: { description: 'The engineering group', visibility: 'Private' },
		});

		expect(requests).toHaveLength(2);
		expect(requests[1]).toMatchObject({
			method: 'PATCH',
			url: `https://graph.microsoft.com/v1.0/groups/${GROUP_ID}`,
		});
		expect(requests[1].body).toEqual({
			description: 'The engineering group',
			visibility: 'Private',
		});
		// The follow-up values are merged into what the create returned.
		expect(output[0].json).toMatchObject({ id: GROUP_ID, description: 'The engineering group' });
	});

	it('deletes the half-configured group when the follow-up PATCH fails', async () => {
		const { ctx, requests } = operationContext({
			parameters: {
				...BASE,
				groupType: '',
				additionalFields: { description: 'nope' },
			},
			responses: [
				{ body: { id: GROUP_ID } },
				{ statusCode: 400, body: { error: { code: 'BadRequest', message: 'no' } } },
			],
		});

		await expect(execute.call(ctx, 0)).rejects.toThrow();

		expect(requests.map((r) => r.method)).toEqual(['POST', 'PATCH', 'DELETE']);
		expect(requests[2].url).toBe(`https://graph.microsoft.com/v1.0/groups/${GROUP_ID}`);
	});
});

describe('group create validation', () => {
	it('rejects a display name over 256 characters', async () => {
		await expect(run({ groupType: '', displayName: 'a'.repeat(257) })).rejects.toThrow(
			/maximum length of 256/,
		);
	});

	it('rejects a mail nickname that carries a domain', async () => {
		await expect(run({ groupType: '', mailNickname: 'eng@contoso.com' })).rejects.toThrow(
			/without @contoso\.com/,
		);
	});

	it('rejects a mail nickname over 64 characters', async () => {
		await expect(run({ groupType: '', mailNickname: 'a'.repeat(65) })).rejects.toThrow(
			/maximum length of 64/,
		);
	});

	it('rejects a mail nickname with characters outside the allowed ASCII set', async () => {
		await expect(run({ groupType: '', mailNickname: 'eng team' })).rejects.toThrow(
			/ASCII character set/,
		);
	});

	it.each([
		[
			'security disabled',
			{
				groupType: '',
				securityEnabled: false,
				additionalFields: { isAssignableToRole: true, visibility: 'Private' },
			},
			/'Security Enabled' must be set to true/,
		],
		[
			'not private',
			{ groupType: '', additionalFields: { isAssignableToRole: true, visibility: 'Public' } },
			/'Visibility' must be set to 'Private'/,
		],
		[
			'unified without mail',
			{
				groupType: 'Unified',
				mailEnabled: false,
				additionalFields: { isAssignableToRole: true, visibility: 'Private' },
			},
			/'Mail Enabled' must be set to true/,
		],
	])('refuses assignable-to-role when %s', async (_label, parameters, match) => {
		await expect(run(parameters)).rejects.toThrow(match);
	});
});
