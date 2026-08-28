import { describe, expect, it } from 'vitest';

import { operationContext } from '../operation.fixtures';
import { execute } from './queryHolders.operation';

const SKU = '05e9a617-0261-4cee-bb44-138d3ef5d965';

async function run(parameters: Record<string, unknown>) {
	const { ctx, requests } = operationContext({
		parameters: { skuId: SKU, returnAll: false, limit: 50, ...parameters },
		responses: [{ body: { value: [{ id: 'u1' }] } }],
	});
	const output = await execute.call(ctx, 0);
	return { output, requests };
}

describe('license queryHolders', () => {
	it('filters users on the SKU with the advanced-query headers Graph demands', async () => {
		const { requests } = await run({ holderType: 'users' });

		expect(requests[0].url).toBe('https://graph.microsoft.com/v1.0/users');
		expect(requests[0].qs?.$filter).toBe(`assignedLicenses/any(license:license/skuId eq ${SKU})`);
		expect(requests[0].qs?.$count).toBe('true');
		expect(requests[0].headers?.ConsistencyLevel).toBe('eventual');
	});

	it('reads the group collection when asked for group holders', async () => {
		const { requests } = await run({ holderType: 'groups' });

		expect(requests[0].url).toBe('https://graph.microsoft.com/v1.0/groups');
		expect(requests[0].qs?.$select).toContain('licenseProcessingState');
	});

	it('projects the assignment state for users', async () => {
		const { requests } = await run({ holderType: 'users' });
		expect(requests[0].qs?.$select).toContain('licenseAssignmentStates');
	});

	it('falls back to the user projection for an unrecognised holder type', async () => {
		const { requests } = await run({ holderType: 'somethingElse' });
		expect(requests[0].url).toBe('https://graph.microsoft.com/v1.0/users');
		expect(requests[0].qs?.$select).toContain('userPrincipalName');
	});

	it('caps the page at the limit when Return All is off', async () => {
		const { requests } = await run({ holderType: 'users', limit: 5 });
		expect(requests[0].qs?.$top).toBe(5);
	});

	it('unwraps the collection into one item per holder', async () => {
		const { output } = await run({ holderType: 'users' });
		expect(output.map((item) => item.json.id)).toEqual(['u1']);
	});
});
