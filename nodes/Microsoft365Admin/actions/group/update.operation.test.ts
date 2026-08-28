import { describe, expect, it } from 'vitest';

import { operationContext } from '../operation.fixtures';
import { execute } from './update.operation';

const GROUP = '4b8b8d7e-1a2b-4c3d-9e0f-5a6b7c8d9e0f';
const URL = `https://graph.microsoft.com/v1.0/groups/${GROUP}`;

async function run(updateFields: Record<string, unknown>) {
	const { ctx, requests } = operationContext({ parameters: { group: GROUP, updateFields } });
	const output = await execute.call(ctx, 0);
	return { output, requests };
}

describe('group update', () => {
	it('PATCHes the ordinary fields in one request', async () => {
		const { requests, output } = await run({ description: 'new', visibility: 'Private' });

		expect(requests).toHaveLength(1);
		expect(requests[0]).toMatchObject({ method: 'PATCH', url: URL });
		expect(requests[0].body).toEqual({ description: 'new', visibility: 'Private' });
		expect(output[0].json).toEqual({ updated: true });
	});

	it('sends the fields Graph refuses to batch in a request of their own', async () => {
		const { requests } = await run({ description: 'new', allowExternalSenders: true });

		expect(requests).toHaveLength(2);
		expect(requests[0].body).toEqual({ description: 'new' });
		expect(requests[1].body).toEqual({ allowExternalSenders: true });
	});

	it('sends only the separate request when nothing else changed', async () => {
		const { requests } = await run({ autoSubscribeNewMembers: true });

		expect(requests).toHaveLength(1);
		expect(requests[0].body).toEqual({ autoSubscribeNewMembers: true });
	});

	it('sends nothing at all for a no-op update', async () => {
		// Declarative routing had to send an empty PATCH and swallow Graph's 400.
		const { requests, output } = await run({});

		expect(requests).toHaveLength(0);
		expect(output[0].json).toEqual({ updated: true });
	});

	it('rejects a mail nickname that carries a domain', async () => {
		await expect(run({ mailNickname: 'eng@contoso.com' })).rejects.toThrow(/without @contoso\.com/);
	});
});
