import { describe, expect, it } from 'vitest';

import { operationContext, type OperationScenario } from './operation.fixtures';
import { router } from './router';

const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';
const GROUP = '4b8b8d7e-1a2b-4c3d-9e0f-5a6b7c8d9e0f';
const SKU = '05e9a617-0261-4cee-bb44-138d3ef5d965';

async function run(scenario: OperationScenario) {
	const { ctx, requests } = operationContext(scenario);
	const [output] = await router.call(ctx);
	return { output, requests };
}

describe('dispatch', () => {
	it.each([
		['user', 'get', `https://graph.microsoft.com/v1.0/users/${USER}`, 'GET'],
		['user', 'delete', `https://graph.microsoft.com/v1.0/users/${USER}`, 'DELETE'],
		[
			'user',
			'revokeSessions',
			`https://graph.microsoft.com/v1.0/users/${USER}/revokeSignInSessions`,
			'POST',
		],
		[
			'authentication',
			'getAllMethods',
			`https://graph.microsoft.com/v1.0/users/${USER}/authentication/methods`,
			'GET',
		],
		['license', 'queryTenant', 'https://graph.microsoft.com/v1.0/subscribedSkus', 'GET'],
		[
			'license',
			'queryUser',
			`https://graph.microsoft.com/v1.0/users/${USER}/licenseDetails`,
			'GET',
		],
	])('routes %s.%s to the right request', async (resource, operation, url, method) => {
		const { requests } = await run({
			parameters: { resource, operation, user: USER, output: 'raw' },
		});

		expect(requests[0].method).toBe(method);
		expect(requests[0].url).toBe(url);
	});

	it('rejects a resource it does not know', async () => {
		await expect(run({ parameters: { resource: 'nope', operation: 'get' } })).rejects.toThrow(
			/resource "nope" is not known/,
		);
	});
});

describe('the per-item loop', () => {
	it('runs once per input item, in order', async () => {
		const { output, requests } = await run({
			items: 3,
			parameters: {
				resource: 'user',
				operation: 'delete',
				user: (index: number) => `user-${index}`,
			},
		});

		expect(requests.map((r) => r.url.split('/users/')[1])).toEqual(['user-0', 'user-1', 'user-2']);
		expect(output.map((item) => item.pairedItem)).toEqual([{ item: 0 }, { item: 1 }, { item: 2 }]);
	});

	it('stops at the first failure when continue-on-fail is off', async () => {
		await expect(
			run({
				items: 3,
				parameters: {
					resource: 'user',
					operation: 'delete',
					user: (index: number) => (index === 1 ? '' : `user-${index}`),
				},
			}),
		).rejects.toThrow(/No user was given/);
	});

	it('reports the failure per item and keeps going when continue-on-fail is on', async () => {
		const { output, requests } = await run({
			items: 3,
			continueOnFail: true,
			parameters: {
				resource: 'user',
				operation: 'delete',
				user: (index: number) => (index === 1 ? '' : `user-${index}`),
			},
		});

		expect(requests).toHaveLength(2);
		expect(output).toHaveLength(3);
		expect(output[1].json.error).toMatch(/No user was given/);
		expect(output[1].pairedItem).toEqual({ item: 1 });
	});
});

describe('whole-run license writes', () => {
	it.each(['assign', 'assignGroup', 'unassign'])(
		'runs %s once for the whole input instead of per item',
		async (operation) => {
			const target = operation === 'assignGroup' ? { group: GROUP } : { user: USER };

			const { requests } = await run({
				items: 3,
				parameters: {
					resource: 'license',
					operation,
					...target,
					skuId: [SKU],
					options: {},
				},
			});

			// Three items aimed at the same target fold into a single Graph write.
			expect(requests).toHaveLength(1);
			expect(requests[0].url).toContain('/assignLicense');
		},
	);

	it('still answers every input item', async () => {
		const { output } = await run({
			items: 3,
			parameters: {
				resource: 'license',
				operation: 'assign',
				user: USER,
				skuId: [SKU],
				options: {},
			},
		});

		expect(output).toHaveLength(3);
	});

	it('leaves the license read operations on the per-item path', async () => {
		const { requests } = await run({
			items: 2,
			parameters: { resource: 'license', operation: 'queryTenant' },
		});

		expect(requests).toHaveLength(2);
	});
});
