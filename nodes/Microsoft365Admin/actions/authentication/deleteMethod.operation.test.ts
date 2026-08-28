import { describe, expect, it } from 'vitest';

import { operationContext } from '../operation.fixtures';
import { execute } from './deleteMethod.operation';

const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';
const METHOD = '-2_GRUg2-HYz6_1YG4YRAQ2';

async function run(parameters: Record<string, unknown>) {
	const { ctx, requests } = operationContext({
		parameters: { user: USER, method: METHOD, methodType: 'fido2Methods', ...parameters },
	});
	const output = await execute.call(ctx, 0);
	return { output, requests };
}

describe('delete authentication method', () => {
	it('addresses the method through its own collection segment', async () => {
		const { requests, output } = await run({});

		expect(requests[0]).toMatchObject({
			method: 'DELETE',
			url: `https://graph.microsoft.com/v1.0/users/${USER}/authentication/fido2Methods/${METHOD}`,
		});
		expect(output[0].json).toEqual({ deleted: true });
	});

	it('refuses a method type Graph has no delete for', async () => {
		await expect(run({ methodType: 'passwordMethods' })).rejects.toThrow(
			/not a deletable method type/,
		);
	});

	it('refuses a method type an expression invented', async () => {
		await expect(run({ methodType: 'made/up?x=' })).rejects.toThrow(/not a deletable method type/);
	});

	it.each([
		['user', { user: 'alice/x?y=' }],
		['method', { method: '../../bob' }],
	])('refuses an unsafe %s before any request goes out', async (_label, parameters) => {
		const { ctx, httpRequestWithAuthentication } = operationContext({
			parameters: { user: USER, method: METHOD, methodType: 'fido2Methods', ...parameters },
		});

		await expect(execute.call(ctx, 0)).rejects.toThrow(/characters that are not allowed/);
		expect(httpRequestWithAuthentication).not.toHaveBeenCalled();
	});
});
