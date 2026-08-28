import { describe, expect, it } from 'vitest';

import { operationContext } from '../operation.fixtures';
import { USER_RAW_SELECT, USER_SIMPLE_SELECT } from './constants';
import { execute as getAll } from './getAll.operation';
import { execute as get } from './get.operation';

const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';

describe('user get', () => {
	it('projects the simplified field set by default', async () => {
		const { ctx, requests } = operationContext({ parameters: { user: USER, output: 'simple' } });
		await get.call(ctx, 0);

		expect(requests[0].url).toBe(`https://graph.microsoft.com/v1.0/users/${USER}`);
		expect(requests[0].qs?.$select).toBe(USER_SIMPLE_SELECT);
	});

	it('names every readable property for Raw output', async () => {
		const { ctx, requests } = operationContext({ parameters: { user: USER, output: 'raw' } });
		await get.call(ctx, 0);

		expect(requests[0].qs?.$select).toBe(USER_RAW_SELECT);
	});

	it('always projects id alongside the chosen fields', async () => {
		const { ctx, requests } = operationContext({
			parameters: { user: USER, output: 'fields', fields: ['displayName', 'jobTitle'] },
		});
		await get.call(ctx, 0);

		expect(requests[0].qs?.$select).toBe('displayName,jobTitle,id');
	});

	it('does not repeat id when it was chosen explicitly', async () => {
		const { ctx, requests } = operationContext({
			parameters: { user: USER, output: 'fields', fields: ['id', 'displayName'] },
		});
		await get.call(ctx, 0);

		expect(requests[0].qs?.$select).toBe('id,displayName');
	});
});

describe('user getAll', () => {
	it('asks for one page of `limit` when Return All is off', async () => {
		const { ctx, requests } = operationContext({
			parameters: { output: 'simple', returnAll: false, limit: 25 },
		});
		await getAll.call(ctx, 0);

		expect(requests[0].url).toBe('https://graph.microsoft.com/v1.0/users');
		expect(requests[0].qs?.$top).toBe(25);
	});

	it('walks every page when Return All is on', async () => {
		const { ctx, requests } = operationContext({
			parameters: { output: 'simple', returnAll: true },
			pages: [{ body: { value: [{ id: 'u1' }] } }, { body: { value: [{ id: 'u2' }] } }],
		});

		const output = await getAll.call(ctx, 0);

		// Paging goes through the paginated helper, not the single-request transport.
		expect(requests).toHaveLength(0);
		expect(output.map((item) => item.json.id)).toEqual(['u1', 'u2']);
	});

	it('passes a filter through untouched', async () => {
		const { ctx, requests } = operationContext({
			parameters: {
				output: 'raw',
				returnAll: false,
				limit: 50,
				filter: "startswith(displayName,'a')",
			},
		});
		await getAll.call(ctx, 0);

		expect(requests[0].qs?.$filter).toBe("startswith(displayName,'a')");
	});

	it('sends no filter when the parameter is empty', async () => {
		const { ctx, requests } = operationContext({
			parameters: { output: 'raw', returnAll: false, limit: 50, filter: '' },
		});
		await getAll.call(ctx, 0);

		expect(requests[0].qs?.$filter).toBeUndefined();
	});

	it('unwraps the collection into one item per user', async () => {
		const { ctx } = operationContext({
			parameters: { output: 'simple', returnAll: false, limit: 50 },
			responses: [{ body: { value: [{ id: 'u1' }, { id: 'u2' }] } }],
		});

		const output = await getAll.call(ctx, 0);
		expect(output.map((item) => item.json.id)).toEqual(['u1', 'u2']);
		expect(output.every((item) => item.pairedItem)).toBe(true);
	});
});
