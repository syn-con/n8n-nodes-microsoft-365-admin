import type { IDataObject, IHttpRequestOptions, ILoadOptionsFunctions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { getAuthenticationMethods, getGroups, getUsers } from './listSearch';

const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';

interface Stub {
	/** The Graph body the single request answers with. */
	body?: IDataObject;
	/** Values `getCurrentNodeParameter` hands back, for the method picker. */
	current?: Record<string, unknown>;
}

function mockContext(stub: Stub = {}) {
	const requests: IHttpRequestOptions[] = [];

	const httpRequestWithAuthentication = vi.fn(
		async (_credentialType: string, options: IHttpRequestOptions) => {
			requests.push(options);
			return { statusCode: 200, headers: {}, body: stub.body ?? {} };
		},
	);

	const ctx = {
		getCredentials: vi.fn(async () => ({ graphApiBaseUrl: 'https://graph.microsoft.com' })),
		getNode: vi.fn(() => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' })),
		getNodeParameter: vi.fn(() => ''),
		getCurrentNodeParameter: vi.fn((name: string) => stub.current?.[name]),
		helpers: { httpRequestWithAuthentication, requestWithAuthenticationPaginated: vi.fn() },
	} as unknown as ILoadOptionsFunctions;

	return { ctx, requests };
}

/** A directory collection, the shape both pickers read. */
const listContext = (value: unknown[], extra: IDataObject = {}) =>
	mockContext({ body: { value, ...extra } as IDataObject });

describe('directory searches', () => {
	it('maps groups to search results', async () => {
		const { ctx } = listContext([{ id: 'g1', displayName: 'Engineering' }]);
		expect((await getGroups.call(ctx)).results).toEqual([{ name: 'Engineering', value: 'g1' }]);
	});

	it('maps users to search results', async () => {
		const { ctx } = listContext([{ id: 'u1', displayName: 'Ada Lovelace' }]);
		expect((await getUsers.call(ctx)).results).toEqual([{ name: 'Ada Lovelace', value: 'u1' }]);
	});

	it('returns an empty result set when nothing matches', async () => {
		const { ctx } = listContext([]);
		expect((await getUsers.call(ctx)).results).toEqual([]);
	});

	it('sorts results case- and numeric-insensitively', async () => {
		const { ctx } = listContext([
			{ id: '1', displayName: 'zeta' },
			{ id: '2', displayName: 'Alpha' },
			{ id: '3', displayName: 'item10' },
			{ id: '4', displayName: 'item2' },
		]);

		expect((await getUsers.call(ctx)).results.map((r) => r.name)).toEqual([
			'Alpha',
			'item2',
			'item10',
			'zeta',
		]);
	});

	it('surfaces the nextLink as a pagination token', async () => {
		const { ctx } = listContext([], {
			'@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
		});
		expect((await getUsers.call(ctx)).paginationToken).toContain('$skiptoken=abc');
	});

	it('searches groups by display name with the eventual-consistency header', async () => {
		const { ctx, requests } = listContext([]);
		await getGroups.call(ctx, 'Engineering');

		expect(requests[0].qs?.$search).toBe('"displayName:Engineering"');
		expect(requests[0].headers?.ConsistencyLevel).toBe('eventual');
	});

	it('filters users across display name and UPN', async () => {
		const { ctx, requests } = listContext([]);
		await getUsers.call(ctx, 'ada');

		expect(requests[0].qs?.$filter).toContain("startsWith(displayName, 'ada')");
		expect(requests[0].qs?.$filter).toContain("startsWith(userPrincipalName, 'ada')");
	});

	it('doubles an apostrophe so a name like O’Brien does not break the filter', async () => {
		const { ctx, requests } = listContext([]);
		await getUsers.call(ctx, "o'brien");

		expect(requests[0].qs?.$filter).toBe(
			"startsWith(displayName, 'o''brien') OR startsWith(userPrincipalName, 'o''brien')",
		);
	});

	it('escapes a quote in a group search phrase', async () => {
		const { ctx, requests } = listContext([]);
		await getGroups.call(ctx, 'say "hi"');

		expect(requests[0].qs?.$search).toBe('"displayName:say \\"hi\\""');
	});

	it.each([
		['groups', getGroups],
		['users', getUsers],
	])('falls back to the ID when %s have no display name', async (label, search) => {
		const { ctx } = listContext([{ id: 'no-name' }]);
		expect((await search.call(ctx)).results, label).toEqual([
			{ name: 'no-name', value: 'no-name' },
		]);
	});

	it.each([
		['groups', getGroups],
		['users', getUsers],
	])('follows the pagination token for %s instead of re-filtering', async (_label, search) => {
		const { ctx, requests } = listContext([]);
		const token = 'https://graph.microsoft.com/v1.0/x?$skiptoken=next';

		await search.call(ctx, undefined, token);

		expect(requests[0].url).toBe(token);
		expect(requests[0].qs).toBeUndefined();
	});
});

describe('the method picker', () => {
	it('lists the methods of the chosen type for the chosen user', async () => {
		const { ctx, requests } = mockContext({
			current: { user: USER, methodType: 'phoneMethods' },
			body: {
				value: [
					{
						'@odata.type': '#microsoft.graph.phoneAuthenticationMethod',
						id: 'b',
						phoneNumber: '+370 600 00002',
						phoneType: 'alternateMobile',
					},
					{
						'@odata.type': '#microsoft.graph.phoneAuthenticationMethod',
						id: 'a',
						phoneNumber: '+370 600 00001',
						phoneType: 'mobile',
					},
				],
			},
		});

		const result = await getAuthenticationMethods.call(ctx);

		expect(requests[0].url).toBe(
			`https://graph.microsoft.com/v1.0/users/${USER}/authentication/methods`,
		);
		expect(result.results.map((r) => r.value)).toEqual(['a', 'b']);
	});

	it('filters the aggregate response to the selected method type', async () => {
		const { ctx } = mockContext({
			current: { user: USER, methodType: 'platformCredentialMethods' },
			body: {
				value: [{ '@odata.type': '#microsoft.graph.passwordAuthenticationMethod', id: 'password' }],
			},
		});

		expect((await getAuthenticationMethods.call(ctx)).results).toEqual([]);
	});

	it('filters on the label, since these collections take no $filter', async () => {
		const { ctx } = mockContext({
			current: { user: USER, methodType: 'fido2Methods' },
			body: {
				value: [
					{
						'@odata.type': '#microsoft.graph.fido2AuthenticationMethod',
						id: 'k1',
						displayName: 'Red Key',
					},
					{
						'@odata.type': '#microsoft.graph.fido2AuthenticationMethod',
						id: 'k2',
						displayName: 'Blue Key',
					},
				],
			},
		});

		expect((await getAuthenticationMethods.call(ctx, 'red')).results).toEqual([
			{ name: 'Red Key', value: 'k1' },
		]);
	});

	it('asks for a user before it can offer anything', async () => {
		const { ctx } = mockContext({ current: { user: '', methodType: 'phoneMethods' } });
		await expect(getAuthenticationMethods.call(ctx)).rejects.toThrow(/Choose a user first/);
	});

	it('refuses a user value that would reshape the request path', async () => {
		const { ctx } = mockContext({
			current: { user: 'alice/authentication/methods?x=', methodType: 'phoneMethods' },
		});
		await expect(getAuthenticationMethods.call(ctx)).rejects.toThrow(/Choose a user first/);
	});

	it('asks for a method type it recognizes', async () => {
		const { ctx } = mockContext({ current: { user: USER, methodType: 'madeUpMethods' } });
		await expect(getAuthenticationMethods.call(ctx)).rejects.toThrow(/Choose a method type first/);
	});

	it('copes with a user who has no methods of that type', async () => {
		const { ctx } = mockContext({ current: { user: USER, methodType: 'fido2Methods' }, body: {} });
		expect((await getAuthenticationMethods.call(ctx)).results).toEqual([]);
	});
});
