import type {
	IDataObject,
	IExecuteFunctions,
	ILoadOptionsFunctions,
	IN8nHttpFullResponse,
	IExecuteSingleFunctions,
} from 'n8n-workflow';
import { NodeApiError } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import {
	getGroupProperties,
	getGroups,
	getSubscribedSkus,
	getUserProperties,
	getUsers,
	microsoftApiPaginateRequest,
	microsoftApiRequest,
} from './GenericFunctions';
import { handleErrorPostReceive } from './GraphErrors';
import { deepMerge } from './utils';

const CREDENTIAL = 'microsoft365AdminServicePrincipalApi';

const METADATA = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
<edmx:DataServices>
<Schema Namespace="microsoft.graph">
<EntityType Name="entity" Abstract="true">
<Key><PropertyRef Name="id"/></Key>
<Property Name="id" Type="Edm.String" Nullable="false"/>
</EntityType>
<EntityType Name="directoryObject" BaseType="graph.entity">
<Property Name="deletedDateTime" Type="Edm.DateTimeOffset"/>
</EntityType>
<EntityType Name="user" BaseType="graph.directoryObject">
<Property Name="displayName" Type="Edm.String"/>
<Property Name="usageLocation" Type="Edm.String"/>
<Property Name="mailboxSettings" Type="graph.mailboxSettings"/>
<NavigationProperty Name="manager" Type="graph.directoryObject"/>
</EntityType>
<EntityType Name="group" BaseType="graph.directoryObject">
<Property Name="mailNickname" Type="Edm.String"/>
<Property Name="isArchived" Type="Edm.Boolean"/>
</EntityType>
<EntityType Name="device" BaseType="graph.directoryObject">
<Property Name="deviceId" Type="Edm.String"/>
</EntityType>
</Schema>
<Schema Namespace="microsoft.graph.callRecords">
<EntityType Name="session"><Property Name="fromOtherNamespace" Type="Edm.String"/></EntityType>
</Schema>
</edmx:DataServices>
</edmx:Edmx>`;

function mockContext(overrides: Record<string, unknown> = {}) {
	const httpRequestWithAuthentication = vi.fn(async () => ({}));
	const requestWithAuthenticationPaginated = vi.fn(async () => []);

	return {
		getCredentials: vi.fn(async () => ({ graphApiBaseUrl: 'https://graph.microsoft.com' })),
		getNode: vi.fn(() => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' })),
		getNodeParameter: vi.fn(),
		helpers: { httpRequestWithAuthentication, requestWithAuthenticationPaginated },
		...overrides,
	} as unknown as IExecuteFunctions & ILoadOptionsFunctions & IExecuteSingleFunctions;
}

describe('deepMerge', () => {
	it('copies new keys onto the target', () => {
		expect(deepMerge({ a: 1 }, { b: 2 })).toEqual({ a: 1, b: 2 });
	});

	it('overwrites scalars', () => {
		expect(deepMerge({ a: 1 }, { a: 2 })).toEqual({ a: 2 });
	});

	it('merges nested objects rather than replacing them', () => {
		const result = deepMerge(
			{ nested: { keep: 1, replace: 1 } },
			{ nested: { replace: 2, add: 3 } },
		);
		expect(result).toEqual({ nested: { keep: 1, replace: 2, add: 3 } });
	});

	it('replaces arrays wholesale instead of merging by index', () => {
		expect(deepMerge({ list: [1, 2, 3] }, { list: [9] })).toEqual({ list: [9] });
	});

	it('replaces an object with a scalar when types differ', () => {
		expect(deepMerge({ a: { b: 1 } }, { a: 'scalar' })).toEqual({ a: 'scalar' });
	});

	it('treats null as a scalar rather than recursing into it', () => {
		expect(deepMerge({ a: { b: 1 } }, { a: null })).toEqual({ a: null });
	});

	it('mutates and returns the target, matching the previous lodash behaviour', () => {
		const target: IDataObject = { a: 1 };
		expect(deepMerge(target, { b: 2 })).toBe(target);
		expect(target).toEqual({ a: 1, b: 2 });
	});
});

describe('microsoftApiRequest', () => {
	it('builds a v1.0 URL against the credential base and authenticates', async () => {
		const context = mockContext();
		await microsoftApiRequest.call(context, 'GET', '/users');

		expect(context.helpers.httpRequestWithAuthentication).toHaveBeenCalledWith(
			CREDENTIAL,
			expect.objectContaining({
				method: 'GET',
				url: 'https://graph.microsoft.com/v1.0/users',
				json: true,
			}),
		);
	});

	it('strips a trailing slash from a stored base URL', async () => {
		const context = mockContext({
			getCredentials: vi.fn(async () => ({ graphApiBaseUrl: 'https://graph.microsoft.us///' })),
		});
		await microsoftApiRequest.call(context, 'GET', '/users');

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.url).toBe('https://graph.microsoft.us/v1.0/users');
	});

	it('falls back to the global cloud when the credential has no base URL', async () => {
		const context = mockContext({ getCredentials: vi.fn(async () => ({})) });
		await microsoftApiRequest.call(context, 'GET', '/users');

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.url).toBe('https://graph.microsoft.com/v1.0/users');
	});

	it('honours an explicit absolute URL over the constructed one', async () => {
		const context = mockContext();
		const absolute = 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc';
		await microsoftApiRequest.call(context, 'GET', '/users', {}, { url: absolute });

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.url).toBe(absolute);
	});

	it('passes body, query string and headers through', async () => {
		const context = mockContext();
		await microsoftApiRequest.call(
			context,
			'POST',
			'/users',
			{ displayName: 'Ada' },
			{ qs: { $top: 10 }, headers: { ConsistencyLevel: 'eventual' } },
		);

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.body).toEqual({ displayName: 'Ada' });
		expect(options.qs).toEqual({ $top: 10 });
		expect(options.headers).toEqual({ ConsistencyLevel: 'eventual' });
	});
});

describe('microsoftApiPaginateRequest', () => {
	it('concatenates the value arrays of every page', async () => {
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(),
				requestWithAuthenticationPaginated: vi.fn(async () => [
					{ body: { value: [{ id: '1' }, { id: '2' }] } },
					{ body: { value: [{ id: '3' }] } },
				]),
			},
		});

		const result = await microsoftApiPaginateRequest.call(context, 'GET', '/users');
		expect(result).toEqual([{ id: '1' }, { id: '2' }, { id: '3' }]);
	});

	it('skips pages that carry no value array', async () => {
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(),
				requestWithAuthenticationPaginated: vi.fn(async () => [
					{ body: {} },
					{ body: { value: [{ id: '1' }] } },
				]),
			},
		});

		expect(await microsoftApiPaginateRequest.call(context, 'GET', '/users')).toEqual([{ id: '1' }]);
	});

	it('follows @odata.nextLink through the pagination options', async () => {
		const paginated = vi.fn(async () => []);
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(),
				requestWithAuthenticationPaginated: paginated,
			},
		});

		await microsoftApiPaginateRequest.call(context, 'GET', '/users');

		const [, , paginationOptions, credentialType] = paginated.mock.calls[0] as unknown[];
		expect(credentialType).toBe(CREDENTIAL);
		expect(JSON.stringify(paginationOptions)).toContain('@odata.nextLink');
	});
});

describe('getSubscribedSkus', () => {
	it('maps SKUs to options labelled with seat usage', async () => {
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(async () => ({
					value: [
						{
							skuId: 'sku-1',
							skuPartNumber: 'ENTERPRISEPACK',
							consumedUnits: 142,
							prepaidUnits: { enabled: 200 },
						},
					],
				})),
				requestWithAuthenticationPaginated: vi.fn(),
			},
		});

		expect(await getSubscribedSkus.call(context)).toEqual([
			{ name: 'ENTERPRISEPACK (142/200 used)', value: 'sku-1' },
		]);
	});

	it('defaults missing seat counts to zero rather than rendering undefined', async () => {
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(async () => ({
					value: [{ skuId: 'sku-2', skuPartNumber: 'FLOW_FREE' }],
				})),
				requestWithAuthenticationPaginated: vi.fn(),
			},
		});

		expect(await getSubscribedSkus.call(context)).toEqual([
			{ name: 'FLOW_FREE (0/0 used)', value: 'sku-2' },
		]);
	});

	it('returns an empty list when the tenant has no subscriptions', async () => {
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(async () => ({})),
				requestWithAuthenticationPaginated: vi.fn(),
			},
		});

		expect(await getSubscribedSkus.call(context)).toEqual([]);
	});
});

describe('metadata-driven property loaders', () => {
	function metadataContext() {
		return mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(async () => METADATA),
				requestWithAuthenticationPaginated: vi.fn(),
			},
		});
	}

	it('collects user properties from the inherited entity chain', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);

		expect(names).toContain('displayName');
		expect(names).toContain('usageLocation');
		expect(names).toContain('deletedDateTime');
	});

	it('excludes properties that need permissions beyond the node’s scope', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);

		expect(names).not.toContain('mailboxSettings');
		expect(names).not.toContain('id');
	});

	it('ignores NavigationProperty elements', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('manager');
	});

	it('ignores entity types that were not requested', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('deviceId');
	});

	it('ignores schemas outside the microsoft.graph namespace', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('fromOtherNamespace');
	});

	it('returns options sorted alphabetically', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).toEqual([...names].sort());
	});

	it('collects group properties and drops the excluded ones', async () => {
		const names = (await getGroupProperties.call(metadataContext())).map((o) => o.value);

		expect(names).toContain('mailNickname');
		expect(names).not.toContain('isArchived');
		expect(names).not.toContain('id');
		expect(names).not.toContain('displayName');
	});
});

describe('resource locator searches', () => {
	function listContext(value: unknown[]) {
		return mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(async () => ({ value })),
				requestWithAuthenticationPaginated: vi.fn(),
			},
		});
	}

	it('maps groups to search results', async () => {
		const context = listContext([{ id: 'g1', displayName: 'Engineering' }]);
		const result = await getGroups.call(context);

		expect(result.results).toEqual([{ name: 'Engineering', value: 'g1' }]);
	});

	it('maps users to search results', async () => {
		const context = listContext([{ id: 'u1', displayName: 'Ada Lovelace' }]);
		const result = await getUsers.call(context);

		expect(result.results).toEqual([{ name: 'Ada Lovelace', value: 'u1' }]);
	});

	it('returns an empty result set when nothing matches', async () => {
		const result = await getUsers.call(listContext([]));
		expect(result.results).toEqual([]);
	});

	it('sorts results case- and numeric-insensitively', async () => {
		const context = listContext([
			{ id: '1', displayName: 'zeta' },
			{ id: '2', displayName: 'Alpha' },
			{ id: '3', displayName: 'item10' },
			{ id: '4', displayName: 'item2' },
		]);

		const names = (await getUsers.call(context)).results.map((r) => r.name);
		expect(names).toEqual(['Alpha', 'item2', 'item10', 'zeta']);
	});

	it('surfaces the nextLink as a pagination token', async () => {
		const context = mockContext({
			helpers: {
				httpRequestWithAuthentication: vi.fn(async () => ({
					value: [],
					'@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc',
				})),
				requestWithAuthenticationPaginated: vi.fn(),
			},
		});

		const result = await getUsers.call(context);
		expect(result.paginationToken).toContain('$skiptoken=abc');
	});

	it('searches groups by display name with the eventual-consistency header', async () => {
		const context = listContext([]);
		await getGroups.call(context, 'Engineering');

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.qs?.$search).toBe('"displayName:Engineering"');
		expect(options.headers?.ConsistencyLevel).toBe('eventual');
	});

	it('filters users across display name and UPN', async () => {
		const context = listContext([]);
		await getUsers.call(context, 'ada');

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.qs?.$filter).toContain("startsWith(displayName, 'ada')");
		expect(options.qs?.$filter).toContain("startsWith(userPrincipalName, 'ada')");
	});

	it('doubles an apostrophe so a name like O’Brien does not break the filter', async () => {
		const context = listContext([]);
		await getUsers.call(context, "o'brien");

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.qs?.$filter).toBe(
			"startsWith(displayName, 'o''brien') OR startsWith(userPrincipalName, 'o''brien')",
		);
	});

	it('escapes a quote in a group search phrase', async () => {
		const context = listContext([]);
		await getGroups.call(context, 'say "hi"');

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.qs?.$search).toBe('"displayName:say \\"hi\\""');
	});

	it.each([
		['groups', getGroups],
		['users', getUsers],
	])('falls back to the ID when %s have no display name', async (label, search) => {
		const context = listContext([{ id: 'no-name' }] as Array<{ id: string; displayName: string }>);

		const { results } = await search.call(context);

		expect(results, label).toEqual([{ name: 'no-name', value: 'no-name' }]);
	});

	it.each([
		['groups', getGroups],
		['users', getUsers],
	])('follows the pagination token for %s instead of re-filtering', async (_label, search) => {
		const context = listContext([]);
		const token = 'https://graph.microsoft.com/v1.0/x?$skiptoken=next';

		await search.call(context, undefined, token);

		const [, options] = vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0];
		expect(options.url).toBe(token);
		expect(options.qs).toBeUndefined();
	});
});

describe('handleErrorPostReceive', () => {
	function response(statusCode: number, error: unknown): IN8nHttpFullResponse {
		return { statusCode, body: { error }, headers: {} } as unknown as IN8nHttpFullResponse;
	}

	function errorContext(resource: string, operation: string, params: Record<string, string> = {}) {
		return mockContext({
			getNodeParameter: vi.fn((name: string) => {
				if (name === 'resource') return resource;
				if (name === 'operation') return operation;
				return params[name];
			}),
		});
	}

	const items = [{ json: {} }];

	it('passes successful responses through untouched', async () => {
		const result = await handleErrorPostReceive.call(errorContext('user', 'get'), items, {
			statusCode: 200,
			body: {},
			headers: {},
		} as unknown as IN8nHttpFullResponse);
		expect(result).toBe(items);
	});

	it('raises a friendly message when a user lookup misses', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'get'),
				items,
				response(404, { code: 'Request_ResourceNotFound', message: 'not found' }),
			),
		).rejects.toThrow(NodeApiError);
	});

	it('raises a friendly message when a group lookup misses', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('group', 'delete'),
				items,
				response(404, { code: 'Request_ResourceNotFound', message: 'not found' }),
			),
		).rejects.toThrow(NodeApiError);
	});

	it('swallows the empty-payload error n8n produces for a no-op group update', async () => {
		const result = await handleErrorPostReceive.call(
			errorContext('group', 'update'),
			items,
			response(400, { code: 'BadRequest', message: 'Empty Payload. JSON content expected.' }),
		);
		expect(result).toBe(items);
	});

	it('explains that a user is already a member when adding a duplicate', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'addGroup'),
				items,
				response(400, {
					code: 'Request_BadRequest',
					message:
						"One or more added object references already exist for the following modified properties: 'members'.",
				}),
			),
		).rejects.toThrow(/already in the group/);
	});

	it('explains that a user is not a member when removing a non-member', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'removeGroup'),
				items,
				response(404, { code: 'Request_ResourceNotFound', message: 'not found' }),
			),
		).rejects.toThrow(/not in the group/);
	});

	it('still raises for server errors it has no specific message for', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'get'),
				items,
				response(500, { code: 'InternalServerError', message: 'boom' }),
			),
		).rejects.toThrow(NodeApiError);
	});

	const NOT_FOUND = { code: 'Request_ResourceNotFound', message: 'not found' };

	it.each([
		['group', 'delete', NOT_FOUND, /group doesn't match/],
		['group', 'get', NOT_FOUND, /group doesn't match/],
		['group', 'update', NOT_FOUND, /group doesn't match/],
		['user', 'delete', NOT_FOUND, /user doesn't match/],
		['user', 'get', NOT_FOUND, /user doesn't match/],
		['user', 'update', NOT_FOUND, /user doesn't match/],
	])('names the right parameter for %s.%s not-found', async (resource, operation, error, match) => {
		await expect(
			handleErrorPostReceive.call(errorContext(resource, operation), items, response(404, error)),
		).rejects.toThrow(match);
	});

	it('swallows the empty-payload error for a no-op user update too', async () => {
		const result = await handleErrorPostReceive.call(
			errorContext('user', 'update'),
			items,
			response(400, { code: 'BadRequest', message: 'Empty Payload. JSON content expected.' }),
		);
		expect(result).toBe(items);
	});

	it('blames the group when an addGroup miss mentions the group ID', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'addGroup', { 'group.value': 'group-1' }),
				items,
				response(404, { code: 'Request_ResourceNotFound', message: 'no group-1 here' }),
			),
		).rejects.toThrow(/group doesn't match/);
	});

	it('blames the user when an addGroup miss does not mention the group ID', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'addGroup', { 'group.value': 'group-1' }),
				items,
				response(404, { code: 'Request_ResourceNotFound', message: 'missing directory object' }),
			),
		).rejects.toThrow(/user doesn't match/);
	});

	it('explains an unusable member reference on removeGroup', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('user', 'removeGroup'),
				items,
				response(400, {
					code: 'Request_UnsupportedQuery',
					message: "Unsupported referenced-object resource identifier for link property 'members'.",
				}),
			),
		).rejects.toThrow(/user ID is invalid/);
	});

	it.each([
		['group', 'group ID is invalid'],
		['user', 'user ID is invalid'],
	])('reports an invalid object identifier for %s', async (resource, message) => {
		await expect(
			handleErrorPostReceive.call(
				errorContext(resource, 'get', { 'group.value': 'group-1' }),
				items,
				response(400, {
					code: 'Request_BadRequest',
					message: 'Invalid object identifier "abc"',
				}),
			),
		).rejects.toThrow(new RegExp(message));
	});

	it.each(['ObjectConflict', 'ConflictingObjects'])(
		'reports an already-existing resource for %s',
		async (code) => {
			await expect(
				handleErrorPostReceive.call(
					errorContext('group', 'create'),
					items,
					response(400, {
						code: 'Request_BadRequest',
						message: 'conflict',
						details: [{ code, message: 'conflict' }],
					}),
				),
			).rejects.toThrow(/group already exists/);
		},
	);

	it('falls through to a bare API error for an unrecognised failure', async () => {
		await expect(
			handleErrorPostReceive.call(
				errorContext('license', 'assign'),
				items,
				response(403, { code: 'Authorization_RequestDenied', message: 'denied' }),
			),
		).rejects.toThrow(NodeApiError);
	});
});
