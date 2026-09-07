import {
	NodeApiError,
	type IDataObject,
	type IExecuteFunctions,
	type IHttpRequestOptions,
} from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { getGraphApiBaseUrl, microsoftApiPaginateRequest, microsoftApiRequest } from './index';

const CREDENTIAL = 'microsoft365AdminServicePrincipalApi';

interface Stub {
	/** What the transport's own request answers with. */
	response?: { statusCode?: number; body?: unknown; headers?: IDataObject };
	credentials?: IDataObject;
	parameters?: Record<string, unknown>;
	/** Consumed one per request, for the multi-request pagination walk. */
	responses?: Array<{ statusCode?: number; body?: unknown; headers?: IDataObject }>;
	/** Thrown by the auth helper instead of answering, to exercise the failure path. */
	throws?: unknown;
}

function mockContext(stub: Stub = {}) {
	let call = 0;
	const httpRequestWithAuthentication = vi.fn(
		async (_credentialType: string, _options: IHttpRequestOptions) => {
			if (stub.throws !== undefined) {
				throw stub.throws;
			}
			const response = stub.responses ? stub.responses[call++] : stub.response;
			return { statusCode: 200, headers: {}, body: {}, ...response };
		},
	);

	return {
		getCredentials: vi.fn(
			async () => stub.credentials ?? { graphApiBaseUrl: 'https://graph.microsoft.com' },
		),
		getNode: vi.fn(() => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' })),
		getNodeParameter: vi.fn(
			(name: string, _index?: number, fallback?: unknown) =>
				stub.parameters?.[name] ?? fallback ?? '',
		),
		helpers: { httpRequestWithAuthentication },
	} as unknown as IExecuteFunctions;
}

function sentOptions(context: IExecuteFunctions) {
	return vi.mocked(context.helpers.httpRequestWithAuthentication).mock.calls[0][1];
}

describe('getGraphApiBaseUrl', () => {
	it('strips a stored trailing slash', async () => {
		const context = mockContext({
			credentials: { graphApiBaseUrl: 'https://graph.microsoft.us///' },
		});
		expect(await getGraphApiBaseUrl.call(context)).toBe('https://graph.microsoft.us');
	});

	it('falls back to the global cloud when the credential has none', async () => {
		const context = mockContext({ credentials: {} });
		expect(await getGraphApiBaseUrl.call(context)).toBe('https://graph.microsoft.com');
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

	it('honours an explicit absolute URL over the constructed one', async () => {
		const context = mockContext();
		const absolute = 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc';
		await microsoftApiRequest.call(context, 'GET', '/users', {}, { url: absolute });

		expect(sentOptions(context).url).toBe(absolute);
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

		const options = sentOptions(context);
		expect(options.body).toEqual({ displayName: 'Ada' });
		expect(options.qs).toEqual({ $top: 10 });
		expect(options.headers).toEqual({ ConsistencyLevel: 'eventual' });
	});

	it('turns JSON parsing off for the $metadata documents', async () => {
		const context = mockContext();
		await microsoftApiRequest.call(context, 'GET', '/$metadata#users');
		expect(sentOptions(context).json).toBe(false);
	});

	it('always asks for the full response so it can read the status itself', async () => {
		const context = mockContext();
		await microsoftApiRequest.call(context, 'GET', '/users');
		expect(sentOptions(context).returnFullResponse).toBe(true);
	});

	it('lets a 401 reach the auth helper so an expired token can refresh', async () => {
		const context = mockContext();
		await microsoftApiRequest.call(context, 'GET', '/users');
		expect(sentOptions(context).ignoreHttpStatusErrors).toEqual({ ignore: true, except: [401] });
	});

	it('unwraps the body by default', async () => {
		const context = mockContext({ response: { body: { id: 'u1' } } });
		expect(await microsoftApiRequest.call(context, 'GET', '/users/u1')).toEqual({ id: 'u1' });
	});

	it('returns the whole response when asked for it', async () => {
		const context = mockContext({ response: { body: { id: 'u1' } } });
		const result = (await microsoftApiRequest.call(
			context,
			'GET',
			'/users/u1',
			{},
			{ returnFullResponse: true },
		)) as { statusCode: number };
		expect(result.statusCode).toBe(200);
	});

	it('translates a Graph failure into a node error', async () => {
		const context = mockContext({
			response: {
				statusCode: 404,
				body: { error: { code: 'Request_ResourceNotFound', message: 'not found' } },
			},
			parameters: { resource: 'user', operation: 'get' },
		});

		await expect(microsoftApiRequest.call(context, 'GET', '/users/u1')).rejects.toThrow(
			NodeApiError,
		);
	});

	it('hands a failure back untouched when the caller says it will inspect it', async () => {
		const context = mockContext({
			response: { statusCode: 400, body: { error: { code: 'X', message: 'm' } } },
		});

		const result = (await microsoftApiRequest.call(
			context,
			'POST',
			'/users/u1/assignLicense',
			{},
			{ returnFullResponse: true, ignoreHttpStatusErrors: { ignore: true, except: [401] } },
		)) as { statusCode: number };

		expect(result.statusCode).toBe(400);
	});
});

describe('microsoftApiPaginateRequest', () => {
	const page = (value: IDataObject[], nextLink?: string) => ({
		body: { value, ...(nextLink === undefined ? {} : { '@odata.nextLink': nextLink }) },
	});

	it('concatenates the value arrays of every page', async () => {
		const context = mockContext({
			responses: [
				page([{ id: '1' }, { id: '2' }], 'https://graph.microsoft.com/v1.0/users?$skiptoken=a'),
				page([{ id: '3' }]),
			],
		});

		expect(await microsoftApiPaginateRequest.call(context, 'GET', '/users')).toEqual([
			{ id: '1' },
			{ id: '2' },
			{ id: '3' },
		]);
	});

	it('skips pages that carry no value array', async () => {
		const context = mockContext({
			responses: [
				{ body: { '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=a' } },
				page([{ id: '1' }]),
			],
		});

		expect(await microsoftApiPaginateRequest.call(context, 'GET', '/users')).toEqual([{ id: '1' }]);
	});

	it('requests each page through the authenticated helper, so an expired token refreshes', async () => {
		const nextLink = 'https://graph.microsoft.com/v1.0/users?$skiptoken=a';
		const context = mockContext({ responses: [page([{ id: '1' }], nextLink), page([{ id: '2' }])] });

		await microsoftApiPaginateRequest.call(context, 'GET', '/users', {}, { qs: { $top: 999 } });

		const { calls } = vi.mocked(context.helpers.httpRequestWithAuthentication).mock;
		expect(calls).toHaveLength(2);
		expect(calls[0][0]).toBe(CREDENTIAL);
		expect(calls[0][1].url).toBe('https://graph.microsoft.com/v1.0/users');
		expect(calls[0][1].qs).toEqual({ $top: 999 });
		// The nextLink already carries the query string; re-applying `qs` would duplicate it.
		expect(calls[1][1].url).toBe(nextLink);
		expect(calls[1][1].qs).toBeUndefined();
	});

	it('surfaces a failing page instead of dropping it as an empty result', async () => {
		const context = mockContext({
			responses: [
				{ statusCode: 403, body: { error: { code: 'Authorization_RequestDenied', message: 'no' } } },
			],
		});

		await expect(microsoftApiPaginateRequest.call(context, 'GET', '/users')).rejects.toBeInstanceOf(
			NodeApiError,
		);
	});

	it('stops when a nextLink repeats rather than looping forever', async () => {
		const nextLink = 'https://graph.microsoft.com/v1.0/users?$skiptoken=a';
		const context = mockContext({
			responses: [page([{ id: '1' }], nextLink), page([{ id: '2' }], nextLink)],
		});

		expect(await microsoftApiPaginateRequest.call(context, 'GET', '/users')).toEqual([
			{ id: '1' },
			{ id: '2' },
		]);
		expect(context.helpers.httpRequestWithAuthentication).toHaveBeenCalledTimes(2);
	});

	it('passes the item index through so errors point at the right item', async () => {
		const context = mockContext({
			responses: [
				{ statusCode: 404, body: { error: { code: 'Request_ResourceNotFound', message: 'no' } } },
			],
		});

		await expect(
			microsoftApiPaginateRequest.call(context, 'GET', '/users', {}, { itemIndex: 4 }),
		).rejects.toMatchObject({ context: { itemIndex: 4 } });
	});
});

describe('a 401 that survives the token refresh', () => {
	it('is reported as a rejected access token rather than a raw Graph envelope', async () => {
		const context = mockContext({
			throws: Object.assign(new Error('401 - {"error":{"code":"InvalidAuthenticationToken"}}'), {
				httpCode: '401',
			}),
		});

		await expect(microsoftApiRequest.call(context, 'GET', '/users')).rejects.toMatchObject({
			message: 'Microsoft Graph rejected the access token',
		});
	});

	it('leaves any other transport failure untouched', async () => {
		const failure = Object.assign(new Error('socket hang up'), { code: 'ECONNRESET' });
		const context = mockContext({ throws: failure });

		await expect(microsoftApiRequest.call(context, 'GET', '/users')).rejects.toBe(failure);
	});
});
