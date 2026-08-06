import type { ICredentialDataDecryptedObject, IHttpRequestOptions } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import {
	getAccessToken,
	Microsoft365AdminServicePrincipalApi,
} from './Microsoft365AdminServicePrincipalApi.credentials';

const TENANT = '35eb41ef-cdeb-42c3-ad87-f101a0bff350';
const CLIENT = 'a09519e2-abb5-4efa-a21f-41c86831d152';

function credentials(
	overrides: ICredentialDataDecryptedObject = {},
): ICredentialDataDecryptedObject {
	return {
		authentication: 'clientSecret',
		tenantId: TENANT,
		clientId: CLIENT,
		clientSecret: 'a-secret',
		graphApiBaseUrl: 'https://graph.microsoft.com',
		...overrides,
	};
}

/** Captures the request and returns a token, standing in for `helpers.httpRequest`. */
function tokenPoster(response: unknown = { access_token: 'token-abc' }) {
	const calls: IHttpRequestOptions[] = [];
	const post = vi.fn(async (options: IHttpRequestOptions) => {
		calls.push(options);
		return response;
	});
	return { post, calls };
}

function bodyOf(options: IHttpRequestOptions): URLSearchParams {
	return new URLSearchParams(options.body as string);
}

describe('getAccessToken', () => {
	let poster: ReturnType<typeof tokenPoster>;

	beforeEach(() => {
		poster = tokenPoster();
	});

	it('returns the access token from a successful exchange', async () => {
		await expect(getAccessToken(credentials(), poster.post)).resolves.toBe('token-abc');
	});

	it('posts the client_credentials grant to the tenant-specific token endpoint', async () => {
		await getAccessToken(credentials(), poster.post);

		const [request] = poster.calls;
		expect(request.method).toBe('POST');
		expect(request.url).toBe(
			`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
		);
		expect(request.headers?.['Content-Type']).toBe('application/x-www-form-urlencoded');

		const body = bodyOf(request);
		expect(body.get('grant_type')).toBe('client_credentials');
		expect(body.get('client_id')).toBe(CLIENT);
		expect(body.get('client_secret')).toBe('a-secret');
		expect(body.get('scope')).toBe('https://graph.microsoft.com/.default');
	});

	it('trims whitespace from pasted credential values', async () => {
		await getAccessToken(
			credentials({ tenantId: `  ${TENANT}  `, clientId: ` ${CLIENT} ` }),
			poster.post,
		);

		expect(poster.calls[0].url).toContain(TENANT);
		expect(bodyOf(poster.calls[0]).get('client_id')).toBe(CLIENT);
	});

	it('accepts a verified-domain tenant as well as a GUID', async () => {
		await expect(
			getAccessToken(credentials({ tenantId: 'contoso.onmicrosoft.com' }), poster.post),
		).resolves.toBe('token-abc');
	});

	it.each([
		['a sovereign US Gov cloud', 'https://graph.microsoft.us', 'https://login.microsoftonline.us'],
		['the DOD cloud', 'https://dod-graph.microsoft.us', 'https://login.microsoftonline.us'],
		[
			'the China cloud',
			'https://microsoftgraph.chinacloudapi.cn',
			'https://login.partner.microsoftonline.cn',
		],
	])('authenticates against the matching login host for %s', async (_label, graph, login) => {
		await getAccessToken(credentials({ graphApiBaseUrl: graph }), poster.post);

		expect(poster.calls[0].url).toBe(`${login}/${TENANT}/oauth2/v2.0/token`);
		expect(bodyOf(poster.calls[0]).get('scope')).toBe(`${graph}/.default`);
	});

	it('normalises a trailing slash on the base URL before resolving the cloud', async () => {
		await getAccessToken(credentials({ graphApiBaseUrl: 'https://graph.microsoft.us/' }), poster.post);

		expect(poster.calls[0].url).toContain('login.microsoftonline.us');
	});

	it('falls back to the global cloud when no base URL is stored', async () => {
		await getAccessToken(credentials({ graphApiBaseUrl: '' }), poster.post);

		expect(poster.calls[0].url).toContain('login.microsoftonline.com');
		expect(bodyOf(poster.calls[0]).get('scope')).toBe('https://graph.microsoft.com/.default');
	});

	it('sends a signed client assertion instead of a secret in certificate mode', async () => {
		const { privateKey, certificate } = await import('./common/client-assertion.fixtures');

		await getAccessToken(
			credentials({ authentication: 'certificate', clientSecret: '', privateKey, certificate }),
			poster.post,
		);

		const body = bodyOf(poster.calls[0]);
		expect(body.get('client_secret')).toBeNull();
		expect(body.get('client_assertion_type')).toBe(
			'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
		);
		expect(body.get('client_assertion')?.split('.')).toHaveLength(3);
	});

	describe('validation, before any network call', () => {
		it.each([
			['tenantId', { tenantId: '' }],
			['clientId', { clientId: '' }],
			['clientSecret', { clientSecret: '' }],
		])('rejects when %s is missing', async (_field, overrides) => {
			await expect(getAccessToken(credentials(overrides), poster.post)).rejects.toThrow(
				/credentials are incomplete/,
			);
			expect(poster.post).not.toHaveBeenCalled();
		});

		it('requires both key and certificate in certificate mode', async () => {
			await expect(
				getAccessToken(
					credentials({ authentication: 'certificate', clientSecret: '', privateKey: 'k' }),
					poster.post,
				),
			).rejects.toThrow(/credentials are incomplete/);
			expect(poster.post).not.toHaveBeenCalled();
		});

		it.each([
			['a path traversal attempt', '../../evil'],
			['an embedded slash', 'tenant/evil'],
			['an at sign', 'tenant@evil'],
			['whitespace inside', 'ten ant'],
			['a query separator', 'tenant?x=1'],
		])('rejects %s in the tenant ID without echoing it', async (_label, tenantId) => {
			await expect(getAccessToken(credentials({ tenantId }), poster.post)).rejects.toThrow(
				/not a valid GUID or domain/,
			);
			expect(poster.post).not.toHaveBeenCalled();
		});

		it('rejects a Graph base URL outside the known Microsoft clouds', async () => {
			await expect(
				getAccessToken(credentials({ graphApiBaseUrl: 'https://evil.example.com' }), poster.post),
			).rejects.toThrow(/not a recognized Microsoft cloud/);
			expect(poster.post).not.toHaveBeenCalled();
		});
	});

	describe('response handling', () => {
		it.each([
			['an empty object', {}],
			['a null body', null],
			['a string body', 'nope'],
			['an empty access_token', { access_token: '' }],
			['a non-string access_token', { access_token: 42 }],
		])('throws a static message for %s', async (_label, response) => {
			const bad = tokenPoster(response);

			await expect(getAccessToken(credentials(), bad.post)).rejects.toThrow(
				/did not return an access token/,
			);
		});

		it('never interpolates the error body into the thrown message', async () => {
			const bad = tokenPoster({ error: 'AADSTS700016', correlation_id: 'leak-me' });

			await expect(getAccessToken(credentials(), bad.post)).rejects.toThrow(
				/^Microsoft Entra authentication did not return an access token$/,
			);
		});
	});
});

describe('Microsoft365AdminServicePrincipalApi', () => {
	const credential = new Microsoft365AdminServicePrincipalApi();

	it('is registered under the Synergy-specific type name', () => {
		expect(credential.name).toBe('microsoft365AdminServicePrincipalApi');
	});

	it('marks the access token expirable so core can refresh it', () => {
		const accessToken = credential.properties.find((p) => p.name === 'accessToken');
		expect(accessToken?.typeOptions?.expirable).toBe(true);
	});

	it('masks every sensitive field', () => {
		for (const name of ['accessToken', 'clientSecret', 'privateKey', 'certificate']) {
			const property = credential.properties.find((p) => p.name === name);
			expect(property?.typeOptions?.password, `${name} should be masked`).toBe(true);
		}
	});

	it('tests the credential against the organization endpoint', () => {
		expect(credential.test.request.url).toBe('/v1.0/organization');
		expect(credential.test.request.method).toBe('GET');
	});

	it('attaches the cached bearer token and preserves existing headers', async () => {
		const result = await credential.authenticate(
			{ accessToken: 'token-abc' },
			{ url: 'https://graph.microsoft.com/v1.0/users', headers: { Accept: 'application/json' } },
		);

		expect(result.headers).toEqual({
			Accept: 'application/json',
			Authorization: 'Bearer token-abc',
		});
	});

	it('mints a token through the node helper during preAuthentication', async () => {
		const httpRequest = vi.fn(async () => ({ access_token: 'fresh-token' }));
		const helper = { helpers: { httpRequest } } as never;

		const result = await credential.preAuthentication.call(helper, credentials());

		expect(result).toEqual({ accessToken: 'fresh-token' });
		expect(httpRequest).toHaveBeenCalledOnce();
	});
});
