import type {
	IDataObject,
	IExecuteFunctions,
	IExecuteSingleFunctions,
	IHttpRequestOptions,
	INodeProperties,
	ILoadOptionsFunctions,
	IN8nHttpFullResponse,
	INodeExecutionData,
} from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import {
	annotateMethodsPostReceive,
	validatePathSegmentsPreSend,
	describeMethod,
	executeResetPassword,
	generatePassword,
	getAuthenticationMethods,
	methodTypeOf,
} from './AuthenticationFunctions';

const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';
const RESPONSE = { statusCode: 200, body: {}, headers: {} } as unknown as IN8nHttpFullResponse;

interface Scenario {
	items?: number;
	parameters?: Record<string, unknown>;
	/** What the Graph GET answers with, for the picker. */
	body?: IDataObject;
	continueOnFail?: boolean;
}

function context(scenario: Scenario = {}) {
	const requests: Array<{ method?: string; url: string; body: unknown }> = [];

	const httpRequestWithAuthentication = vi.fn(
		async (_credentialType: string, options: IHttpRequestOptions) => {
			requests.push({ method: options.method, url: options.url, body: options.body });
			return scenario.body ?? {};
		},
	);

	const parameter = (name: string, itemIndex = 0) => {
		const value = scenario.parameters?.[name];
		if (typeof value === 'function') return (value as (index: number) => unknown)(itemIndex);
		return value;
	};

	/**
	 * Stands in for n8n's parameter resolution, including the two behaviours that bite:
	 * `extractValue` resolves against the properties the operation *displays* and throws
	 * for one it does not, while a `name.value` path is a plain lookup that just misses.
	 */
	const resolve = (name: string, itemIndex = 0, options?: { extractValue?: boolean }) => {
		const declared = scenario.parameters ?? {};

		if (name.endsWith('.value')) {
			const base = parameter(name.slice(0, -'.value'.length), itemIndex);
			if (base !== null && typeof base === 'object') return (base as IDataObject).value;
			return base;
		}
		if (options?.extractValue && !(name in declared)) {
			throw new Error('Could not find property');
		}
		return parameter(name, itemIndex);
	};

	const ctx = {
		getInputData: () =>
			Array.from({ length: scenario.items ?? 1 }, (_, index) => ({ json: { index } })),
		getNodeParameter: (
			name: string,
			itemIndex: number,
			fallback?: unknown,
			options?: { extractValue?: boolean },
		) => resolve(name, itemIndex, options) ?? fallback,
		getCurrentNodeParameter: (name: string, options?: { extractValue?: boolean }) =>
			resolve(name, 0, options),
		getNode: () => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' }),
		getCredentials: async () => ({}),
		continueOnFail: () => scenario.continueOnFail ?? false,
		helpers: { httpRequestWithAuthentication },
	};

	return { ctx, requests, httpRequestWithAuthentication };
}

describe('method types', () => {
	it.each([
		['#microsoft.graph.phoneAuthenticationMethod', 'phoneMethods'],
		[
			'#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
			'microsoftAuthenticatorMethods',
		],
		['#microsoft.graph.fido2AuthenticationMethod', 'fido2Methods'],
		['#microsoft.graph.passwordAuthenticationMethod', 'passwordMethods'],
		// Graph is not always consistent about the leading hash.
		['microsoft.graph.emailAuthenticationMethod', 'emailMethods'],
	])('maps %s to the collection that addresses it', (odataType, expected) => {
		expect(methodTypeOf(odataType)).toBe(expected);
	});

	it.each([undefined, null, 42, '#microsoft.graph.somethingNewAuthenticationMethod'])(
		'has nothing to say about %s',
		(odataType) => {
			expect(methodTypeOf(odataType)).toBeUndefined();
		},
	);
});

describe('method labels', () => {
	it('uses the name a security key was registered under', () => {
		expect(describeMethod({ displayName: 'Red key', model: 'NFC key' }, 'fido2Methods')).toBe(
			'Red key (NFC key)',
		);
	});

	it('names a phone by its number and type', () => {
		expect(
			describeMethod({ phoneNumber: '+370 600 00000', phoneType: 'mobile' }, 'phoneMethods'),
		).toBe('+370 600 00000 (mobile)');
	});

	it('falls back to the method type when Graph gives no name at all', () => {
		expect(
			describeMethod(
				{ id: 'tap-1', createdDateTime: '2026-08-21T09:00:00Z' },
				'temporaryAccessPassMethods',
			),
		).toBe('Temporary Access Pass (2026-08-21T09:00:00Z)');
	});

	it('falls back to the ID when even the type is unknown', () => {
		expect(describeMethod({ id: 'method-1' })).toBe('method-1');
	});
});

describe('annotating a list of methods', () => {
	const items = (): INodeExecutionData[] => [
		{
			json: {
				'@odata.type': '#microsoft.graph.phoneAuthenticationMethod',
				id: 'p1',
				phoneNumber: '+1',
			},
			index: 0,
		},
		{
			json: { '@odata.type': '#microsoft.graph.passwordAuthenticationMethod', id: 'pw' },
			index: 1,
		},
		{
			json: { '@odata.type': '#microsoft.graph.newFangledAuthenticationMethod', id: 'x' },
			index: 2,
		},
	];

	async function run() {
		const { ctx } = context();
		return await annotateMethodsPostReceive.call(
			ctx as unknown as IExecuteSingleFunctions,
			items(),
			RESPONSE,
		);
	}

	it('adds the collection a Delete Method step needs', async () => {
		const result = await run();
		expect(result[0].json.methodType).toBe('phoneMethods');
	});

	it('marks the password method as one that cannot be deleted', async () => {
		const result = await run();
		expect(result[1].json).toMatchObject({ methodType: 'passwordMethods', deletable: false });
		expect(result[0].json.deletable).toBe(true);
	});

	it('reports an unrecognized method type rather than guessing one', async () => {
		const result = await run();
		expect(result[2].json).toMatchObject({ methodType: null, deletable: false });
	});

	it('keeps everything Graph returned', async () => {
		const result = await run();
		expect(result[0].json).toMatchObject({ id: 'p1', phoneNumber: '+1' });
		expect(result).toHaveLength(3);
	});
});

describe('the method picker', () => {
	async function list(scenario: Scenario) {
		const { ctx, requests } = context(scenario);
		const result = await getAuthenticationMethods.call(ctx as unknown as ILoadOptionsFunctions);
		return { result, requests };
	}

	it('lists the methods of the chosen type for the chosen user', async () => {
		const { result, requests } = await list({
			parameters: { user: USER, methodType: 'phoneMethods' },
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

		expect(requests[0].url).toBe(
			`https://graph.microsoft.com/v1.0/users/${USER}/authentication/methods`,
		);
		expect(result.results.map((r) => r.value)).toEqual(['a', 'b']);
	});

	it('filters the aggregate response to the selected method type', async () => {
		const { result } = await list({
			parameters: { user: USER, methodType: 'platformCredentialMethods' },
			body: {
				value: [
					{
						'@odata.type': '#microsoft.graph.passwordAuthenticationMethod',
						id: 'password',
					},
				],
			},
		});

		expect(result.results).toEqual([]);
	});

	it('filters on the label, since these collections take no $filter', async () => {
		const { ctx } = context({
			parameters: { user: USER, methodType: 'fido2Methods' },
			body: {
				value: [
					{
						'@odata.type': '#microsoft.graph.fido2AuthenticationMethod',
						id: 'k1',
						displayName: 'Red key',
					},
					{
						'@odata.type': '#microsoft.graph.fido2AuthenticationMethod',
						id: 'k2',
						displayName: 'Blue key',
					},
				],
			},
		});

		const result = await getAuthenticationMethods.call(
			ctx as unknown as ILoadOptionsFunctions,
			'red',
		);

		expect(result.results).toEqual([{ name: 'Red key', value: 'k1' }]);
	});

	it('asks for a user before it can offer anything', async () => {
		await expect(list({ parameters: { user: '', methodType: 'phoneMethods' } })).rejects.toThrow(
			/Choose a user first/,
		);
	});

	it('refuses a user value that would reshape the request path', async () => {
		await expect(
			list({ parameters: { user: 'alice/authentication/methods?x=', methodType: 'phoneMethods' } }),
		).rejects.toThrow(/Choose a user first/);
	});

	it('asks for a method type it recognizes', async () => {
		await expect(list({ parameters: { user: USER, methodType: 'madeUpMethods' } })).rejects.toThrow(
			/Choose a method type first/,
		);
	});

	it('copes with a user who has no methods of that type', async () => {
		const { result } = await list({
			parameters: { user: USER, methodType: 'fido2Methods' },
			body: {},
		});
		expect(result.results).toEqual([]);
	});
});

describe('path safety', () => {
	const options = () =>
		({
			url: 'https://graph.microsoft.com/v1.0/users/x/authentication/methods',
		}) as IHttpRequestOptions;

	async function guard(parameters: Record<string, unknown>) {
		const { ctx } = context({ parameters });
		return await validatePathSegmentsPreSend.call(
			{
				...ctx,
				// IExecuteSingleFunctions takes the fallback as the second argument.
				getNodeParameter: (name: string, fallback?: unknown) =>
					ctx.getNodeParameter(name, 0, fallback),
			} as unknown as IExecuteSingleFunctions,
			options(),
		);
	}

	it('lets an object ID and a userPrincipalName through', async () => {
		await expect(guard({ user: USER })).resolves.toBeDefined();
		await expect(guard({ user: "o'brien.a@contoso.com" })).resolves.toBeDefined();
	});

	it('lets a base64url method ID through', async () => {
		await expect(
			guard({ user: USER, methodType: 'fido2Methods', method: '-2_GRUg2-HYz6_1YG4YRAQ2' }),
		).resolves.toBeDefined();
	});

	// Left unchecked, this would move the request to a different Graph endpoint and leave
	// the intended path in the query string.
	it.each([
		['user', 'alice/revokeSignInSessions?x='],
		['method', '../../../users/bob'],
		['methodType', 'methods%2F..'],
	])('refuses an unsafe %s', async (name, value) => {
		await expect(guard({ user: USER, [name]: value })).rejects.toThrow(
			/characters that are not allowed/,
		);
	});

	// Get Many Methods does not display `method` or `methodType`, and asking n8n to extract
	// a value from a property it is not showing fails with "Could not find property".
	it('says nothing about a parameter the operation does not have', async () => {
		await expect(guard({ user: USER })).resolves.toBeDefined();
	});

	it('reads the ID out of a resource locator', async () => {
		await expect(
			guard({ user: { __rl: true, mode: 'list', value: 'alice/revokeSignInSessions?x=' } }),
		).rejects.toThrow(/characters that are not allowed/);
	});
});

describe('generated passwords', () => {
	it('is as long as asked', () => {
		expect(generatePassword(24)).toHaveLength(24);
	});

	it('carries all four character classes, so Entra accepts it', () => {
		for (let attempt = 0; attempt < 50; attempt++) {
			const password = generatePassword(12);
			expect(password).toMatch(/[a-z]/);
			expect(password).toMatch(/[A-Z]/);
			expect(password).toMatch(/[0-9]/);
			expect(password).toMatch(/[!#$%&*+\-=?@^_]/);
		}
	});

	it('leaves out the characters that get misread', () => {
		for (let attempt = 0; attempt < 50; attempt++) {
			expect(generatePassword(32)).not.toMatch(/[lIO01]/);
		}
	});

	it('does not repeat itself', () => {
		const passwords = new Set(Array.from({ length: 20 }, () => generatePassword(16)));
		expect(passwords.size).toBe(20);
	});
});

describe('reset password', () => {
	async function run(scenario: Scenario) {
		const { ctx, requests } = context(scenario);
		const [output] = await executeResetPassword.call(ctx as unknown as IExecuteFunctions);
		return { output, requests };
	}

	it('patches passwordProfile with the password it was given', async () => {
		const { output, requests } = await run({
			parameters: { user: USER, options: { password: 'Correct-Horse-42' } },
		});

		expect(requests[0]).toMatchObject({
			method: 'PATCH',
			url: `https://graph.microsoft.com/v1.0/users/${USER}`,
		});
		expect(requests[0].body).toEqual({
			passwordProfile: { password: 'Correct-Horse-42', forceChangePasswordNextSignIn: true },
		});
		expect(output[0].json).toMatchObject({
			id: USER,
			passwordReset: true,
			password: 'Correct-Horse-42',
			generated: false,
		});
	});

	it('generates a password when none is given and hands it back', async () => {
		const { output, requests } = await run({ parameters: { user: USER } });

		const body = requests[0].body as { passwordProfile: { password: string } };
		expect(output[0].json.generated).toBe(true);
		expect(output[0].json.password).toBe(body.passwordProfile.password);
		expect(String(output[0].json.password)).toHaveLength(16);
	});

	it('honours the requested length of a generated password', async () => {
		const { output } = await run({ parameters: { user: USER, options: { passwordLength: 32 } } });
		expect(String(output[0].json.password)).toHaveLength(32);
	});

	it('keeps a length Graph would reject inside the allowed range', async () => {
		const { output } = await run({ parameters: { user: USER, options: { passwordLength: 2 } } });
		expect(String(output[0].json.password).length).toBeGreaterThanOrEqual(8);
	});

	it('asks for MFA before the change when told to', async () => {
		const { requests } = await run({
			parameters: {
				user: USER,
				options: { forceChangePassword: 'forceChangePasswordNextSignInWithMfa' },
			},
		});

		expect(requests[0].body).toMatchObject({
			passwordProfile: { forceChangePasswordNextSignInWithMfa: true },
		});
	});

	it('leaves the password in place when no change is required', async () => {
		const { output, requests } = await run({
			parameters: { user: USER, options: { forceChangePassword: 'never' } },
		});

		expect(requests[0].body).toMatchObject({
			passwordProfile: { forceChangePasswordNextSignIn: false },
		});
		expect(output[0].json.forceChangePasswordNextSignIn).toBe(false);
	});

	it('resets one user per item, in order', async () => {
		const { output, requests } = await run({
			items: 3,
			parameters: { user: (index: number) => `user-${index}` },
		});

		expect(requests.map((r) => r.url.split('/users/')[1])).toEqual(['user-0', 'user-1', 'user-2']);
		expect(output.map((item) => item.pairedItem)).toEqual([{ item: 0 }, { item: 1 }, { item: 2 }]);
	});

	it('refuses a user value that would reshape the request path', async () => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'alice/authentication/methods?x=' },
		});

		await expect(executeResetPassword.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/characters that are not allowed/,
		);
		expect(httpRequestWithAuthentication).not.toHaveBeenCalled();
	});

	it('rejects an item with no user', async () => {
		const { ctx } = context({ parameters: { user: '' } });

		await expect(executeResetPassword.call(ctx as unknown as IExecuteFunctions)).rejects.toThrow(
			/No user was given/,
		);
	});

	it('reports the failure per item and keeps going', async () => {
		const { output, requests } = await run({
			items: 3,
			continueOnFail: true,
			parameters: { user: (index: number) => (index === 1 ? '' : `user-${index}`) },
		});

		expect(requests).toHaveLength(2);
		expect(output).toHaveLength(3);
		expect(output[1].json.error).toMatch(/No user was given/);
	});
});

beforeEach(() => {
	vi.clearAllMocks();
});
