import type {
	IDataObject,
	IExecuteFunctions,
	IHttpRequestOptions,
	IN8nHttpFullResponse,
} from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('n8n-workflow', async (importOriginal) => {
	const actual = await importOriginal<typeof import('n8n-workflow')>();
	// The real backoff waits tens of seconds.
	return { ...actual, sleep: vi.fn(async () => {}) };
});

const { sleep } = await import('n8n-workflow');
const { executeLicenseWrite } = await import('./LicenseFunctions');

const E3 = '05e9a617-0261-4cee-bb44-138d3ef5d965';
const E5 = 'c7df2760-2c81-4ef7-b578-5b5392b571df';
const TEAMS_PLAN = '57ff2da0-773e-42df-b2af-ffb7a2317929';
const EXCHANGE_PLAN = '9aaf7827-d63c-4b61-89c3-182f06f82e5c';

type Response = Partial<IN8nHttpFullResponse>;

interface Scenario {
	items?: number;
	parameters?: Record<string, unknown>;
	/** Consumed in order; the last entry answers every further request. */
	responses?: Response[];
	continueOnFail?: boolean;
	credentials?: IDataObject;
	/** What `/subscribedSkus` answers, for the disabled-plan lookup. */
	subscribedSkus?: IDataObject;
}

const conflict: Response = {
	statusCode: 400,
	body: {
		error: {
			code: 'Directory_ConcurrencyViolation',
			message:
				'Error due to concurrent requests being made to the tenant. Please wait briefly and retry.',
		},
	},
};

function context(scenario: Scenario = {}) {
	const requests: Array<{ url: string; body: IDataObject }> = [];
	const queue = [...(scenario.responses ?? [])];
	let inFlight = 0;
	let peakInFlight = 0;
	let skuLookups = 0;

	// Stands in for n8n's transport rather than for `microsoftApiRequest`, so the URL the
	// node really builds is the one under test.
	const httpRequestWithAuthentication = vi.fn(
		async (_credentialType: string, options: IHttpRequestOptions) => {
			inFlight++;
			peakInFlight = Math.max(peakInFlight, inFlight);
			// Yield, so overlapping requests would show up in peakInFlight.
			await Promise.resolve();
			inFlight--;

			if (options.url.endsWith('/subscribedSkus')) {
				skuLookups++;
				return scenario.subscribedSkus ?? { value: [] };
			}

			requests.push({ url: options.url, body: structuredClone(options.body) as IDataObject });
			const next = queue.length > 1 ? queue.shift()! : (queue[0] ?? {});
			const response = { statusCode: 200, body: {}, headers: {}, ...next };

			return options.returnFullResponse ? response : response.body;
		},
	);

	const ctx = {
		getInputData: () =>
			Array.from({ length: scenario.items ?? 1 }, (_, index) => ({ json: { index } })),
		getNodeParameter: (name: string, itemIndex: number, fallback?: unknown) => {
			const value = scenario.parameters?.[name];
			if (typeof value === 'function') return (value as (index: number) => unknown)(itemIndex);
			return value ?? fallback;
		},
		getNode: () => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' }),
		getCredentials: async () => scenario.credentials ?? {},
		continueOnFail: () => scenario.continueOnFail ?? false,
		helpers: { httpRequestWithAuthentication },
	};

	return {
		ctx: ctx as unknown as IExecuteFunctions,
		requests,
		httpRequestWithAuthentication,
		peak: () => peakInFlight,
		skuLookups: () => skuLookups,
	};
}

/** The body of a single Graph request, as it went out. */
function body(requests: Array<{ body: IDataObject }>, index = 0) {
	return requests[index].body;
}

/** The two SKUs used by the disabled-plan tests, with one service plan each. */
const TENANT_SKUS: IDataObject = {
	value: [
		{ skuId: E3, servicePlans: [{ servicePlanId: EXCHANGE_PLAN }] },
		{ skuId: E5, servicePlans: [{ servicePlanId: TEAMS_PLAN }] },
	],
};

beforeEach(() => {
	vi.mocked(sleep).mockClear();
});

describe('assign', () => {
	it('adds every selected SKU in one request', async () => {
		const { ctx, requests } = context({ parameters: { user: 'user-1', skuId: [E3, E5] } });

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(1);
		expect(requests[0].url).toBe('https://graph.microsoft.com/v1.0/users/user-1/assignLicense');
		expect(body(requests)).toEqual({
			addLicenses: [
				{ skuId: E3, disabledPlans: [] },
				{ skuId: E5, disabledPlans: [] },
			],
			removeLicenses: [],
		});
	});

	it('accepts a single ID or a comma-separated list from an expression', async () => {
		const { ctx, requests } = context({ parameters: { user: 'user-1', skuId: `${E3}, ${E5}` } });

		await executeLicenseWrite.call(ctx, 'assign');

		expect(body(requests).addLicenses).toEqual([
			{ skuId: E3, disabledPlans: [] },
			{ skuId: E5, disabledPlans: [] },
		]);
	});

	it('swaps licenses in a single request', async () => {
		const { ctx, requests } = context({
			parameters: { user: 'user-1', skuId: [E5], options: { removeSkuIds: [E3] } },
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(1);
		expect(body(requests)).toEqual({
			addLicenses: [{ skuId: E5, disabledPlans: [] }],
			removeLicenses: [E3],
		});
	});

	it('never asks Graph to add and remove the same SKU', async () => {
		const { ctx, requests } = context({
			parameters: { user: 'user-1', skuId: [E5], options: { removeSkuIds: [E5, E3] } },
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(body(requests).removeLicenses).toEqual([E3]);
	});

	it('targets the Graph host from the credential', async () => {
		const { ctx, requests } = context({
			parameters: { user: 'user-1', skuId: [E3] },
			credentials: { graphApiBaseUrl: 'https://graph.microsoft.us/' },
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests[0].url).toBe('https://graph.microsoft.us/v1.0/users/user-1/assignLicense');
	});

	it('rejects an item with no SKU selected', async () => {
		const { ctx, httpRequestWithAuthentication } = context({ parameters: { user: 'user-1' } });

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(/No license SKU/);
		expect(httpRequestWithAuthentication).not.toHaveBeenCalled();
	});

	it('rejects an item with no user', async () => {
		const { ctx } = context({ parameters: { skuId: [E3] } });

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(/No user was given/);
	});

	// A target arriving from workflow data must not be able to point the request at another
	// Graph endpoint: `alice/revokeSignInSessions?x=` would leave `/assignLicense` in the
	// query string and post to whatever path came before it.
	it.each(['alice/revokeSignInSessions?x=', '../../groups/g1', 'alice%2Fbob', 'has space'])(
		'refuses to build a URL from the target %s',
		async (user) => {
			const { ctx, httpRequestWithAuthentication } = context({
				parameters: { user, skuId: [E3] },
			});

			await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(
				/characters that are not allowed/,
			);
			expect(httpRequestWithAuthentication).not.toHaveBeenCalled();
		},
	);

	it('still accepts an object ID or a userPrincipalName', async () => {
		const { ctx, requests } = context({
			items: 2,
			parameters: {
				user: (index: number) =>
					index === 0 ? '02bd9fd6-8f93-4758-87c3-1fb73740a315' : "o'brien.a@contoso.com",
				skuId: [E3],
			},
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(2);
	});
});

describe('unassign', () => {
	it('removes every selected SKU in one request', async () => {
		const { ctx, requests } = context({ parameters: { user: 'user-1', skuId: [E3, E5] } });

		await executeLicenseWrite.call(ctx, 'unassign');

		expect(requests).toHaveLength(1);
		expect(body(requests)).toEqual({ addLicenses: [], removeLicenses: [E3, E5] });
	});
});

describe('assignGroup', () => {
	it('keeps the item usable when Entra accepts the write with an empty body', async () => {
		const { ctx } = context({
			parameters: { group: 'group-1', skuId: [E3] },
			// What a 202 Accepted with no body arrives as.
			responses: [{ statusCode: 202, body: '' as unknown as IDataObject }],
		});

		const [output] = await executeLicenseWrite.call(ctx, 'assignGroup');

		expect(output[0].json).toEqual({});
	});

	it('posts to the group endpoint', async () => {
		const { ctx, requests } = context({ parameters: { group: 'group-1', skuId: [E3] } });

		await executeLicenseWrite.call(ctx, 'assignGroup');

		expect(requests[0].url).toBe('https://graph.microsoft.com/v1.0/groups/group-1/assignLicense');
	});

	it('names the group in the error when it is missing', async () => {
		const { ctx } = context({ parameters: { skuId: [E3] } });

		await expect(executeLicenseWrite.call(ctx, 'assignGroup')).rejects.toThrow(
			/No group was given/,
		);
	});
});

describe('disabled plans', () => {
	it('sends each plan only with the SKU that contains it', async () => {
		const { ctx, requests } = context({
			subscribedSkus: TENANT_SKUS,
			parameters: {
				user: 'user-1',
				skuId: [E3, E5],
				options: { disabledPlans: `${TEAMS_PLAN},${EXCHANGE_PLAN}` },
			},
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(body(requests).addLicenses).toEqual([
			{ skuId: E3, disabledPlans: [EXCHANGE_PLAN] },
			{ skuId: E5, disabledPlans: [TEAMS_PLAN] },
		]);
	});

	it('reads the tenant SKUs once for the whole run', async () => {
		const { ctx, skuLookups } = context({
			items: 3,
			subscribedSkus: TENANT_SKUS,
			parameters: {
				user: (index: number) => `user-${index}`,
				skuId: [E3],
				options: { disabledPlans: EXCHANGE_PLAN },
			},
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(skuLookups()).toBe(1);
	});

	it('does not look up service plans when none are disabled', async () => {
		const { ctx, skuLookups } = context({ parameters: { user: 'user-1', skuId: [E3] } });

		await executeLicenseWrite.call(ctx, 'assign');

		expect(skuLookups()).toBe(0);
	});

	it('reports a plan that belongs to none of the selected SKUs', async () => {
		const { ctx } = context({
			subscribedSkus: TENANT_SKUS,
			parameters: { user: 'user-1', skuId: [E3], options: { disabledPlans: TEAMS_PLAN } },
		});

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(
			/not part of any selected license SKU/,
		);
	});
});

describe('combining items', () => {
	it('folds items for the same user into one request', async () => {
		const { ctx, requests } = context({
			items: 3,
			parameters: {
				user: (index: number) => (index === 2 ? 'user-2' : 'user-1'),
				skuId: (index: number) => [index === 1 ? E5 : E3],
			},
		});

		const [output] = await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(2);
		expect(body(requests, 0).addLicenses).toEqual([
			{ skuId: E3, disabledPlans: [] },
			{ skuId: E5, disabledPlans: [] },
		]);
		expect(body(requests, 1).addLicenses).toEqual([{ skuId: E3, disabledPlans: [] }]);

		// Every input item still gets its own output item, in order.
		expect(output).toHaveLength(3);
		expect(output.map((item) => item.pairedItem)).toEqual([{ item: 0 }, { item: 1 }, { item: 2 }]);
	});

	it('gives each item of a merged request its own copy of the response', async () => {
		const { ctx } = context({
			items: 2,
			parameters: { user: 'user-1', skuId: [E3] },
			responses: [{ statusCode: 200, body: { id: 'user-1', assignedLicenses: [{ skuId: E3 }] } }],
		});

		const [output] = await executeLicenseWrite.call(ctx, 'assign');

		expect(output[0].json).toEqual(output[1].json);
		expect(output[0].json).not.toBe(output[1].json);
		expect(output[0].json.assignedLicenses).not.toBe(output[1].json.assignedLicenses);
	});

	it('sends one request per item when switched off', async () => {
		const { ctx, requests } = context({
			items: 3,
			parameters: { user: 'user-1', skuId: [E3], options: { combineItems: false } },
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(3);
	});

	it('keeps items apart when the same SKU carries different disabled plans', async () => {
		const { ctx, requests } = context({
			items: 2,
			subscribedSkus: { value: [{ skuId: E3, servicePlans: [{ servicePlanId: TEAMS_PLAN }] }] },
			parameters: {
				user: 'user-1',
				skuId: [E3],
				options: (index: number) => ({ disabledPlans: index === 0 ? TEAMS_PLAN : '' }),
			},
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(2);
		expect(body(requests, 0).addLicenses).toEqual([{ skuId: E3, disabledPlans: [TEAMS_PLAN] }]);
		expect(body(requests, 1).addLicenses).toEqual([{ skuId: E3, disabledPlans: [] }]);
	});

	it('does not put a SKU in both lists when merging a swap', async () => {
		const { ctx, requests } = context({
			items: 2,
			parameters: {
				user: 'user-1',
				skuId: (index: number) => [index === 0 ? E3 : E5],
				options: (index: number) => (index === 1 ? { removeSkuIds: [E3] } : {}),
			},
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(2);
		expect(body(requests, 0)).toEqual({
			addLicenses: [{ skuId: E3, disabledPlans: [] }],
			removeLicenses: [],
		});
		expect(body(requests, 1)).toEqual({
			addLicenses: [{ skuId: E5, disabledPlans: [] }],
			removeLicenses: [E3],
		});
	});
});

describe('serialization', () => {
	it('never has two writes in flight at once', async () => {
		const { ctx, requests, peak } = context({
			items: 5,
			parameters: { user: (index: number) => `user-${index}`, skuId: [E3] },
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(5);
		expect(peak()).toBe(1);
	});

	it('pauses between requests when asked to', async () => {
		const { ctx } = context({
			items: 3,
			parameters: {
				user: (index: number) => `user-${index}`,
				skuId: [E3],
				options: { waitBetweenRequests: 250 },
			},
		});

		await executeLicenseWrite.call(ctx, 'assign');

		// Between requests only, so one fewer than the number of requests.
		expect(sleep).toHaveBeenCalledTimes(2);
		expect(sleep).toHaveBeenCalledWith(250);
	});

	it('does not pause by default', async () => {
		const { ctx } = context({
			items: 2,
			parameters: { user: (index: number) => `user-${index}`, skuId: [E3] },
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(sleep).not.toHaveBeenCalled();
	});
});

describe('tenant conflicts', () => {
	it('retries a concurrent-request rejection and returns the eventual response', async () => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'user-1', skuId: [E3] },
			responses: [conflict, conflict, { statusCode: 200, body: { id: 'user-1' } }],
		});

		const [output] = await executeLicenseWrite.call(ctx, 'assign');

		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(3);
		expect(output[0].json).toEqual({ id: 'user-1' });
	});

	it('backs off further on each attempt', async () => {
		const { ctx } = context({
			parameters: { user: 'user-1', skuId: [E3] },
			responses: [conflict, conflict, { statusCode: 200, body: {} }],
		});

		await executeLicenseWrite.call(ctx, 'assign');

		const [first, second] = vi.mocked(sleep).mock.calls.map(([ms]) => ms);
		expect(first).toBeGreaterThanOrEqual(3_750);
		expect(second).toBeGreaterThan(first);
	});

	it('gives up after Max Retries with advice on what to do instead', async () => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'user-1', skuId: [E3], options: { maxRetries: 2 } },
			responses: [conflict],
		});

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(/still busy/);
		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(3);
	});

	it('caps how far Max Retries can be pushed', async () => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'user-1', skuId: [E3], options: { maxRetries: 5_000 } },
			responses: [conflict],
		});

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(/still busy/);
		// 20 retries plus the original attempt.
		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(21);
	});

	it('retries even when Entra sends the concurrency message without its code', async () => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'user-1', skuId: [E3] },
			responses: [
				{
					statusCode: 400,
					body: {
						error: {
							code: 'Request_BadRequest',
							message: 'Error due to concurrent requests being made to the tenant.',
						},
					},
				},
				{ statusCode: 200, body: {} },
			],
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(2);
	});

	it('waits the Retry-After a throttled response asks for', async () => {
		const { ctx } = context({
			parameters: { user: 'user-1', skuId: [E3] },
			responses: [
				{ statusCode: 429, body: {}, headers: { 'retry-after': '12' } },
				{ statusCode: 200, body: {} },
			],
		});

		await executeLicenseWrite.call(ctx, 'assign');

		expect(sleep).toHaveBeenCalledExactlyOnceWith(12_000);
	});

	it.each([
		[
			'the plan-conflict message',
			{
				code: 'Request_BadRequest',
				message:
					'License assignment failed because service plan 5136a095-5cf0-4aff-bec3-e84448b38ea5 conflicts with service plan 43de0ff5-c92c-492b-9116-175376d08c38.',
			},
		],
		[
			'the MutuallyExclusiveViolation detail',
			{
				code: 'Request_BadRequest',
				message: 'License assignment failed.',
				details: [{ code: 'MutuallyExclusiveViolation', message: 'conflict' }],
			},
		],
	])('explains overlapping service plans, given %s', async (_label, error) => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'user-1', skuId: [E3, E5] },
			responses: [{ statusCode: 400, body: { error } }],
		});

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow(
			/refused this combination/,
		);
		// Nothing about it improves with time, so it must not burn the retry budget.
		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(1);
		expect(sleep).not.toHaveBeenCalled();
	});

	it('does not retry a request Entra will keep rejecting', async () => {
		const { ctx, httpRequestWithAuthentication } = context({
			parameters: { user: 'user-1', skuId: [E3] },
			responses: [
				{
					statusCode: 400,
					body: {
						error: {
							code: 'Request_BadRequest',
							message: 'License assignment failed because service plan cannot be assigned.',
						},
					},
				},
			],
		});

		await expect(executeLicenseWrite.call(ctx, 'assign')).rejects.toThrow();
		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(1);
		expect(sleep).not.toHaveBeenCalled();
	});

	it('lets 401 through to the auth helper instead of swallowing it', async () => {
		const { ctx, requests } = context({ parameters: { user: 'user-1', skuId: [E3] } });

		await executeLicenseWrite.call(ctx, 'assign');

		const [, options] = vi.mocked(ctx.helpers.httpRequestWithAuthentication).mock.calls[0] as [
			string,
			IHttpRequestOptions,
		];
		expect(requests).toHaveLength(1);
		expect(options.ignoreHttpStatusErrors).toEqual({ ignore: true, except: [401] });
	});
});

describe('continue on fail', () => {
	it('reports the failure per item and keeps going', async () => {
		const { ctx, requests } = context({
			items: 3,
			continueOnFail: true,
			parameters: {
				user: (index: number) => (index === 1 ? '' : `user-${index}`),
				skuId: [E3],
			},
		});

		const [output] = await executeLicenseWrite.call(ctx, 'assign');

		expect(requests).toHaveLength(2);
		expect(output).toHaveLength(3);
		expect(output[1].json.error).toMatch(/No user was given/);
	});

	it('marks every item of a merged request that failed', async () => {
		const { ctx } = context({
			items: 2,
			continueOnFail: true,
			parameters: { user: 'user-1', skuId: [E3], options: { maxRetries: 0 } },
			responses: [conflict],
		});

		const [output] = await executeLicenseWrite.call(ctx, 'assign');

		expect(output).toHaveLength(2);
		expect(output[0].json.error).toMatch(/still busy/);
		expect(output[1].json.error).toMatch(/still busy/);
	});
});
