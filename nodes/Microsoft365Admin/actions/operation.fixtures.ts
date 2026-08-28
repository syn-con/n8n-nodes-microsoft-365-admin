import type {
	IDataObject,
	IExecuteFunctions,
	IHttpRequestOptions,
	INodeExecutionData,
} from 'n8n-workflow';
import { vi } from 'vitest';

/**
 * A stand-in for n8n's execution context, stubbed at the transport boundary.
 *
 * The stub replaces `httpRequestWithAuthentication` rather than `microsoftApiRequest`, so
 * the URL, body and query string an operation really builds are the ones under test.
 */
export interface OperationScenario {
	items?: number;
	/** Values `getNodeParameter` hands back. A function is called with the item index. */
	parameters?: Record<string, unknown>;
	/** Consumed in order; the last entry answers every further request. */
	responses?: Array<{ statusCode?: number; body?: unknown; headers?: IDataObject }>;
	/** Pages `requestWithAuthenticationPaginated` hands back. */
	pages?: Array<{ body: { value?: IDataObject[] } }>;
	continueOnFail?: boolean;
	credentials?: IDataObject;
}

export interface SentRequest {
	method?: string;
	url: string;
	body: IDataObject;
	qs?: IDataObject;
	headers?: IDataObject;
}

export function operationContext(scenario: OperationScenario = {}) {
	const requests: SentRequest[] = [];
	const queue = [...(scenario.responses ?? [])];

	const httpRequestWithAuthentication = vi.fn(
		async (_credentialType: string, options: IHttpRequestOptions) => {
			requests.push({
				method: options.method,
				url: options.url,
				body: (options.body ?? {}) as IDataObject,
				qs: options.qs,
				headers: options.headers,
			});

			const next = queue.length > 1 ? queue.shift()! : (queue[0] ?? {});
			return { statusCode: 200, body: {}, headers: {}, ...next };
		},
	);

	const requestWithAuthenticationPaginated = vi.fn(async () => scenario.pages ?? []);

	const resolve = (name: string, itemIndex = 0, fallback?: unknown) => {
		const value = scenario.parameters?.[name];
		if (typeof value === 'function') {
			return (value as (index: number) => unknown)(itemIndex);
		}
		return value ?? fallback;
	};

	const ctx = {
		getInputData: () =>
			Array.from({ length: scenario.items ?? 1 }, (_, index) => ({ json: { index } })),
		getNodeParameter: (name: string, itemIndex: number, fallback?: unknown) =>
			resolve(name, itemIndex, fallback),
		getNode: () => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' }),
		getCredentials: async () =>
			scenario.credentials ?? { graphApiBaseUrl: 'https://graph.microsoft.com' },
		continueOnFail: () => scenario.continueOnFail ?? false,
		helpers: {
			httpRequestWithAuthentication,
			requestWithAuthenticationPaginated,
			returnJsonArray: (data: IDataObject | IDataObject[]): INodeExecutionData[] =>
				(Array.isArray(data) ? data : [data]).map((json) => ({ json })),
			constructExecutionMetaData: (
				data: INodeExecutionData[],
				meta: { itemData: { item: number } },
			): INodeExecutionData[] => data.map((item) => ({ ...item, pairedItem: meta.itemData })),
		},
	};

	return { ctx: ctx as unknown as IExecuteFunctions, requests, httpRequestWithAuthentication };
}
