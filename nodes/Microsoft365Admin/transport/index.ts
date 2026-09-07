import {
	NodeApiError,
	type IDataObject,
	type IExecuteFunctions,
	type IExecuteSingleFunctions,
	type IHttpRequestMethods,
	type IHttpRequestOptions,
	type ILoadOptionsFunctions,
	type JsonObject,
} from 'n8n-workflow';

import type { GraphRequestExtras } from '../helpers/interfaces';
import { isErrorStatus, throwGraphError, type GraphErrorContext } from './errors';

export { graphError, resolveGraphError, throwGraphError } from './errors';
export type { GraphErrorContext, ErrorResolution } from './errors';

const CREDENTIAL_TYPE = 'microsoft365AdminServicePrincipalApi';

/**
 * 401 responses must reach `httpRequestWithAuthentication` so an expired token can
 * refresh; everything else in the 4xx/5xx range is inspected here instead of thrown
 * by the HTTP helper, so Graph's error envelope can be translated first.
 */
export const IGNORE_STATUS_ERRORS = {
	ignore: true as const,
	except: [401],
};

type RequestContext = IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions;

/**
 * Resolves the Graph host from the credential, tolerating a stored trailing slash.
 */
export async function getGraphApiBaseUrl(this: RequestContext): Promise<string> {
	const credentials = await this.getCredentials(CREDENTIAL_TYPE);
	return (
		typeof credentials.graphApiBaseUrl === 'string' && credentials.graphApiBaseUrl !== ''
			? credentials.graphApiBaseUrl
			: 'https://graph.microsoft.com'
	).replace(/\/+$/, '');
}

/**
 * Reads the HTTP status off whatever the auth helper threw.
 *
 * Core wraps the failure in a `NodeApiError` carrying `httpCode` as a string, built from an
 * axios error that carries `response.status` as a number; which shape arrives depends on
 * where in the auth path the request died, so both are accepted.
 */
function thrownStatus(error: unknown): number | undefined {
	const candidate = error as {
		httpCode?: unknown;
		status?: unknown;
		response?: { status?: unknown };
	};

	const httpCode = Number(candidate?.httpCode);
	if (Number.isInteger(httpCode) && httpCode !== 0) {
		return httpCode;
	}

	const status = candidate?.response?.status ?? candidate?.status;
	return typeof status === 'number' ? status : undefined;
}

/**
 * Builds the error context a failed response is interpreted against.
 *
 * `resource` and `operation` do not vary per item, so item 0 is always the right index —
 * and it is the only one a load-options context can read.
 */
function errorContext(context: RequestContext, itemIndex = 0): GraphErrorContext {
	const read = (name: string, fallback?: string) => {
		try {
			return String(
				(context as IExecuteFunctions).getNodeParameter(name, itemIndex, fallback) ??
					fallback ??
					'',
			);
		} catch {
			return fallback ?? '';
		}
	};

	return {
		resource: read('resource'),
		operation: read('operation'),
		getParameter: read,
	};
}

/**
 * Issues a single Graph request and returns its body.
 *
 * Failures are translated into a node error naming the parameter at fault where the rules
 * in `./errors` recognise them. Pass `extras.ignoreHttpStatusErrors` to inspect a failure
 * yourself instead — the raw response comes back untouched.
 */
export async function microsoftApiRequest(
	this: RequestContext,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	extras: GraphRequestExtras & { itemIndex?: number } = {},
): Promise<unknown> {
	const baseUrl = await getGraphApiBaseUrl.call(this);
	const callerHandlesErrors = extras.ignoreHttpStatusErrors !== undefined;

	const options: IHttpRequestOptions = {
		method,
		url: extras.url ?? `${baseUrl}/v1.0${endpoint}`,
		// The `$metadata` endpoints answer with XML, which the property loaders parse
		// themselves. Leaving JSON parsing on for those would hand them an unusable value.
		json: !endpoint.startsWith('/$metadata'),
		headers: extras.headers,
		body,
		qs: extras.qs,
		returnFullResponse: true,
		ignoreHttpStatusErrors: extras.ignoreHttpStatusErrors ?? IGNORE_STATUS_ERRORS,
	};

	let response: { statusCode: number; body: unknown; headers: IDataObject };
	try {
		response = (await this.helpers.httpRequestWithAuthentication.call(
			this,
			CREDENTIAL_TYPE,
			options,
		)) as { statusCode: number; body: unknown; headers: IDataObject };
	} catch (error) {
		// 401 is the one status deliberately left un-ignored, so core mints a fresh token and
		// replays the request. By the time a 401 surfaces here that replay has already run and
		// failed, so the token was not merely stale — and core's message is the bare Graph
		// envelope, which reads as an unexplained `401 - {...}` in the workflow.
		if (thrownStatus(error) === 401) {
			throw new NodeApiError(this.getNode(), error as JsonObject, {
				message: 'Microsoft Graph rejected the access token',
				description:
					'A freshly minted token was rejected as well. Check that the app registration still has a valid client secret or certificate, that admin consent is granted for the application permissions this operation needs, and that the credential names the tenant you expect.',
				...(extras.itemIndex === undefined ? {} : { itemIndex: extras.itemIndex }),
			});
		}

		// Every other failure already arrives as core's own `NodeApiError`, built from the
		// response it saw. Re-wrapping it here would bury that detail behind a second envelope.
		// eslint-disable-next-line @n8n/community-nodes/require-node-api-error
		throw error;
	}

	if (callerHandlesErrors) {
		return extras.returnFullResponse ? response : response.body;
	}

	if (isErrorStatus(response.statusCode)) {
		throwGraphError(
			this.getNode(),
			response,
			errorContext(this, extras.itemIndex),
			extras.itemIndex,
		);
	}

	return extras.returnFullResponse ? response : response.body;
}

/** One page of a Graph collection. */
interface CollectionPage {
	value?: IDataObject[];
	'@odata.nextLink'?: string;
}

/**
 * Walks every page of a Graph collection and returns the concatenated `value` arrays.
 *
 * This drives the `@odata.nextLink` chain through `microsoftApiRequest` rather than through
 * core's `requestWithAuthenticationPaginated`. That helper forces `simple: false` on the
 * legacy request path, so a failing page *resolves* instead of throwing: core never sees a
 * 401, never refreshes the expired token, and the error body — which has no `value` array —
 * was silently dropped, turning an auth failure into an empty result set.
 */
export async function microsoftApiPaginateRequest(
	this: RequestContext,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	extras: Pick<GraphRequestExtras, 'qs' | 'headers' | 'url'> & { itemIndex?: number } = {},
): Promise<IDataObject[]> {
	const results: IDataObject[] = [];
	const visited = new Set<string>();

	// Only the first page is built from `endpoint` + `qs`. A `@odata.nextLink` is an absolute
	// URL that already carries the query string, so re-applying `qs` to it would duplicate
	// every parameter.
	let page = (await microsoftApiRequest.call(
		this,
		method,
		endpoint,
		body,
		extras,
	)) as CollectionPage;

	for (;;) {
		if (page?.value) {
			results.push(...page.value);
		}

		const nextLink = page?.['@odata.nextLink'];
		if (typeof nextLink !== 'string' || nextLink === '') {
			return results;
		}

		// Graph advances the skip token on every page; a link repeating itself means the
		// collection is cycling, and following it further would never terminate.
		if (visited.has(nextLink)) {
			return results;
		}
		visited.add(nextLink);

		// Pages must be walked in order: each `@odata.nextLink` only exists once the page
		// before it has been read.
		// eslint-disable-next-line no-await-in-loop
		page = (await microsoftApiRequest.call(this, method, endpoint, body, {
			headers: extras.headers,
			itemIndex: extras.itemIndex,
			url: nextLink,
		})) as CollectionPage;
	}
}
