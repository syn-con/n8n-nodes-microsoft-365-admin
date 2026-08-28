import type {
	IDataObject,
	IExecuteFunctions,
	IExecuteSingleFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
} from 'n8n-workflow';

import type { GraphRequestExtras, PaginatedRequestOptions } from '../helpers/interfaces';
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

	const response = (await this.helpers.httpRequestWithAuthentication.call(
		this,
		CREDENTIAL_TYPE,
		options,
	)) as { statusCode: number; body: unknown; headers: IDataObject };

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

/**
 * Walks every page of a Graph collection and returns the concatenated `value` arrays.
 */
export async function microsoftApiPaginateRequest(
	this: RequestContext,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	extras: Pick<GraphRequestExtras, 'qs' | 'headers' | 'url'> & { itemIndex?: number } = {},
): Promise<IDataObject[]> {
	const baseUrl = await getGraphApiBaseUrl.call(this);

	const options: PaginatedRequestOptions = {
		method,
		uri: extras.url ?? `${baseUrl}/v1.0${endpoint}`,
		json: true,
		headers: extras.headers,
		body,
		qs: extras.qs,
	};

	const pages = await this.helpers.requestWithAuthenticationPaginated.call(
		this,
		options,
		extras.itemIndex ?? 0,
		{
			continue: '={{ !!$response.body?.["@odata.nextLink"] }}',
			request: {
				url: '={{ $response.body?.["@odata.nextLink"] ?? $request.url }}',
			},
			requestInterval: 0,
		},
		CREDENTIAL_TYPE,
	);

	let results: IDataObject[] = [];
	for (const page of pages) {
		const items = page.body.value as IDataObject[];
		if (items) {
			results = results.concat(items);
		}
	}

	return results;
}
