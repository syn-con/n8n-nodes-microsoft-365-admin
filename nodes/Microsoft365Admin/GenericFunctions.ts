import type {
	IDataObject,
	IExecuteFunctions,
	IExecuteSingleFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
	// `requestWithAuthenticationPaginated` is current, but its signature still takes the
	// deprecated `IRequestOptions`, and n8n ships no paginated equivalent that accepts
	// `IHttpRequestOptions`. The type is unavoidable until upstream provides one.
	// eslint-disable-next-line @n8n/community-nodes/no-deprecated-workflow-functions
	IRequestOptions,
	INodePropertyOptions,
	INodeListSearchResult,
	INodeListSearchItems,
} from 'n8n-workflow';

import { extractEntityProperties, type DirectoryListResponse } from './utils';

/**
 * Resolves the Graph host from the credential, tolerating a stored trailing slash.
 *
 * The declarative operations get this from `requestDefaults.baseURL`; anything that
 * builds its own request needs the same value.
 */
export async function getGraphApiBaseUrl(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
): Promise<string> {
	const credentials = await this.getCredentials('microsoft365AdminServicePrincipalApi');
	return (
		typeof credentials.graphApiBaseUrl === 'string' && credentials.graphApiBaseUrl !== ''
			? credentials.graphApiBaseUrl
			: 'https://graph.microsoft.com'
	).replace(/\/+$/, '');
}

/**
 * The parts of a Graph request that most call sites leave alone, kept out of the parameter
 * list so the common `(method, endpoint)` and `(method, endpoint, body)` calls stay short.
 */
export interface GraphRequestExtras {
	qs?: IDataObject;
	headers?: IDataObject;
	/** An absolute URL, e.g. an `@odata.nextLink`, used instead of building one. */
	url?: string;
	/** Returns `{ body, headers, statusCode }` instead of just the body. */
	returnFullResponse?: boolean;
	/** Set to inspect an error response yourself rather than having it thrown. */
	ignoreHttpStatusErrors?: IHttpRequestOptions['ignoreHttpStatusErrors'];
}

export async function microsoftApiRequest(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	extras: GraphRequestExtras = {},
): Promise<unknown> {
	const baseUrl = await getGraphApiBaseUrl.call(this);
	const options: IHttpRequestOptions = {
		method,
		url: extras.url ?? `${baseUrl}/v1.0${endpoint}`,
		// The `$metadata` endpoints answer with XML, which the property loaders parse
		// themselves. Leaving JSON parsing on for those would hand them an unusable value.
		json: !endpoint.startsWith('/$metadata'),
		headers: extras.headers,
		body,
		qs: extras.qs,
		returnFullResponse: extras.returnFullResponse,
		ignoreHttpStatusErrors: extras.ignoreHttpStatusErrors,
	};

	return this.helpers.httpRequestWithAuthentication.call(
		this,
		'microsoft365AdminServicePrincipalApi',
		options,
	);
}

export async function microsoftApiPaginateRequest(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	extras: Pick<GraphRequestExtras, 'qs' | 'headers' | 'url'> & { itemIndex?: number } = {},
): Promise<IDataObject[]> {
	const baseUrl = await getGraphApiBaseUrl.call(this);
	// `IHttpRequestOptions` has no `uri` property, which `requestWithAuthenticationPaginated`
	// requires — so the deprecated option type is the only one that fits here.
	// eslint-disable-next-line @n8n/community-nodes/no-deprecated-workflow-functions
	const options: IRequestOptions = {
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
		'microsoft365AdminServicePrincipalApi',
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

export async function getGroupProperties(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const returnData: INodePropertyOptions[] = [];
	const response = await microsoftApiRequest.call(this, 'GET', '/$metadata#groups');

	let properties = extractEntityProperties(response as string, [
		'entity',
		'directoryObject',
		'group',
	]);

	properties = properties.filter(
		(x) => !['id', 'isArchived', 'hasMembersWithLicenseErrors'].includes(x),
	);

	properties = properties.sort();

	for (const property of properties) {
		returnData.push({
			name: property,
			value: property,
		});
	}

	return returnData;
}

/**
 * Lists the tenant's subscribed license SKUs for the "Assign" operation.
 *
 * `/subscribedSkus` returns the full collection in one response — it supports no
 * paging or `$filter`, so there is nothing to paginate over here.
 */
export async function getSubscribedSkus(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const response = (await microsoftApiRequest.call(this, 'GET', '/subscribedSkus')) as {
		value?: Array<{
			skuId: string;
			skuPartNumber: string;
			consumedUnits?: number;
			prepaidUnits?: { enabled?: number };
		}>;
	};

	return (response.value ?? []).map((sku) => {
		const consumed = sku.consumedUnits ?? 0;
		const enabled = sku.prepaidUnits?.enabled ?? 0;
		return {
			// Surfacing seat usage makes it obvious when a SKU has nothing left to assign.
			name: `${sku.skuPartNumber} (${consumed}/${enabled} used)`,
			value: sku.skuId,
		};
	});
}

export async function getUserProperties(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const returnData: INodePropertyOptions[] = [];
	const response = await microsoftApiRequest.call(this, 'GET', '/$metadata#users');

	let properties = extractEntityProperties(response as string, [
		'entity',
		'directoryObject',
		'user',
	]);

	// signInActivity requires AuditLog.Read.All
	// mailboxSettings MailboxSettings.Read
	properties = properties.filter(
		(x) =>
			!['id', 'deviceEnrollmentLimit', 'mailboxSettings', 'print', 'signInActivity'].includes(x),
	);

	properties = properties.sort();

	for (const property of properties) {
		returnData.push({
			name: property,
			value: property,
		});
	}
	return returnData;
}

/**
 * Escapes a search term for an OData string literal, where a quote is written twice.
 *
 * Without this, searching for a name that contains an apostrophe — O'Brien — closes the
 * literal early and Graph rejects the whole query.
 */
function odataLiteral(value: string): string {
	return value.replace(/'/g, "''");
}

/**
 * Escapes a search term for a `$search` phrase, which is KQL inside double quotes.
 *
 * The backslash goes first, so an escape added here is not itself escaped again.
 */
function searchPhrase(value: string): string {
	return value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

export async function getGroups(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	let response: DirectoryListResponse;
	if (paginationToken) {
		response = (await microsoftApiRequest.call(
			this,
			'GET',
			'/groups',
			{},
			{
				url: paginationToken,
			},
		)) as DirectoryListResponse;
	} else {
		const qs: IDataObject = {
			$select: 'id,displayName',
		};
		const headers: IDataObject = {};
		if (filter) {
			headers.ConsistencyLevel = 'eventual';
			qs.$search = `"displayName:${searchPhrase(filter)}"`;
		}
		response = (await microsoftApiRequest.call(
			this,
			'GET',
			'/groups',
			{},
			{ qs, headers },
		)) as DirectoryListResponse;
	}

	const groups: Array<{
		id: string;
		displayName: string;
	}> = response.value ?? [];

	const results: INodeListSearchItems[] = groups
		.map((g) => ({
			// Sorting compares names, so a directory object without one must not be null.
			name: g.displayName ?? g.id,
			value: g.id,
		}))
		.sort((a, b) =>
			a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }),
		);

	return { results, paginationToken: response['@odata.nextLink'] };
}

export async function getUsers(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	let response: DirectoryListResponse;
	if (paginationToken) {
		response = (await microsoftApiRequest.call(
			this,
			'GET',
			'/users',
			{},
			{
				url: paginationToken,
			},
		)) as DirectoryListResponse;
	} else {
		const qs: IDataObject = {
			$select: 'id,displayName',
		};
		const headers: IDataObject = {};
		if (filter) {
			const term = odataLiteral(filter);
			qs.$filter = `startsWith(displayName, '${term}') OR startsWith(userPrincipalName, '${term}')`;
		}
		response = (await microsoftApiRequest.call(
			this,
			'GET',
			'/users',
			{},
			{ qs, headers },
		)) as DirectoryListResponse;
	}

	const users: Array<{
		id: string;
		displayName: string;
	}> = response.value ?? [];

	const results: INodeListSearchItems[] = users
		.map((u) => ({
			name: u.displayName ?? u.id,
			value: u.id,
		}))
		.sort((a, b) =>
			a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }),
		);

	return { results, paginationToken: response['@odata.nextLink'] };
}
