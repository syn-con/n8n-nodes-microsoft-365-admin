import {
	NodeApiError,
	type JsonObject,
	type IDataObject,
	type IExecuteFunctions,
	type IExecuteSingleFunctions,
	type IHttpRequestMethods,
	type IHttpRequestOptions,
	type ILoadOptionsFunctions,
	// `requestWithAuthenticationPaginated` is current, but its signature still takes the
	// deprecated `IRequestOptions`, and n8n ships no paginated equivalent that accepts
	// `IHttpRequestOptions`. The type is unavoidable until upstream provides one.
	// eslint-disable-next-line @n8n/community-nodes/no-deprecated-workflow-functions
	type IRequestOptions,
	type INodeExecutionData,
	type IN8nHttpFullResponse,
	type INodePropertyOptions,
	type INodeListSearchResult,
	type INodeListSearchItems,
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

/** The ID format quoted back whenever an identifier is unusable. */
const ID_FORMAT = 'The ID should be in the format e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315';

/** n8n strips an empty body object, so a no-op update reaches Graph with no payload. */
const EMPTY_PAYLOAD = 'Empty Payload. JSON content expected.';

/** A message to raise in place of Graph's, or `ignore` to let the response through. */
type ErrorResolution = { message: string; description: string } | 'ignore';

interface ErrorRule {
	/** The Graph error code this rule answers. */
	code: string;
	/** Narrows a code that Graph uses for more than one thing. */
	when?: (message: string) => boolean;
	resolve:
		| ErrorResolution
		| ((this: IExecuteSingleFunctions, message: string, resource: string) => ErrorResolution);
}

/** The only thing that varies between the not-found messages is the parameter to blame. */
function notFound(resource: 'group' | 'user', parameterName: string): ErrorResolution {
	return {
		message: `The required ${resource} doesn't match any existing one`,
		description: `Double-check the value in the parameter '${parameterName}' and try again`,
	};
}

/** Keyed by `resource.operation`; the first matching rule wins. */
const OPERATION_RULES: Record<string, ErrorRule[]> = {
	'group.delete': [
		{ code: 'Request_ResourceNotFound', resolve: notFound('group', 'Group to Delete') },
	],
	'group.get': [{ code: 'Request_ResourceNotFound', resolve: notFound('group', 'Group to Get') }],
	'group.update': [
		{ code: 'BadRequest', when: (message) => message === EMPTY_PAYLOAD, resolve: 'ignore' },
		{ code: 'Request_ResourceNotFound', resolve: notFound('group', 'Group to Update') },
	],
	'user.addGroup': [
		{
			code: 'Request_BadRequest',
			when: (message) =>
				message ===
				"One or more added object references already exist for the following modified properties: 'members'.",
			resolve: {
				message: 'The user is already in the group',
				description:
					'The specified user cannot be added to the group because they are already a member',
			},
		},
		{
			code: 'Request_ResourceNotFound',
			// Graph names whichever object it could not find, so the message decides the blame.
			resolve(message) {
				const group = this.getNodeParameter('group.value') as string;
				return message.includes(group)
					? notFound('group', 'Group')
					: notFound('user', 'User to Add');
			},
		},
	],
	'user.delete': [
		{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Delete') },
	],
	'user.get': [{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Get') }],
	'user.removeGroup': [
		{
			code: 'Request_ResourceNotFound',
			resolve: {
				message: 'The user is not in the group',
				description:
					'The specified user cannot be removed from the group because they are not a member of the group',
			},
		},
		{
			code: 'Request_UnsupportedQuery',
			when: (message) =>
				message === "Unsupported referenced-object resource identifier for link property 'members'.",
			resolve: { message: 'The user ID is invalid', description: ID_FORMAT },
		},
	],
	'user.update': [
		{ code: 'BadRequest', when: (message) => message === EMPTY_PAYLOAD, resolve: 'ignore' },
		{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Update') },
	],
};

/** Tried after the operation's own rules, for failures any operation can hit. */
const GENERIC_RULES: ErrorRule[] = [
	{
		code: 'Request_BadRequest',
		when: (message) => message.startsWith('Invalid object identifier'),
		resolve(message, resource) {
			const group = this.getNodeParameter('group.value', '') as string;
			const parameterResource =
				resource === 'group' || message.includes(group) ? 'group' : 'user';

			return { message: `The ${parameterResource} ID is invalid`, description: ID_FORMAT };
		},
	},
];

/** Reads Graph's error envelope, which is absent on some gateway failures. */
export function graphError(body: unknown): {
	code: string;
	message: string;
	details?: Array<{ code: string; message: string }>;
} {
	const error = (body as { error?: { code?: string; message?: string; details?: [] } })?.error;

	return { code: error?.code ?? '', message: error?.message ?? '', details: error?.details };
}

/** The operation's own rules are tried before the generic ones; the first match wins. */
function findRule(
	resource: string,
	operation: string,
	code: string,
	message: string,
): ErrorRule | undefined {
	const rules = [...(OPERATION_RULES[`${resource}.${operation}`] ?? []), ...GENERIC_RULES];

	return rules.find((rule) => rule.code === code && (!rule.when || rule.when(message)));
}

export async function handleErrorPostReceive(
	this: IExecuteSingleFunctions,
	data: INodeExecutionData[],
	response: IN8nHttpFullResponse,
): Promise<INodeExecutionData[]> {
	const statusCode = String(response.statusCode);
	if (!statusCode.startsWith('4') && !statusCode.startsWith('5')) {
		return data;
	}

	const resource = this.getNodeParameter('resource') as string;
	const operation = this.getNodeParameter('operation') as string;
	const { code, message, details } = graphError(response.body);
	const rule = findRule(resource, operation, code, message);

	if (rule) {
		const resolution =
			typeof rule.resolve === 'function' ? rule.resolve.call(this, message, resource) : rule.resolve;

		// The empty-payload case is n8n's doing, not the user's, so the items pass through.
		if (resolution !== 'ignore') {
			throw new NodeApiError(this.getNode(), response as unknown as JsonObject, resolution);
		}

		return data;
	}

	if (details?.some((detail) => ['ObjectConflict', 'ConflictingObjects'].includes(detail.code))) {
		throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
			message: `The ${resource} already exists`,
			description: message,
		});
	}

	throw new NodeApiError(this.getNode(), response as unknown as JsonObject);
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
		response = (await microsoftApiRequest.call(this, 'GET', '/groups', {}, {
			url: paginationToken,
		})) as DirectoryListResponse;
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
		response = (await microsoftApiRequest.call(this, 'GET', '/users', {}, {
			url: paginationToken,
		})) as DirectoryListResponse;
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
