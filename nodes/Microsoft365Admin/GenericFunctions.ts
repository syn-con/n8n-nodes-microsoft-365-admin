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

export async function microsoftApiRequest(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	qs?: IDataObject,
	headers?: IDataObject,
	url?: string,
): Promise<unknown> {
	const credentials = await this.getCredentials('microsoft365AdminServicePrincipalApi');
	const baseUrl = (
		typeof credentials.graphApiBaseUrl === 'string' && credentials.graphApiBaseUrl !== ''
			? credentials.graphApiBaseUrl
			: 'https://graph.microsoft.com'
	).replace(/\/+$/, '');
	const options: IHttpRequestOptions = {
		method,
		url: url ?? `${baseUrl}/v1.0${endpoint}`,
		// The `$metadata` endpoints answer with XML, which the property loaders parse
		// themselves. Leaving JSON parsing on for those would hand them an unusable value.
		json: !endpoint.startsWith('/$metadata'),
		headers,
		body,
		qs,
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
	qs?: IDataObject,
	headers?: IDataObject,
	url?: string,
	itemIndex: number = 0,
): Promise<IDataObject[]> {
	const credentials = await this.getCredentials('microsoft365AdminServicePrincipalApi');
	const baseUrl = (
		typeof credentials.graphApiBaseUrl === 'string' && credentials.graphApiBaseUrl !== ''
			? credentials.graphApiBaseUrl
			: 'https://graph.microsoft.com'
	).replace(/\/+$/, '');
	// `IHttpRequestOptions` has no `uri` property, which `requestWithAuthenticationPaginated`
	// requires — so the deprecated option type is the only one that fits here.
	// eslint-disable-next-line @n8n/community-nodes/no-deprecated-workflow-functions
	const options: IRequestOptions = {
		method,
		uri: url ?? `${baseUrl}/v1.0${endpoint}`,
		json: true,
		headers,
		body,
		qs,
	};

	const pages = await this.helpers.requestWithAuthenticationPaginated.call(
		this,
		options,
		itemIndex,
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

export async function handleErrorPostReceive(
	this: IExecuteSingleFunctions,
	data: INodeExecutionData[],
	response: IN8nHttpFullResponse,
): Promise<INodeExecutionData[]> {
	if (String(response.statusCode).startsWith('4') || String(response.statusCode).startsWith('5')) {
		const resource = this.getNodeParameter('resource') as string;
		const operation = this.getNodeParameter('operation') as string;
		const {
			code: errorCode,
			message: errorMessage,
			details: errorDetails,
		} = (response.body as IDataObject)?.error as {
			code: string;
			message: string;
			innerError?: {
				code: string;
				'request-id'?: string;
				date?: string;
			};
			details?: Array<{
				code: string;
				message: string;
			}>;
		};

		// Operation specific errors
		if (resource === 'group') {
			if (operation === 'delete') {
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: "The required group doesn't match any existing one",
						description: "Double-check the value in the parameter 'Group to Delete' and try again",
					});
				}
			} else if (operation === 'get') {
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: "The required group doesn't match any existing one",
						description: "Double-check the value in the parameter 'Group to Get' and try again",
					});
				}
			} else if (operation === 'update') {
				if (
					errorCode === 'BadRequest' &&
					errorMessage === 'Empty Payload. JSON content expected.'
				) {
					// Ignore empty payload error. Currently n8n deletes the empty body object from the request.
					return data;
				}
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: "The required group doesn't match any existing one",
						description: "Double-check the value in the parameter 'Group to Update' and try again",
					});
				}
			}
		} else if (resource === 'user') {
			if (operation === 'addGroup') {
				if (
					errorCode === 'Request_BadRequest' &&
					errorMessage ===
						"One or more added object references already exist for the following modified properties: 'members'."
				) {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: 'The user is already in the group',
						description:
							'The specified user cannot be added to the group because they are already a member',
					});
				} else if (errorCode === 'Request_ResourceNotFound') {
					const group = this.getNodeParameter('group.value') as string;
					if (errorMessage.includes(group)) {
						throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
							message: "The required group doesn't match any existing one",
							description: "Double-check the value in the parameter 'Group' and try again",
						});
					} else {
						throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
							message: "The required user doesn't match any existing one",
							description: "Double-check the value in the parameter 'User to Add' and try again",
						});
					}
				}
			} else if (operation === 'delete') {
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: "The required user doesn't match any existing one",
						description: "Double-check the value in the parameter 'User to Delete' and try again",
					});
				}
			} else if (operation === 'get') {
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: "The required user doesn't match any existing one",
						description: "Double-check the value in the parameter 'User to Get' and try again",
					});
				}
			} else if (operation === 'removeGroup') {
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: 'The user is not in the group',
						description:
							'The specified user cannot be removed from the group because they are not a member of the group',
					});
				} else if (
					errorCode === 'Request_UnsupportedQuery' &&
					errorMessage ===
						"Unsupported referenced-object resource identifier for link property 'members'."
				) {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: 'The user ID is invalid',
						description: 'The ID should be in the format e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
					});
				}
			} else if (operation === 'update') {
				if (
					errorCode === 'BadRequest' &&
					errorMessage === 'Empty Payload. JSON content expected.'
				) {
					// Ignore empty payload error. Currently n8n deletes the empty body object from the request.
					return data;
				}
				if (errorCode === 'Request_ResourceNotFound') {
					throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
						message: "The required user doesn't match any existing one",
						description: "Double-check the value in the parameter 'User to Update' and try again",
					});
				}
			}
		}

		// Generic errors
		if (
			errorCode === 'Request_BadRequest' &&
			errorMessage.startsWith('Invalid object identifier')
		) {
			const group = this.getNodeParameter('group.value', '') as string;
			const parameterResource =
				resource === 'group' || errorMessage.includes(group) ? 'group' : 'user';

			throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
				message: `The ${parameterResource} ID is invalid`,
				description: 'The ID should be in the format e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
			});
		}
		if (errorDetails?.some((x) => x.code === 'ObjectConflict' || x.code === 'ConflictingObjects')) {
			throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
				message: `The ${resource} already exists`,
				description: errorMessage,
			});
		}

		throw new NodeApiError(this.getNode(), response as unknown as JsonObject);
	}

	return data;
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
			undefined,
			undefined,
			paginationToken,
		)) as DirectoryListResponse;
	} else {
		const qs: IDataObject = {
			$select: 'id,displayName',
		};
		const headers: IDataObject = {};
		if (filter) {
			headers.ConsistencyLevel = 'eventual';
			qs.$search = `"displayName:${filter}"`;
		}
		response = (await microsoftApiRequest.call(
			this,
			'GET',
			'/groups',
			{},
			qs,
			headers,
		)) as DirectoryListResponse;
	}

	const groups: Array<{
		id: string;
		displayName: string;
	}> = response.value ?? [];

	const results: INodeListSearchItems[] = groups
		.map((g) => ({
			name: g.displayName,
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
			undefined,
			undefined,
			paginationToken,
		)) as DirectoryListResponse;
	} else {
		const qs: IDataObject = {
			$select: 'id,displayName',
		};
		const headers: IDataObject = {};
		if (filter) {
			qs.$filter = `startsWith(displayName, '${filter}') OR startsWith(userPrincipalName, '${filter}')`;
		}
		response = (await microsoftApiRequest.call(
			this,
			'GET',
			'/users',
			{},
			qs,
			headers,
		)) as DirectoryListResponse;
	}

	const users: Array<{
		id: string;
		displayName: string;
	}> = response.value ?? [];

	const results: INodeListSearchItems[] = users
		.map((u) => ({
			name: u.displayName,
			value: u.id,
		}))
		.sort((a, b) =>
			a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }),
		);

	return { results, paginationToken: response['@odata.nextLink'] };
}
