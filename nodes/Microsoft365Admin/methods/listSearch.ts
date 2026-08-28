import {
	NodeOperationError,
	type IDataObject,
	type ILoadOptionsFunctions,
	type INodeListSearchItems,
	type INodeListSearchResult,
} from 'n8n-workflow';

import { describeMethod, METHOD_TYPE_LABELS, methodTypeOf } from '../helpers/authentication';
import type { DirectoryListResponse } from '../helpers/interfaces';
import { isFilled, isPathSafe } from '../helpers/utils';
import { microsoftApiRequest } from '../transport';

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

/** Turns a directory collection into picker entries, sorted by the name shown. */
function toSearchItems(response: DirectoryListResponse): INodeListSearchResult {
	const results: INodeListSearchItems[] = (response.value ?? [])
		.map((entry) => ({
			// Sorting compares names, so a directory object without one must not be null.
			name: entry.displayName ?? entry.id,
			value: entry.id,
		}))
		.sort((a, b) =>
			a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }),
		);

	return { results, paginationToken: response['@odata.nextLink'] };
}

export async function getGroups(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	if (paginationToken) {
		return toSearchItems(
			(await microsoftApiRequest.call(
				this,
				'GET',
				'/groups',
				{},
				{ url: paginationToken },
			)) as DirectoryListResponse,
		);
	}

	const qs: IDataObject = { $select: 'id,displayName' };
	const headers: IDataObject = {};
	if (filter) {
		headers.ConsistencyLevel = 'eventual';
		qs.$search = `"displayName:${searchPhrase(filter)}"`;
	}

	return toSearchItems(
		(await microsoftApiRequest.call(
			this,
			'GET',
			'/groups',
			{},
			{ qs, headers },
		)) as DirectoryListResponse,
	);
}

export async function getUsers(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	if (paginationToken) {
		return toSearchItems(
			(await microsoftApiRequest.call(
				this,
				'GET',
				'/users',
				{},
				{ url: paginationToken },
			)) as DirectoryListResponse,
		);
	}

	const qs: IDataObject = { $select: 'id,displayName' };
	if (filter) {
		const term = odataLiteral(filter);
		qs.$filter = `startsWith(displayName, '${term}') OR startsWith(userPrincipalName, '${term}')`;
	}

	return toSearchItems(
		(await microsoftApiRequest.call(
			this,
			'GET',
			'/users',
			{},
			{ qs, headers: {} },
		)) as DirectoryListResponse,
	);
}

/** Reads the user whose picker is being opened, as a path-safe segment. */
function currentUser(this: ILoadOptionsFunctions): string {
	const user = this.getCurrentNodeParameter('user', { extractValue: true });

	if (!isPathSafe(user)) {
		throw new NodeOperationError(this.getNode(), 'Choose a user first', {
			description: 'The methods on offer are the ones registered to that user',
		});
	}

	return user;
}

/**
 * Lists the methods of the selected type registered to the selected user.
 *
 * The type-specific endpoints are inconsistent when a user has none: most answer with an
 * empty collection, while platformCredentialMethods can answer 404. The aggregate methods
 * endpoint has stable empty semantics, so the picker reads it and filters by `@odata.type`
 * here. The collections are small and support no `$filter`, so the search term is also
 * matched locally.
 */
export async function getAuthenticationMethods(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const user = currentUser.call(this);
	const methodType = this.getCurrentNodeParameter('methodType');

	if (!isFilled(methodType) || METHOD_TYPE_LABELS[methodType] === undefined) {
		throw new NodeOperationError(this.getNode(), 'Choose a method type first');
	}

	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		`/users/${user}/authentication/methods`,
	)) as { value?: IDataObject[] };

	const term = filter?.toLowerCase();
	const results: INodeListSearchItems[] = (response.value ?? [])
		.filter((method) => methodTypeOf(method['@odata.type']) === methodType)
		.map((method) => ({ name: describeMethod(method, methodType), value: String(method.id) }))
		.filter((item) => !term || item.name.toLowerCase().includes(term))
		.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

	return { results };
}
