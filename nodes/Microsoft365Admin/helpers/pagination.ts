import type { IDataObject, IExecuteFunctions } from 'n8n-workflow';

import { microsoftApiPaginateRequest, microsoftApiRequest } from '../transport';

/**
 * Reads a Graph collection, honouring the operation's Return All / Limit parameters.
 *
 * Declarative routing expressed this as `send.paginate` plus a `$top` mapping; doing it
 * here keeps the two halves of the decision — walk every page, or ask for one page of
 * `limit` — in one place instead of split across two parameter definitions.
 */
export async function fetchCollection(
	this: IExecuteFunctions,
	endpoint: string,
	index: number,
	qs: IDataObject = {},
	headers?: IDataObject,
): Promise<IDataObject[]> {
	const returnAll = this.getNodeParameter('returnAll', index, false) as boolean;

	if (returnAll) {
		return microsoftApiPaginateRequest.call(
			this,
			'GET',
			endpoint,
			{},
			{ qs, headers, itemIndex: index },
		);
	}

	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		endpoint,
		{},
		{
			qs: { ...qs, $top: this.getNodeParameter('limit', index, 50) as number },
			headers,
			itemIndex: index,
		},
	)) as { value?: IDataObject[] };

	return response.value ?? [];
}

/**
 * Builds the `$select` projection from the Output / Field Names parameters.
 *
 * `rawSelect` is what "Raw" projects. Groups leave it unset, which lets Graph apply its
 * own default; users name every readable property explicitly, because Graph's default
 * projection for a user is far narrower than what the operation is expected to return.
 */
export function buildSelect(
	this: IExecuteFunctions,
	index: number,
	simpleSelect: string,
	rawSelect?: string,
): string | undefined {
	const output = this.getNodeParameter('output', index, 'simple') as string;

	if (output === 'simple') {
		return simpleSelect;
	}

	if (output === 'fields') {
		const fields = this.getNodeParameter('fields', index, []) as string[];
		// `id` is always projected: without it the output cannot be piped into an
		// operation that addresses the object.
		return [...new Set([...fields, 'id'])].join(',');
	}

	return rawSelect;
}
