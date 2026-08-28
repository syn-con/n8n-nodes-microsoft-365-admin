import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { outputProperties, paginationProperties } from '../../helpers/descriptions';
import { buildSelect, fetchCollection } from '../../helpers/pagination';
import { updateDisplayOptions } from '../../helpers/utils';
import { USER_RAW_SELECT, USER_SIMPLE_SELECT } from './constants';

export const properties: INodeProperties[] = [
	...paginationProperties(),
	{
		displayName: 'Filter',
		name: 'filter',
		default: '',
		description:
			'<a href="https://docs.microsoft.com/en-us/graph/query-parameters#filter-parameter">Query parameter</a> to filter results by',
		hint: 'If empty, all the users will be returned',
		placeholder: "e.g. startswith(displayName, 'a')",
		type: 'string',
		validateType: 'string',
	},
	...outputProperties('getUserPropertiesGetAll'),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['getAll'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const filter = this.getNodeParameter('filter', index, '') as string;

	const qs: IDataObject = {};
	const select = buildSelect.call(this, index, USER_SIMPLE_SELECT, USER_RAW_SELECT);
	if (select) {
		qs.$select = select;
	}

	if (filter) {
		qs.$filter = filter;
	}

	const users = await fetchCollection.call(this, '/users', index, qs);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(users), {
		itemData: { item: index },
	});
}
