import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { outputProperties, paginationProperties } from '../../helpers/descriptions';
import { buildSelect, fetchCollection } from '../../helpers/pagination';
import { updateDisplayOptions } from '../../helpers/utils';
import { GROUP_SIMPLE_SELECT } from './constants';

export const properties: INodeProperties[] = [
	...paginationProperties(),
	{
		displayName: 'Filter',
		name: 'filter',
		default: '',
		description:
			'<a href="https://docs.microsoft.com/en-us/graph/query-parameters#filter-parameter">Query parameter</a> to filter results by',
		hint: 'If empty, all the groups will be returned',
		placeholder: "e.g. startswith(displayName, 'a')",
		type: 'string',
		validateType: 'string',
	},
	...outputProperties('getGroupPropertiesGetAll'),
];

const displayOptions = {
	show: {
		resource: ['group'],
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
	const select = buildSelect.call(this, index, GROUP_SIMPLE_SELECT);
	if (select) {
		qs.$select = select;
	}
	if (filter) {
		qs.$filter = filter;
	}

	const groups = await fetchCollection.call(this, '/groups', index, qs);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(groups), {
		itemData: { item: index },
	});
}
