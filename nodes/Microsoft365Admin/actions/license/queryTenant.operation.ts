import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [];

const displayOptions = {
	show: {
		resource: ['license'],
		operation: ['queryTenant'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	// `/subscribedSkus` returns the whole collection in one response — it supports no
	// paging, so there is nothing to walk here.
	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		'/subscribedSkus',
		{},
		{
			itemIndex: index,
		},
	)) as { value?: IDataObject[] };

	return this.helpers.constructExecutionMetaData(
		this.helpers.returnJsonArray(response.value ?? []),
		{ itemData: { item: index } },
	);
}
