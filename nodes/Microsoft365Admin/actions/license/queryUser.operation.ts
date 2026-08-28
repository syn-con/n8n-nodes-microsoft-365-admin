import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user whose licenses should be retrieved'),
];

const displayOptions = {
	show: {
		resource: ['license'],
		operation: ['queryUser'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const user = assertPathSafe(
		this.getNode(),
		this.getNodeParameter('user', index, '', { extractValue: true }),
		'user',
		index,
	);

	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		`/users/${user}/licenseDetails`,
		{},
		{ itemIndex: index },
	)) as { value?: IDataObject[] };

	return this.helpers.constructExecutionMetaData(
		this.helpers.returnJsonArray(response.value ?? []),
		{ itemData: { item: index } },
	);
}
