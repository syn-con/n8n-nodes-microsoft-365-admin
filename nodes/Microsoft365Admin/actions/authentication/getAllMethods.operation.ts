import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { annotateMethod } from '../../helpers/authentication';
import { userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user whose authentication methods should be retrieved'),
];

const displayOptions = {
	show: {
		resource: ['authentication'],
		operation: ['getAllMethods'],
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
		`/users/${user}/authentication/methods`,
		{},
		{ itemIndex: index },
	)) as { value?: IDataObject[] };

	// Graph never reports the collection segment a method lives under, which Delete Method
	// needs, so it is derived from `@odata.type` here.
	const methods = (response.value ?? []).map(annotateMethod);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(methods), {
		itemData: { item: index },
	});
}
