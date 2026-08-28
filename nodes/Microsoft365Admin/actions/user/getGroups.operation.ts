import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { paginationProperties, userRLC } from '../../helpers/descriptions';
import { fetchCollection } from '../../helpers/pagination';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user whose group memberships should be retrieved'),
	...paginationProperties(),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['getGroups'],
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

	const groups = await fetchCollection.call(this, `/users/${user}/memberOf`, index);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(groups), {
		itemData: { item: index },
	});
}
