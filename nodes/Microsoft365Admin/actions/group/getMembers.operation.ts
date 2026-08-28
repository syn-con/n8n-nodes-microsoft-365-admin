import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { groupRLC, paginationProperties } from '../../helpers/descriptions';
import { fetchCollection } from '../../helpers/pagination';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';

export const properties: INodeProperties[] = [
	groupRLC('Group', 'The group whose members should be retrieved'),
	...paginationProperties(),
];

const displayOptions = {
	show: {
		resource: ['group'],
		operation: ['getMembers'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const group = assertPathSafe(
		this.getNode(),
		this.getNodeParameter('group', index, '', { extractValue: true }),
		'group',
		index,
	);

	const members = await fetchCollection.call(this, `/groups/${group}/members`, index);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(members), {
		itemData: { item: index },
	});
}
