import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { groupRLC, userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	groupRLC('Group', 'The group to remove the user from'),
	userRLC('User to Remove', 'The user to remove from the group'),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['removeGroup'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const node = this.getNode();
	const group = assertPathSafe(
		node,
		this.getNodeParameter('group', index, '', { extractValue: true }),
		'group',
		index,
	);
	const user = assertPathSafe(
		node,
		this.getNodeParameter('user', index, '', { extractValue: true }),
		'user',
		index,
	);

	await microsoftApiRequest.call(
		this,
		'DELETE',
		`/groups/${group}/members/${user}/$ref`,
		{},
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ removed: true }), {
		itemData: { item: index },
	});
}
