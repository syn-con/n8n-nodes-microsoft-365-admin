import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { groupRLC, userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { getGraphApiBaseUrl, microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	groupRLC('Group', 'The group to add the user to'),
	userRLC('User to Add', 'The user to add to the group'),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['addGroup'],
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

	// A `$ref` write identifies the object by absolute URL, which has to sit on the same
	// cloud the credential points at.
	const baseUrl = await getGraphApiBaseUrl.call(this);

	await microsoftApiRequest.call(
		this,
		'POST',
		`/groups/${group}/members/$ref`,
		{ '@odata.id': `${baseUrl}/v1.0/directoryObjects/${user}` },
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ added: true }), {
		itemData: { item: index },
	});
}
