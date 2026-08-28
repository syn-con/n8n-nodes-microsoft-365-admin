import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { groupRLC, userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { getGraphApiBaseUrl, microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	groupRLC('Group', 'The group to add the owner to'),
	userRLC('User to Add as Owner', 'The user to make an owner of the group'),
];

const displayOptions = {
	show: {
		resource: ['group'],
		operation: ['addOwner'],
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
		`/groups/${group}/owners/$ref`,
		{ '@odata.id': `${baseUrl}/v1.0/directoryObjects/${user}` },
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ added: true }), {
		itemData: { item: index },
	});
}
