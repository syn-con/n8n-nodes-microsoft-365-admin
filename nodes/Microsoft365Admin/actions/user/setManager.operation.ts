import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { getGraphApiBaseUrl, microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user whose manager should be set'),
	userRLC('Manager', 'The user to set as manager', 'manager'),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['setManager'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const node = this.getNode();
	const user = assertPathSafe(
		node,
		this.getNodeParameter('user', index, '', { extractValue: true }),
		'user',
		index,
	);
	const manager = assertPathSafe(
		node,
		this.getNodeParameter('manager', index, '', { extractValue: true }),
		'manager',
		index,
	);

	// A `$ref` write identifies the object by absolute URL, which has to sit on the same
	// cloud the credential points at.
	const baseUrl = await getGraphApiBaseUrl.call(this);

	await microsoftApiRequest.call(
		this,
		'PUT',
		`/users/${user}/manager/$ref`,
		{ '@odata.id': `${baseUrl}/v1.0/users/${manager}` },
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ updated: true }), {
		itemData: { item: index },
	});
}
