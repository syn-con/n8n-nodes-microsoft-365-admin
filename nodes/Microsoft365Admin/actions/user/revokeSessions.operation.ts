import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user whose sessions should be revoked'),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['revokeSessions'],
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

	await microsoftApiRequest.call(
		this,
		'POST',
		`/users/${user}/revokeSignInSessions`,
		{},
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ revoked: true }), {
		itemData: { item: index },
	});
}
