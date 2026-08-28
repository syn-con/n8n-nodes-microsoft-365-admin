import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { groupRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [groupRLC('Group to Delete')];

const displayOptions = {
	show: {
		resource: ['group'],
		operation: ['delete'],
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

	await microsoftApiRequest.call(this, 'DELETE', `/groups/${group}`, {}, { itemIndex: index });

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ deleted: true }), {
		itemData: { item: index },
	});
}
