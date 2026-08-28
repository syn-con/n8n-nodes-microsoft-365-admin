import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { groupRLC, outputProperties } from '../../helpers/descriptions';
import { buildSelect } from '../../helpers/pagination';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';
import { GROUP_MEMBER_EXPAND, GROUP_SIMPLE_SELECT } from './constants';

export const properties: INodeProperties[] = [
	groupRLC('Group to Get'),
	...outputProperties('getGroupProperties'),
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		options: [
			{
				displayName: 'Include Members',
				name: 'includeMembers',
				default: false,
				type: 'boolean',
				validateType: 'boolean',
			},
		],
		placeholder: 'Add Option',
		type: 'collection',
	},
];

const displayOptions = {
	show: {
		resource: ['group'],
		operation: ['get'],
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
	const options = this.getNodeParameter('options', index, {}) as { includeMembers?: boolean };

	const qs: IDataObject = {};
	const select = buildSelect.call(this, index, GROUP_SIMPLE_SELECT);
	if (select) {
		qs.$select = select;
	}
	if (options.includeMembers) {
		qs.$expand = GROUP_MEMBER_EXPAND;
	}

	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		`/groups/${group}`,
		{},
		{ qs, itemIndex: index },
	)) as IDataObject;

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(response), {
		itemData: { item: index },
	});
}
