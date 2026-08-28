import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { outputProperties, userRLC } from '../../helpers/descriptions';
import { buildSelect } from '../../helpers/pagination';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';
import { USER_RAW_SELECT, USER_SIMPLE_SELECT } from './constants';

export const properties: INodeProperties[] = [
	userRLC('User to Get', 'The user to retrieve'),
	...outputProperties('getUserProperties'),
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['get'],
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

	const qs: IDataObject = {};
	const select = buildSelect.call(this, index, USER_SIMPLE_SELECT, USER_RAW_SELECT);
	if (select) {
		qs.$select = select;
	}

	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		`/users/${user}`,
		{},
		{ qs, itemIndex: index },
	)) as IDataObject;

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(response), {
		itemData: { item: index },
	});
}
