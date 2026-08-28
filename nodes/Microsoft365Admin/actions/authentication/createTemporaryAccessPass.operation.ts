import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user to act on'),
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		options: [
			{
				displayName: 'One-Time Use',
				name: 'isUsableOnce',
				default: false,
				description:
					'Whether the pass stops working after a single sign-in. A multi-use pass is only accepted if the Temporary Access Pass policy allows it.',
				type: 'boolean',
			},
			{
				displayName: 'Lifetime (Minutes)',
				name: 'lifetimeInMinutes',
				default: 60,
				description:
					'How long the pass stays valid, between 10 minutes and 43200 (30 days). Leave the option out to use the tenant policy default.',
				type: 'number',
				typeOptions: {
					minValue: 10,
					maxValue: 43200,
				},
				validateType: 'number',
			},
			{
				displayName: 'Start Time',
				name: 'startDateTime',
				default: '',
				description:
					'When the pass becomes usable. Leave the option out to make it usable immediately.',
				type: 'dateTime',
			},
		],
		placeholder: 'Add option',
		type: 'collection',
	},
];

const displayOptions = {
	show: {
		resource: ['authentication'],
		operation: ['createTemporaryAccessPass'],
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

	// Graph requires a JSON representation even though every property is optional, so an
	// untouched Options collection still has to go out as `{}` rather than as no payload.
	const body = this.getNodeParameter('options', index, {}) as IDataObject;

	const pass = (await microsoftApiRequest.call(
		this,
		'POST',
		`/users/${user}/authentication/temporaryAccessPassMethods`,
		body,
		{ itemIndex: index },
	)) as IDataObject;

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(pass), {
		itemData: { item: index },
	});
}
