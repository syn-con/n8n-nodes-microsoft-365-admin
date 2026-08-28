import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiPaginateRequest, microsoftApiRequest } from '../../transport';

/**
 * `assignedLicenses` is not part of the default projection for either collection, so
 * without an explicit `$select` the results say nothing about the licences held.
 */
const SELECT_BY_HOLDER_TYPE: Record<string, string> = {
	groups: 'id,displayName,assignedLicenses,licenseProcessingState',
	users: 'id,displayName,userPrincipalName,assignedLicenses,licenseAssignmentStates',
};

export const properties: INodeProperties[] = [
	{
		displayName: 'Holder Type',
		name: 'holderType',
		default: 'users',
		description: 'Whether to list users or groups holding the license',
		options: [
			{
				name: 'Groups',
				value: 'groups',
				description: 'Groups the license is assigned to (group-based licensing)',
			},
			{
				name: 'Users',
				value: 'users',
				description: 'Users holding the license, directly or inherited from a group',
			},
		],
		type: 'options',
	},
	{
		displayName: 'License SKU Name or ID',
		name: 'skuId',
		default: '',
		description:
			'The license to search for. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		required: true,
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		type: 'boolean',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		default: 50,
		description: 'Max number of results to return',
		displayOptions: {
			show: {
				returnAll: [false],
			},
		},
		type: 'number',
		typeOptions: {
			minValue: 1,
		},
		validateType: 'number',
	},
];

const displayOptions = {
	show: {
		resource: ['license'],
		operation: ['queryHolders'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const holderType = this.getNodeParameter('holderType', index) as string;
	const skuId = this.getNodeParameter('skuId', index) as string;
	const returnAll = this.getNodeParameter('returnAll', index, false) as boolean;

	// `holderType` is a fixed enum, so it cannot reshape the path; guard anyway in case an
	// expression supplies something else.
	const endpoint = holderType === 'groups' ? '/groups' : '/users';

	const qs: IDataObject = {
		// Filtering on `assignedLicenses` is a Graph "advanced query": it only works with
		// $count and an eventual-consistency header. Without both, Graph rejects the
		// request rather than silently ignoring the filter.
		$count: 'true',
		// `skuId` is Edm.Guid, so the value is unquoted in the OData filter.
		$filter: `assignedLicenses/any(license:license/skuId eq ${skuId})`,
		$select: SELECT_BY_HOLDER_TYPE[holderType] ?? SELECT_BY_HOLDER_TYPE.users,
	};
	const headers: IDataObject = { ConsistencyLevel: 'eventual' };

	let holders: IDataObject[];

	if (returnAll) {
		holders = await microsoftApiPaginateRequest.call(
			this,
			'GET',
			endpoint,
			{},
			{ qs, headers, itemIndex: index },
		);
	} else {
		qs.$top = this.getNodeParameter('limit', index, 50) as number;
		const response = (await microsoftApiRequest.call(
			this,
			'GET',
			endpoint,
			{},
			{ qs, headers, itemIndex: index },
		)) as { value?: IDataObject[] };
		holders = response.value ?? [];
	}

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(holders), {
		itemData: { item: index },
	});
}
