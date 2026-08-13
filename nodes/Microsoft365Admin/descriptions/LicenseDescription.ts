import type { INodeProperties } from 'n8n-workflow';

import { ignoreHttpStatusErrorsConfig } from './common';
import { handleErrorPostReceive } from '../GenericFunctions';

export const licenseOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['license'],
			},
		},
		options: [
			{
				name: 'Assign',
				value: 'assign',
				description: 'Assign one or more licenses to a user',
				// No `routing`: the write operations run through `customOperations` in the node
				// class, which serializes the requests Entra ID refuses to process in parallel.
				action: 'Assign license to user',
			},
			{
				name: 'Assign to Group',
				value: 'assignGroup',
				description: 'Assign one or more licenses to a group, licensing every member of it',
				action: 'Assign license to group',
			},
			{
				name: 'Query License Holders',
				value: 'queryHolders',
				description: 'List the users or groups that a given license is assigned to',
				routing: {
					request: {
						method: 'GET',
						url: '=/{{ $parameter["holderType"] }}',
						// Filtering on `assignedLicenses` is a Graph "advanced query": it only
						// works with $count and an eventual-consistency header. Without both,
						// Graph rejects the request rather than silently ignoring the filter.
						headers: {
							ConsistencyLevel: 'eventual',
						},
						qs: {
							$count: 'true',
						},
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'rootProperty',
								properties: {
									property: 'value',
								},
							},
						],
					},
				},
				action: 'Query license holders',
			},
			{
				name: 'Query Tenant Licenses',
				value: 'queryTenant',
				description: 'Retrieve the license SKUs the tenant has subscribed to',
				routing: {
					request: {
						method: 'GET',
						url: '/subscribedSkus',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'rootProperty',
								properties: {
									property: 'value',
								},
							},
						],
					},
				},
				action: 'Query tenant licenses',
			},
			{
				name: 'Query User Licenses',
				value: 'queryUser',
				description: 'Retrieve the licenses currently assigned to a user',
				routing: {
					request: {
						method: 'GET',
						url: '=/users/{{ $parameter["user"] }}/licenseDetails',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'rootProperty',
								properties: {
									property: 'value',
								},
							},
						],
					},
				},
				action: 'Query user licenses',
			},
			{
				name: 'Unassign',
				value: 'unassign',
				description: 'Remove one or more licenses from a user, freeing the seats',
				action: 'Unassign license from user',
			},
		],
		default: 'queryTenant',
	},
];

const assignFields: INodeProperties[] = [
	{
		displayName: 'User',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		description: 'The user to assign the license to',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assign'],
			},
		},
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getUsers',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				placeholder: 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	},
	{
		// A multi-select, because Graph applies every SKU in one `assignLicense` call —
		// and Entra ID spends the same tenant-wide processing time whether that call
		// carries one license or ten.
		displayName: 'License SKU Names or IDs',
		name: 'skuId',
		default: [],
		description:
			'The licenses to assign. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assign'],
			},
		},
		required: true,
		type: 'multiOptions',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
];

const queryUserFields: INodeProperties[] = [
	{
		displayName: 'User',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		description: 'The user whose licenses should be retrieved',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['queryUser'],
			},
		},
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getUsers',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				placeholder: 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	},
];

const assignGroupFields: INodeProperties[] = [
	{
		displayName: 'Group',
		name: 'group',
		default: {
			mode: 'list',
			value: '',
		},
		description: 'The group to assign the license to',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assignGroup'],
			},
		},
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getGroups',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				placeholder: 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	},
	{
		displayName: 'License SKU Names or IDs',
		name: 'skuId',
		default: [],
		description:
			'The licenses to assign. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assignGroup'],
			},
		},
		required: true,
		type: 'multiOptions',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
];

const queryHoldersFields: INodeProperties[] = [
	{
		displayName: 'Holder Type',
		name: 'holderType',
		default: 'users',
		description: 'Whether to list users or groups holding the license',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['queryHolders'],
			},
		},
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
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['queryHolders'],
			},
		},
		required: true,
		routing: {
			send: {
				property: '$filter',
				type: 'query',
				// `skuId` is Edm.Guid, so the value is unquoted in the OData filter.
				value: '=assignedLicenses/any(license:license/skuId eq {{ $value }})',
			},
		},
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
	{
		// `assignedLicenses` is not part of the default projection for either collection,
		// so without an explicit $select the results say nothing about the licences held.
		displayName: 'Select',
		name: 'select',
		default: '',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['queryHolders'],
			},
		},
		routing: {
			send: {
				property: '$select',
				type: 'query',
				value:
					'={{ $parameter["holderType"] === "groups" ? "id,displayName,assignedLicenses,licenseProcessingState" : "id,displayName,userPrincipalName,assignedLicenses,licenseAssignmentStates" }}',
			},
		},
		type: 'hidden',
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['queryHolders'],
			},
		},
		routing: {
			send: {
				paginate: '={{ $value }}',
			},
			operations: {
				pagination: {
					type: 'generic',
					properties: {
						continue: '={{ !!$response.body?.["@odata.nextLink"] }}',
						request: {
							url: '={{ $response.body?.["@odata.nextLink"] ?? $request.url }}',
						},
					},
				},
			},
		},
		type: 'boolean',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		default: 50,
		description: 'Max number of results to return',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['queryHolders'],
				returnAll: [false],
			},
		},
		routing: {
			send: {
				property: '$top',
				type: 'query',
				value: '={{ $value }}',
			},
		},
		type: 'number',
		typeOptions: {
			minValue: 1,
		},
		validateType: 'number',
	},
];

const unassignFields: INodeProperties[] = [
	{
		displayName: 'User',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		description: 'The user to remove the license from',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['unassign'],
			},
		},
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getUsers',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				placeholder: 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	},
	{
		displayName: 'License SKU Names or IDs',
		name: 'skuId',
		default: [],
		description:
			'The licenses to remove. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['unassign'],
			},
		},
		required: true,
		type: 'multiOptions',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
];

/**
 * Shared by the three write operations, which all post to `assignLicense` and all have to
 * cope with Entra ID processing one license change per tenant at a time.
 */
const writeOptionsFields: INodeProperties[] = [
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assign', 'assignGroup', 'unassign'],
			},
		},
		options: [
			{
				displayName: 'Combine Items for the Same Target',
				name: 'combineItems',
				default: true,
				description:
					'Whether to merge every input item aimed at the same user or group into one Graph request. A single request can add and remove any number of licenses at no extra cost, so leaving this on is what makes a run over a long list of license changes finish in minutes rather than hours. Turn it off to send one request per item.',
				type: 'boolean',
			},
			{
				displayName: 'Disabled Plans',
				name: 'disabledPlans',
				default: '',
				description:
					'Comma-separated service plan IDs to leave switched off. Each plan is applied to whichever selected SKU contains it. Leave empty to enable every plan.',
				displayOptions: {
					show: {
						'/operation': ['assign', 'assignGroup'],
					},
				},
				placeholder: 'e.g. 8c7d2df8-86f0-4902-b2ed-a0458298f3b3',
				type: 'string',
			},
			{
				displayName: 'License SKU Names or IDs to Remove',
				name: 'removeSkuIds',
				default: [],
				description:
					'Licenses to remove in the same request that assigns the ones above — a swap costs one round of tenant processing instead of two. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
				displayOptions: {
					show: {
						'/operation': ['assign', 'assignGroup'],
					},
				},
				type: 'multiOptions',
				typeOptions: {
					loadOptionsMethod: 'getSubscribedSkus',
				},
			},
			{
				displayName: 'Max Retries',
				name: 'maxRetries',
				default: 5,
				description:
					'How many times to retry a write that Entra rejected because the tenant was busy with another license change. Each retry waits longer than the last, up to a minute.',
				type: 'number',
				typeOptions: {
					minValue: 0,
				},
				validateType: 'number',
			},
			{
				displayName: 'Wait Between Requests',
				name: 'waitBetweenRequests',
				default: 0,
				description:
					'Milliseconds to pause between requests. Only useful when something outside this workflow is also changing licenses in the tenant and the retries are not keeping up.',
				type: 'number',
				typeOptions: {
					minValue: 0,
				},
				validateType: 'number',
			},
		],
		placeholder: 'Add option',
		type: 'collection',
	},
];

export const licenseFields: INodeProperties[] = [
	...assignFields,
	...assignGroupFields,
	...queryHoldersFields,
	...queryUserFields,
	...unassignFields,
	...writeOptionsFields,
];
