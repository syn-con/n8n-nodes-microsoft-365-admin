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
				description: 'Assign a license to a user',
				routing: {
					request: {
						method: 'POST',
						url: '=/users/{{ $parameter["user"] }}/assignLicense',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
				action: 'Assign license to user',
			},
			{
				name: 'Assign to Group',
				value: 'assignGroup',
				description: 'Assign a license to a group, licensing every member of it',
				routing: {
					request: {
						method: 'POST',
						url: '=/groups/{{ $parameter["group"] }}/assignLicense',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
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
				description: 'Remove a license from a user, freeing the seat',
				routing: {
					request: {
						method: 'POST',
						url: '=/users/{{ $parameter["user"] }}/assignLicense',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
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
		displayName: 'License SKU Name or ID',
		name: 'skuId',
		default: '',
		description:
			'The license SKU to assign. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assign'],
			},
		},
		required: true,
		routing: {
			send: {
				property: 'addLicenses',
				type: 'body',
				// Graph expects an array of license assignments. `disabledPlans` is always
				// present (empty means "enable every service plan in the SKU").
				value:
					'={{ [{ skuId: $value, disabledPlans: ($parameter.options?.disabledPlans || "").split(",").map(plan => plan.trim()).filter(plan => plan.length > 0) }] }}',
			},
		},
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
	{
		// Graph rejects the request unless `removeLicenses` is present, even when empty.
		displayName: 'Remove Licenses',
		name: 'removeLicenses',
		default: '',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assign'],
			},
		},
		routing: {
			send: {
				property: 'removeLicenses',
				type: 'body',
				value: '={{ [] }}',
			},
		},
		type: 'hidden',
	},
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assign'],
			},
		},
		options: [
			{
				displayName: 'Disabled Plans',
				name: 'disabledPlans',
				default: '',
				description:
					'Comma-separated service plan IDs to leave switched off within the assigned SKU. Leave empty to enable every plan.',
				placeholder: 'e.g. 8c7d2df8-86f0-4902-b2ed-a0458298f3b3',
				type: 'string',
			},
		],
		placeholder: 'Add option',
		type: 'collection',
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
		displayName: 'License SKU Name or ID',
		name: 'skuId',
		default: '',
		description:
			'The license SKU to assign. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assignGroup'],
			},
		},
		required: true,
		routing: {
			send: {
				property: 'addLicenses',
				type: 'body',
				value: '={{ [{ skuId: $value, disabledPlans: [] }] }}',
			},
		},
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
	{
		displayName: 'Remove Licenses',
		name: 'removeLicenses',
		default: '',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['assignGroup'],
			},
		},
		routing: {
			send: {
				property: 'removeLicenses',
				type: 'body',
				value: '={{ [] }}',
			},
		},
		type: 'hidden',
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
		displayName: 'License SKU Name or ID',
		name: 'skuId',
		default: '',
		description:
			'The license SKU to remove. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['unassign'],
			},
		},
		required: true,
		routing: {
			send: {
				property: 'removeLicenses',
				type: 'body',
				value: '={{ [$value] }}',
			},
		},
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	},
	{
		// Graph requires both keys on every assignLicense call.
		displayName: 'Add Licenses',
		name: 'addLicenses',
		default: '',
		displayOptions: {
			show: {
				resource: ['license'],
				operation: ['unassign'],
			},
		},
		routing: {
			send: {
				property: 'addLicenses',
				type: 'body',
				value: '={{ [] }}',
			},
		},
		type: 'hidden',
	},
];

export const licenseFields: INodeProperties[] = [
	...assignFields,
	...assignGroupFields,
	...queryHoldersFields,
	...queryUserFields,
	...unassignFields,
];
