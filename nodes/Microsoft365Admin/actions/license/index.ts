import type { INodeProperties } from 'n8n-workflow';

import * as assign from './assign.operation';
import * as assignGroup from './assignGroup.operation';
import * as queryHolders from './queryHolders.operation';
import * as queryTenant from './queryTenant.operation';
import * as queryUser from './queryUser.operation';
import * as unassign from './unassign.operation';

export { assign, assignGroup, queryHolders, queryTenant, queryUser, unassign };

export const description: INodeProperties[] = [
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
				action: 'Query license holders',
			},
			{
				name: 'Query Tenant Licenses',
				value: 'queryTenant',
				description: 'Retrieve the license SKUs the tenant has subscribed to',
				action: 'Query tenant licenses',
			},
			{
				name: 'Query User Licenses',
				value: 'queryUser',
				description: 'Retrieve the licenses currently assigned to a user',
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
	...assign.description,
	...assignGroup.description,
	...queryHolders.description,
	...queryTenant.description,
	...queryUser.description,
	...unassign.description,
];
