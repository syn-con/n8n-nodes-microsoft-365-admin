import type { INodeProperties } from 'n8n-workflow';

import * as addGroup from './addGroup.operation';
import * as create from './create.operation';
import * as del from './delete.operation';
import * as get from './get.operation';
import * as getAll from './getAll.operation';
import * as getGroups from './getGroups.operation';
import * as getManager from './getManager.operation';
import * as removeGroup from './removeGroup.operation';
import * as revokeSessions from './revokeSessions.operation';
import * as setManager from './setManager.operation';
import * as update from './update.operation';

export {
	addGroup,
	create,
	del as delete,
	get,
	getAll,
	getGroups,
	getManager,
	removeGroup,
	revokeSessions,
	setManager,
	update,
};

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['user'],
			},
		},
		options: [
			{
				name: 'Add to Group',
				value: 'addGroup',
				description: 'Add user to group',
				action: 'Add user to group',
			},
			{
				name: 'Create',
				value: 'create',
				description: 'Create a user',
				action: 'Create user',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a user',
				action: 'Delete user',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Retrieve data for a specific user',
				action: 'Get user',
			},
			{
				name: 'Get Groups',
				value: 'getGroups',
				description: 'Retrieve the groups a user is a member of',
				action: 'Get groups for user',
			},
			{
				name: 'Get Manager',
				value: 'getManager',
				description: "Retrieve a user's manager",
				action: 'Get manager for user',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Retrieve a list of users',
				action: 'Get many users',
			},
			{
				name: 'Remove From Group',
				value: 'removeGroup',
				description: 'Remove user from group',
				action: 'Remove user from group',
			},
			{
				name: 'Revoke Sessions',
				value: 'revokeSessions',
				description: 'Invalidate all refresh and session tokens for a user',
				action: 'Revoke sessions for user',
			},
			{
				name: 'Set Manager',
				value: 'setManager',
				description: "Set a user's manager",
				action: 'Set manager for user',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a user',
				action: 'Update user',
			},
		],
		default: 'getAll',
	},
	...addGroup.description,
	...create.description,
	...del.description,
	...get.description,
	...getAll.description,
	...getGroups.description,
	...getManager.description,
	...removeGroup.description,
	...revokeSessions.description,
	...setManager.description,
	...update.description,
];
