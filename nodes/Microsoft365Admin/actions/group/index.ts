import type { INodeProperties } from 'n8n-workflow';

import * as addOwner from './addOwner.operation';
import * as create from './create.operation';
import * as del from './delete.operation';
import * as get from './get.operation';
import * as getAll from './getAll.operation';
import * as getMembers from './getMembers.operation';
import * as getOwners from './getOwners.operation';
import * as removeOwner from './removeOwner.operation';
import * as update from './update.operation';

export { addOwner, create, del as delete, get, getAll, getMembers, getOwners, removeOwner, update };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['group'],
			},
		},
		options: [
			{
				name: 'Add Owner',
				value: 'addOwner',
				description: 'Add an owner to a group',
				action: 'Add owner to group',
			},
			{
				name: 'Create',
				value: 'create',
				description: 'Create a group',
				action: 'Create group',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a group',
				action: 'Delete group',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Retrieve data for a specific group',
				action: 'Get group',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Retrieve a list of groups',
				action: 'Get many groups',
			},
			{
				name: 'Get Members',
				value: 'getMembers',
				description: 'Retrieve the members of a group',
				action: 'Get members of group',
			},
			{
				name: 'Get Owners',
				value: 'getOwners',
				description: 'Retrieve the owners of a group',
				action: 'Get owners of group',
			},
			{
				name: 'Remove Owner',
				value: 'removeOwner',
				description: 'Remove an owner from a group',
				action: 'Remove owner from group',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a group',
				action: 'Update group',
			},
		],
		default: 'getAll',
	},
	...addOwner.description,
	...create.description,
	...del.description,
	...get.description,
	...getAll.description,
	...getMembers.description,
	...getOwners.description,
	...removeOwner.description,
	...update.description,
];
