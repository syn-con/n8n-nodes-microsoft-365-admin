import type { INodeProperties } from 'n8n-workflow';

import * as createTemporaryAccessPass from './createTemporaryAccessPass.operation';
import * as deleteMethod from './deleteMethod.operation';
import * as getAllMethods from './getAllMethods.operation';
import * as getPasswordMethod from './getPasswordMethod.operation';
import * as resetPassword from './resetPassword.operation';

export { createTemporaryAccessPass, deleteMethod, getAllMethods, getPasswordMethod, resetPassword };

export const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['authentication'],
			},
		},
		options: [
			{
				name: 'Create Temporary Access Pass',
				value: 'createTemporaryAccessPass',
				description:
					'Issue a time-limited passcode the user can sign in with, e.g. to register a new authenticator',
				action: 'Create temporary access pass for user',
			},
			{
				name: 'Delete Method',
				value: 'deleteMethod',
				description:
					'Remove one registered authentication method, e.g. a lost phone or security key',
				action: 'Delete authentication method',
			},
			{
				name: 'Get Many Methods',
				value: 'getAllMethods',
				description:
					'List every authentication method registered to a user, with the method type each one needs for Delete Method',
				action: 'Get many authentication methods',
			},
			{
				name: 'Get Password Method',
				value: 'getPasswordMethod',
				description: 'Retrieve the password method registered to a user, including its ID',
				action: 'Get password method',
			},
			{
				name: 'Reset Password',
				value: 'resetPassword',
				description: "Set a new password on a user's account, generating one if none is given",
				action: 'Reset user password',
			},
		],
		default: 'getAllMethods',
	},
	...createTemporaryAccessPass.description,
	...deleteMethod.description,
	...getAllMethods.description,
	...getPasswordMethod.description,
	...resetPassword.description,
];
