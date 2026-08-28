import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import {
	DEFAULT_PASSWORD_LENGTH,
	generatePassword,
	MAX_PASSWORD_LENGTH,
	MIN_PASSWORD_LENGTH,
} from '../../helpers/authentication';
import { userRLC } from '../../helpers/descriptions';
import { assertPathSafe, isFilled, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

interface ResetPasswordOptions {
	password?: string;
	passwordLength?: number;
	forceChangePassword?: string;
}

export const properties: INodeProperties[] = [
	{
		displayName:
			'Requires the User-PasswordProfile.ReadWrite.All application permission with admin consent, and the service principal must have at least the User Administrator Microsoft Entra role.',
		name: 'resetPasswordPermissionNotice',
		type: 'notice',
		default: '',
	},
	userRLC('User', 'The user to act on'),
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		options: [
			{
				displayName: 'Force Change',
				name: 'forceChangePassword',
				default: 'forceChangePasswordNextSignIn',
				description: 'Whether the user has to choose their own password, and when',
				options: [
					{
						name: 'At Next Sign-In',
						value: 'forceChangePasswordNextSignIn',
						description: 'The user is prompted to set a new password when they next sign in',
					},
					{
						name: 'At Next Sign-In, With MFA',
						value: 'forceChangePasswordNextSignInWithMfa',
						description: 'As above, but the user must pass MFA before changing it',
					},
					{
						name: 'Never',
						value: 'never',
						description: 'The password is set and the user is not asked to change it',
					},
				],
				type: 'options',
			},
			{
				displayName: 'Password',
				name: 'password',
				default: '',
				description:
					'The password to set. Leave empty to have the node generate one and return it on the output item.',
				type: 'string',
				typeOptions: {
					password: true,
				},
			},
			{
				displayName: 'Password Length',
				name: 'passwordLength',
				default: 16,
				description: 'Length of the generated password. Ignored when a password is given.',
				type: 'number',
				typeOptions: {
					minValue: 8,
					maxValue: 256,
				},
				validateType: 'number',
			},
		],
		placeholder: 'Add option',
		type: 'collection',
	},
];

const displayOptions = {
	show: {
		resource: ['authentication'],
		operation: ['resetPassword'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

/**
 * Sets a new password on the user through `passwordProfile`.
 *
 * Graph does have a dedicated password reset —
 * `POST /users/{id}/authentication/methods/{id}/resetPassword` — but it supports delegated
 * access only, so an app-only credential like this node's can never call it. Updating
 * **passwordProfile** is the app-only equivalent: it writes to Entra ID and, where password
 * writeback is configured, on to on-premises AD.
 *
 * The router runs operations one item at a time, which is what a password reset needs:
 * it is a directory write, and Entra rejects writes that overlap within a tenant.
 */
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

	const options = this.getNodeParameter('options', index, {}) as ResetPasswordOptions;
	const length = Math.min(
		MAX_PASSWORD_LENGTH,
		Math.max(MIN_PASSWORD_LENGTH, options.passwordLength ?? DEFAULT_PASSWORD_LENGTH),
	);
	const supplied = isFilled(options.password) ? options.password.trim() : undefined;
	const password = supplied ?? generatePassword(length);

	const forceChange = options.forceChangePassword ?? 'forceChangePasswordNextSignIn';
	const passwordProfile: IDataObject = { password };
	if (forceChange === 'forceChangePasswordNextSignInWithMfa') {
		passwordProfile.forceChangePasswordNextSignInWithMfa = true;
	} else {
		passwordProfile.forceChangePasswordNextSignIn = forceChange === 'forceChangePasswordNextSignIn';
	}

	await microsoftApiRequest.call(
		this,
		'PATCH',
		`/users/${user}`,
		{ passwordProfile },
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(
		this.helpers.returnJsonArray({
			id: user,
			passwordReset: true,
			// Carried out so the workflow can deliver it: a generated password exists
			// nowhere else, and Graph never reads one back.
			password,
			generated: supplied === undefined,
			forceChangePasswordNextSignIn: passwordProfile.forceChangePasswordNextSignIn === true,
			forceChangePasswordNextSignInWithMfa:
				passwordProfile.forceChangePasswordNextSignInWithMfa === true,
		}),
		{ itemData: { item: index } },
	);
}
