import type { INodeProperties } from 'n8n-workflow';

import { ignoreHttpStatusErrorsConfig, userLocator } from './common';
import {
	annotateMethodsPostReceive,
	validatePathSegmentsPreSend,
} from '../AuthenticationFunctions';
import { handleErrorPostReceive } from '../GraphErrors';

/** Every write operation on this resource targets a single user. */
const WRITE_OPERATIONS = ['createTemporaryAccessPass', 'deleteMethod', 'resetPassword'];

/**
 * Method types that can be deleted, by the URL segment their collection lives under.
 *
 * Password is absent because Graph has no delete for it, and the beta-only types are
 * absent because this node talks to v1.0.
 */
const DELETABLE_METHOD_TYPES: Array<{ name: string; value: string; description: string }> = [
	{
		name: 'Email',
		value: 'emailMethods',
		description: 'The email address registered for self-service password reset',
	},
	{
		name: 'FIDO2 Security Key',
		value: 'fido2Methods',
		description: 'A passkey or security key',
	},
	{
		name: 'Microsoft Authenticator',
		value: 'microsoftAuthenticatorMethods',
		description: 'A device registered with the Microsoft Authenticator app',
	},
	{
		name: 'Phone',
		value: 'phoneMethods',
		description: 'A phone number used for SMS or voice call sign-in',
	},
	{
		name: 'Platform Credential',
		value: 'platformCredentialMethods',
		description: 'A platform credential registered on macOS',
	},
	{
		name: 'Software OATH Token',
		value: 'softwareOathMethods',
		description: 'A third-party authenticator app producing time-based codes',
	},
	{
		name: 'Temporary Access Pass',
		value: 'temporaryAccessPassMethods',
		description: 'A pass issued earlier, whether or not it has been used',
	},
	{
		name: 'Windows Hello for Business',
		value: 'windowsHelloForBusinessMethods',
		description: 'A Windows Hello registration on one device',
	},
];

export const authenticationOperations: INodeProperties[] = [
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
				routing: {
					request: {
						method: 'POST',
						url: '=/users/{{ $parameter["user"] }}/authentication/temporaryAccessPassMethods',
						// Graph requires a JSON representation even though every property is
						// optional. Without an explicit body n8n sends no payload when Options
						// is untouched, which Graph rejects instead of applying the policy defaults.
						body: '={{ $parameter["options"] || {} }}',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
				action: 'Create temporary access pass for user',
			},
			{
				name: 'Delete Method',
				value: 'deleteMethod',
				description:
					'Remove one registered authentication method, e.g. a lost phone or security key',
				routing: {
					request: {
						method: 'DELETE',
						// The collection segment is part of the address: an ID alone does not
						// identify a method. See AuthenticationFunctions.ts.
						url: '=/users/{{ $parameter["user"] }}/authentication/{{ $parameter["methodType"] }}/{{ $parameter["method"] }}',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'set',
								properties: {
									value: '={{ { "deleted": true } }}',
								},
							},
						],
					},
				},
				action: 'Delete authentication method',
			},
			{
				name: 'Get Many Methods',
				value: 'getAllMethods',
				description:
					'List every authentication method registered to a user, with the method type each one needs for Delete Method',
				routing: {
					request: {
						method: 'GET',
						url: '=/users/{{ $parameter["user"] }}/authentication/methods',
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
							annotateMethodsPostReceive,
						],
					},
				},
				action: 'Get many authentication methods',
			},
			{
				name: 'Get Password Method',
				value: 'getPasswordMethod',
				description: 'Retrieve the password method registered to a user, including its ID',
				routing: {
					request: {
						method: 'GET',
						url: '=/users/{{ $parameter["user"] }}/authentication/passwordMethods',
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
				action: 'Get password method',
			},
			{
				name: 'Reset Password',
				value: 'resetPassword',
				description: "Set a new password on a user's account, generating one if none is given",
				// No `routing`: handled by a custom operation so the generated password can be
				// returned. See AuthenticationFunctions.ts.
				action: 'Reset user password',
			},
		],
		default: 'getAllMethods',
	},
];

const deleteMethodFields: INodeProperties[] = [
	{
		displayName: 'Method Type',
		name: 'methodType',
		default: 'microsoftAuthenticatorMethods',
		description: 'Which kind of method to remove',
		displayOptions: {
			show: {
				resource: ['authentication'],
				operation: ['deleteMethod'],
			},
		},
		options: DELETABLE_METHOD_TYPES,
		type: 'options',
	},
	{
		displayName: 'Method',
		name: 'method',
		default: {
			mode: 'list',
			value: '',
		},
		description:
			'The registered method to remove. The list holds the methods of the chosen type that this user has registered.',
		displayOptions: {
			show: {
				resource: ['authentication'],
				operation: ['deleteMethod'],
			},
		},
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getAuthenticationMethods',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				// Phone methods are the exception: their IDs are three fixed GUIDs, one per
				// phone type, the same in every tenant.
				placeholder: 'e.g. 3179e48a-750b-4051-897c-87b9720928f7 (mobile phone)',
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	},
];

const createTemporaryAccessPassFields: INodeProperties[] = [
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		displayOptions: {
			show: {
				resource: ['authentication'],
				operation: ['createTemporaryAccessPass'],
			},
		},
		options: [
			{
				displayName: 'One-Time Use',
				name: 'isUsableOnce',
				default: false,
				description:
					'Whether the pass stops working after a single sign-in. A multi-use pass is only accepted if the Temporary Access Pass policy allows it.',
				routing: {
					send: {
						property: 'isUsableOnce',
						type: 'body',
					},
				},
				type: 'boolean',
			},
			{
				displayName: 'Lifetime (Minutes)',
				name: 'lifetimeInMinutes',
				default: 60,
				description:
					'How long the pass stays valid, between 10 minutes and 43200 (30 days). Leave the option out to use the tenant policy default.',
				routing: {
					send: {
						property: 'lifetimeInMinutes',
						type: 'body',
					},
				},
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
				routing: {
					send: {
						property: 'startDateTime',
						type: 'body',
					},
				},
				type: 'dateTime',
			},
		],
		placeholder: 'Add option',
		type: 'collection',
	},
];

const resetPasswordFields: INodeProperties[] = [
	{
		displayName:
			'Requires the User-PasswordProfile.ReadWrite.All application permission with admin consent, and the service principal must have at least the User Administrator Microsoft Entra role.',
		name: 'resetPasswordPermissionNotice',
		type: 'notice',
		default: '',
		displayOptions: {
			show: {
				resource: ['authentication'],
				operation: ['resetPassword'],
			},
		},
	},
	{
		displayName: 'Options',
		name: 'options',
		default: {},
		displayOptions: {
			show: {
				resource: ['authentication'],
				operation: ['resetPassword'],
			},
		},
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

/**
 * Every operation here addresses the user in its URL, so the picker is also where the
 * path-safety check hangs — one hook covering the whole resource.
 */
const guarded = (property: INodeProperties): INodeProperties => ({
	...property,
	routing: { send: { preSend: [validatePathSegmentsPreSend] } },
});

export const authenticationFields: INodeProperties[] = [
	guarded(
		userLocator(
			'authentication',
			['getAllMethods', 'getPasswordMethod'],
			'The user whose authentication methods should be retrieved',
		),
	),
	guarded(userLocator('authentication', WRITE_OPERATIONS, 'The user to act on')),
	...createTemporaryAccessPassFields,
	...deleteMethodFields,
	...resetPasswordFields,
];
