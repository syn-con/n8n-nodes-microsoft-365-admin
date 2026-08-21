import type { DateTime } from 'luxon';
import {
	NodeApiError,
	NodeOperationError,
	type IDataObject,
	type JsonObject,
	type IExecuteSingleFunctions,
	type IHttpRequestOptions,
	type IN8nHttpFullResponse,
	type INodeExecutionData,
	type INodeProperties,
} from 'n8n-workflow';

import { ignoreHttpStatusErrorsConfig, userLocator as sharedUserLocator } from './common';
import { microsoftApiRequest } from '../GenericFunctions';
import { handleErrorPostReceive } from '../GraphErrors';
import { deepMerge } from '../utils';

/**
 * Graph refuses these unless they travel in a PATCH of their own, with no other property
 * alongside them.
 */
const SEPARATE_ONLY_PROPERTIES = [
	'aboutMe',
	'birthday',
	'interests',
	'mySite',
	'pastProjects',
	'responsibilities',
	'schools',
	'skills',
];

/** Rewrites the collected fields into the shapes Graph accepts. */
function toGraphUserBody(fields: IDataObject): IDataObject {
	const body: IDataObject = { ...fields };

	for (const key of ['birthday', 'employeeHireDate', 'employeeLeaveDateTime']) {
		if (body[key]) {
			body[key] = (body[key] as DateTime).toUTC().toISO();
		}
	}
	if (body.businessPhones) {
		body.businessPhones = [body.businessPhones as string];
	}
	if (body.employeeOrgData) {
		body.employeeOrgData = (body.employeeOrgData as IDataObject).employeeOrgValues;
	}
	if (body.passwordPolicies) {
		body.passwordPolicies = (body.passwordPolicies as string[]).join(',');
	}
	// forceChangePasswordNextSignInWithMfa doesn't seem to take effect when providing it in
	// the initial create request, so it is applied here instead. Only the two known flags are
	// honoured, so an expression cannot name an arbitrary property of the password profile.
	if (body.forceChangePassword) {
		const flag = body.forceChangePassword as string;
		if (['forceChangePasswordNextSignIn', 'forceChangePasswordNextSignInWithMfa'].includes(flag)) {
			body.passwordProfile ??= {};
			(body.passwordProfile as IDataObject)[flag] = true;
		}
		delete body.forceChangePassword;
	}

	return body;
}

/** Moves the properties Graph only accepts alone out of `body` and into their own object. */
function splitSeparateOnly(body: IDataObject): IDataObject {
	const separateBody: IDataObject = {};

	for (const key of SEPARATE_ONLY_PROPERTIES) {
		if (key in body) {
			separateBody[key] = body[key];
			delete body[key];
		}
	}

	return separateBody;
}

export const userOperations: INodeProperties[] = [
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
				routing: {
					request: {
						method: 'POST',
						url: '=/groups/{{ $parameter["group"] }}/members/$ref',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'set',
								properties: {
									value: '={{ { "added": true } }}',
								},
							},
						],
					},
				},
				action: 'Add user to group',
			},
			{
				name: 'Create',
				value: 'create',
				description: 'Create a user',
				routing: {
					request: {
						method: 'POST',
						url: '/users',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
				action: 'Create user',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a user',
				routing: {
					request: {
						method: 'DELETE',
						url: '=/users/{{ $parameter["user"] }}',
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
				action: 'Delete user',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Retrieve data for a specific user',
				routing: {
					request: {
						method: 'GET',
						url: '=/users/{{ $parameter["user"] }}',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
				action: 'Get user',
			},
			{
				name: 'Get Groups',
				value: 'getGroups',
				description: 'Retrieve the groups a user is a member of',
				routing: {
					request: {
						method: 'GET',
						url: '=/users/{{ $parameter["user"] }}/memberOf',
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
				action: 'Get groups for user',
			},
			{
				name: 'Get Manager',
				value: 'getManager',
				description: "Retrieve a user's manager",
				routing: {
					request: {
						method: 'GET',
						url: '=/users/{{ $parameter["user"] }}/manager',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [handleErrorPostReceive],
					},
				},
				action: 'Get manager for user',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Retrieve a list of users',
				routing: {
					request: {
						method: 'GET',
						url: '/users',
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
				action: 'Get many users',
			},
			{
				// eslint-disable-next-line n8n-nodes-base/node-param-display-name-miscased
				name: 'Remove from Group',
				value: 'removeGroup',
				description: 'Remove user from group',
				routing: {
					request: {
						method: 'DELETE',
						url: '=/groups/{{ $parameter["group"] }}/members/{{ $parameter["user"] }}/$ref',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'set',
								properties: {
									value: '={{ { "removed": true } }}',
								},
							},
						],
					},
				},
				action: 'Remove user from group',
			},
			{
				name: 'Revoke Sessions',
				value: 'revokeSessions',
				description: 'Invalidate all refresh tokens and browser sessions for a user',
				routing: {
					request: {
						method: 'POST',
						url: '=/users/{{ $parameter["user"] }}/revokeSignInSessions',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'set',
								properties: {
									value: '={{ { "revoked": true } }}',
								},
							},
						],
					},
				},
				action: 'Revoke sessions for user',
			},
			{
				name: 'Set Manager',
				value: 'setManager',
				description: "Set a user's manager",
				routing: {
					request: {
						method: 'PUT',
						url: '=/users/{{ $parameter["user"] }}/manager/$ref',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'set',
								properties: {
									value: '={{ { "updated": true } }}',
								},
							},
						],
					},
				},
				action: 'Set manager for user',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update a user',
				routing: {
					request: {
						method: 'PATCH',
						url: '=/users/{{ $parameter["user"] }}',
						ignoreHttpStatusErrors: ignoreHttpStatusErrorsConfig,
					},
					output: {
						postReceive: [
							handleErrorPostReceive,
							{
								type: 'set',
								properties: {
									value: '={{ { "updated": true } }}',
								},
							},
						],
					},
				},
				action: 'Update user',
			},
		],
		default: 'getAll',
	},
];

const addGroupFields: INodeProperties[] = [
	{
		displayName: 'Group',
		name: 'group',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['addGroup'],
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
		displayName: 'User to Add',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['addGroup'],
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
		routing: {
			send: {
				property: '@odata.id',
				propertyInDotNotation: false,
				type: 'body',
				value:
					'={{ ($credentials.graphApiBaseUrl || "https://graph.microsoft.com").replace(/\\/+$/, "") }}/v1.0/directoryObjects/{{ $value }}',
			},
		},
		type: 'resourceLocator',
	},
];

const createFields: INodeProperties[] = [
	{
		displayName: 'Account Enabled',
		name: 'accountEnabled',
		default: true,
		description: 'Whether the account is enabled',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['create'],
			},
		},
		required: true,
		routing: {
			send: {
				property: 'accountEnabled',
				type: 'body',
			},
		},
		type: 'boolean',
		validateType: 'boolean',
	},
	{
		displayName: 'Display Name',
		name: 'displayName',
		default: '',
		description: 'The name to display in the address book for the user',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['create'],
			},
		},
		placeholder: 'e.g. Nathan Smith',
		required: true,
		routing: {
			send: {
				property: 'displayName',
				type: 'body',
			},
		},
		type: 'string',
		validateType: 'string',
	},
	{
		displayName: 'User Principal Name',
		name: 'userPrincipalName',
		default: '',
		description: 'The user principal name (UPN)',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['create'],
			},
		},
		placeholder: 'e.g. NathanSmith@contoso.com',
		required: true,
		routing: {
			send: {
				property: 'userPrincipalName',
				type: 'body',
				preSend: [
					async function (
						this: IExecuteSingleFunctions,
						requestOptions: IHttpRequestOptions,
					): Promise<IHttpRequestOptions> {
						const userPrincipalName = this.getNodeParameter('userPrincipalName') as string;
						if (!/^[A-Za-z0-9'._\-!#^~@]+$/.test(userPrincipalName)) {
							throw new NodeOperationError(
								this.getNode(),
								"Only the following characters are allowed for 'User Principal Name': A-Z, a-z, 0-9, ' . - _ ! # ^ ~",
							);
						}
						return requestOptions;
					},
				],
			},
		},
		type: 'string',
		validateType: 'string',
	},
	{
		displayName: 'Mail Nickname',
		name: 'mailNickname',
		default: '',
		description: 'The mail alias for the user',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['create'],
			},
		},
		placeholder: 'e.g. NathanSmith',
		required: true,
		routing: {
			send: {
				property: 'mailNickname',
				type: 'body',
			},
		},
		type: 'string',
		validateType: 'string',
	},
	{
		displayName: 'Password',
		name: 'password',
		default: '',
		description: 'The password for the user',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['create'],
			},
		},
		required: true,
		routing: {
			send: {
				property: 'passwordProfile.password',
				type: 'body',
			},
		},
		type: 'string',
		typeOptions: {
			password: true,
		},
		validateType: 'string',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		default: {},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['create'],
			},
		},
		options: [
			{
				displayName: 'About Me',
				name: 'aboutMe',
				default: '',
				description: 'A freeform text entry field for the user to describe themselves',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Age Group',
				name: 'ageGroup',
				default: 'Adult',
				description: 'Sets the age group of the user',
				options: [
					{
						name: 'Adult',
						value: 'Adult',
					},
					{
						name: 'Minor',
						value: 'Minor',
					},
					{
						name: 'Not Adult',
						value: 'NotAdult',
					},
				],
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Birthday',
				name: 'birthday',
				default: '',
				description: 'The birthday of the user',
				type: 'dateTime',
				validateType: 'dateTime',
			},
			{
				displayName: 'Business Phone',
				name: 'businessPhones',
				default: '',
				description: 'The telephone number for the user',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'City',
				name: 'city',
				default: '',
				description: 'The city in which the user is located',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Company Name',
				name: 'companyName',
				default: '',
				description: 'The name of the company associated with the user',
				routing: {
					send: {
						preSend: [
							async function (
								this: IExecuteSingleFunctions,
								requestOptions: IHttpRequestOptions,
							): Promise<IHttpRequestOptions> {
								const companyName = this.getNodeParameter('additionalFields.companyName') as string;
								if (companyName?.length > 64) {
									throw new NodeOperationError(
										this.getNode(),
										"'Company Name' should have a maximum length of 64",
									);
								}
								return requestOptions;
							},
						],
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Consent Provided',
				name: 'consentProvidedForMinor',
				default: 'Denied',
				description: 'Specifies if consent is provided for minors',
				options: [
					{
						name: 'Denied',
						value: 'Denied',
					},
					{
						name: 'Granted',
						value: 'Granted',
					},
					{
						name: 'Not Required',
						value: 'NotRequired',
					},
				],
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Country',
				name: 'country',
				default: '',
				description: 'The country/region of the user',
				placeholder: 'e.g. US',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Department',
				name: 'department',
				default: '',
				description: 'The department name where the user works',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Employee Hire Date',
				name: 'employeeHireDate',
				default: '',
				description: 'The hire date of the user',
				placeholder: 'e.g. 2014-01-01T00:00:00Z',
				type: 'dateTime',
				validateType: 'dateTime',
			},
			{
				displayName: 'Employee ID',
				name: 'employeeId',
				default: '',
				description: 'Employee identifier assigned by the organization',
				routing: {
					send: {
						preSend: [
							async function (
								this: IExecuteSingleFunctions,
								requestOptions: IHttpRequestOptions,
							): Promise<IHttpRequestOptions> {
								const employeeId = this.getNodeParameter('additionalFields.employeeId') as string;
								if (employeeId?.length > 16) {
									throw new NodeOperationError(
										this.getNode(),
										"'Employee ID' should have a maximum length of 16",
									);
								}
								return requestOptions;
							},
						],
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Employee Leave Date',
				name: 'employeeLeaveDateTime',
				default: '',
				description: 'The date and time when the user left or will leave the organization',
				placeholder: 'e.g. 2014-01-01T00:00:00Z',
				type: 'dateTime',
				validateType: 'dateTime',
			},
			{
				displayName: 'Employee Organization Data',
				name: 'employeeOrgData',
				default: {},
				description:
					'Represents organization data (for example, division and costCenter) associated with a user',
				options: [
					{
						displayName: 'Employee Organization Data',
						name: 'employeeOrgValues',
						values: [
							{
								displayName: 'Cost Center',
								name: 'costCenter',
								description: 'The cost center associated with the user',
								type: 'string',
								default: '',
							},
							{
								displayName: 'Division',
								name: 'division',
								description: 'The name of the division in which the user works',
								type: 'string',
								default: '',
							},
						],
					},
				],
				type: 'fixedCollection',
				validateType: 'string',
			},
			{
				displayName: 'Employee Type',
				name: 'employeeType',
				default: '',
				description: 'Defines enterprise worker type',
				placeholder: 'e.g. Contractor',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'First Name',
				name: 'givenName',
				default: '',
				description: 'The given name (first name) of the user',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Force Change Password',
				name: 'forceChangePassword',
				default: 'forceChangePasswordNextSignIn',
				description: 'Whether the user must change their password on the next sign-in',
				options: [
					{
						name: 'Next Sign In',
						value: 'forceChangePasswordNextSignIn',
					},
					{
						name: 'Next Sign In with MFA',
						value: 'forceChangePasswordNextSignInWithMfa',
					},
				],
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Interests',
				name: 'interests',
				default: [],
				description: 'A list for the user to describe their interests',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Job Title',
				name: 'jobTitle',
				default: '',
				description: "The user's job title",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Last Name',
				name: 'surname',
				default: '',
				description: "The user's last name (family name)",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Mail',
				name: 'mail',
				default: '',
				description: 'The SMTP address for the user',
				placeholder: 'e.g. jeff@contoso.com',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Mobile Phone',
				name: 'mobilePhone',
				default: '',
				description: 'The primary cellular telephone number for the user',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'My Site',
				name: 'mySite',
				default: '',
				description: "The URL for the user's personal site",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Office Location',
				name: 'officeLocation',
				default: '',
				description: 'The office location for the user',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'On Premises Immutable ID',
				name: 'onPremisesImmutableId',
				default: '',
				description:
					'This property is used to associate an on-premises Active Directory user account to their Microsoft Entra user object',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Other Emails',
				name: 'otherMails',
				default: [],
				description: 'Additional email addresses for the user',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Password Policies',
				name: 'passwordPolicies',
				default: [],
				description: 'Specifies password policies',
				options: [
					{
						name: 'Disable Password Expiration',
						value: 'DisablePasswordExpiration',
					},
					{
						name: 'Disable Strong Password',
						value: 'DisableStrongPassword',
					},
				],
				type: 'multiOptions',
			},
			{
				displayName: 'Past Projects',
				name: 'pastProjects',
				default: [],
				description: 'A list of past projects the user has worked on',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Postal Code',
				name: 'postalCode',
				default: '',
				description: "The postal code for the user's address",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Preferred Language',
				name: 'preferredLanguage',
				default: '',
				description: "User's preferred language in ISO 639-1 code",
				placeholder: 'e.g. en-US',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Responsibilities',
				name: 'responsibilities',
				default: [],
				description: 'A list of responsibilities the user has',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Schools Attended',
				name: 'schools',
				default: [],
				description: 'A list of schools the user attended',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Skills',
				name: 'skills',
				default: [],
				description: 'A list of skills the user possesses',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'State',
				name: 'state',
				default: '',
				description: "The state or province of the user's address",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Street Address',
				name: 'streetAddress',
				default: '',
				description: "The street address of the user's place of business",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Usage Location',
				name: 'usageLocation',
				default: '',
				description: 'Two-letter country code where the user is located',
				placeholder: 'e.g. US',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'User Type',
				name: 'userType',
				default: 'Guest',
				description: 'Classifies the user type',
				options: [
					{
						name: 'Guest',
						value: 'Guest',
					},
					{
						name: 'Member',
						value: 'Member',
					},
				],
				type: 'options',
				validateType: 'options',
			},
		],
		placeholder: 'Add Field',
		routing: {
			output: {
				postReceive: [
					async function (
						this: IExecuteSingleFunctions,
						items: INodeExecutionData[],
						_response: IN8nHttpFullResponse,
					): Promise<INodeExecutionData[]> {
						for (const item of items) {
							const userId = item.json.id as string;
							const fields = this.getNodeParameter('additionalFields', item.index) as IDataObject;
							if (!Object.keys(fields).length) {
								continue;
							}

							const body = toGraphUserBody(fields);
							const separateBody = splitSeparateOnly(body);

							try {
								if (Object.keys(separateBody).length) {
									await microsoftApiRequest.call(this, 'PATCH', `/users/${userId}`, separateBody);
									deepMerge(item.json, separateBody);
								}
								if (Object.keys(body).length) {
									await microsoftApiRequest.call(this, 'PATCH', `/users/${userId}`, body);
									deepMerge(item.json, body);
								}
							} catch (error) {
								try {
									await microsoftApiRequest.call(this, 'DELETE', `/users/${userId}`);
								} catch {}
								// Keep an already-typed node error as-is; wrap anything else so the
								// workflow always surfaces a NodeApiError rather than a bare Error.
								throw error instanceof NodeApiError
									? error
									: new NodeApiError(this.getNode(), error as JsonObject);
							}
						}
						return items;
					},
				],
			},
		},
		type: 'collection',
	},
];

const deleteFields: INodeProperties[] = [
	{
		displayName: 'User to Delete',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['delete'],
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

const getFields: INodeProperties[] = [
	{
		displayName: 'User to Get',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['get'],
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
		displayName: 'Output',
		name: 'output',
		default: 'simple',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['get'],
			},
		},
		options: [
			{
				name: 'Simplified',
				value: 'simple',
				routing: {
					send: {
						property: '$select',
						type: 'query',
						value:
							'id,createdDateTime,displayName,userPrincipalName,mail,mailNickname,securityIdentifier',
					},
				},
			},
			{
				name: 'Raw',
				value: 'raw',
				routing: {
					send: {
						property: '$select',
						type: 'query',
						value:
							'id,accountEnabled,ageGroup,assignedLicenses,assignedPlans,authorizationInfo,businessPhones,city,companyName,consentProvidedForMinor,country,createdDateTime,creationType,customSecurityAttributes,deletedDateTime,department,displayName,employeeHireDate,employeeId,employeeLeaveDateTime,employeeOrgData,employeeType,externalUserState,externalUserStateChangeDateTime,faxNumber,givenName,identities,imAddresses,isManagementRestricted,isResourceAccount,jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,mail,mailNickname,mobilePhone,officeLocation,onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,securityIdentifier,serviceProvisioningErrors,showInAddressList,signInSessionsValidFromDateTime,state,streetAddress,surname,usageLocation,userPrincipalName,userType',
					},
				},
			},
			{
				name: 'Selected Fields',
				value: 'fields',
			},
		],
		type: 'options',
	},
	{
		// eslint-disable-next-line n8n-nodes-base/node-param-display-name-wrong-for-dynamic-multi-options
		displayName: 'Fields',
		name: 'fields',
		default: [],
		// eslint-disable-next-line n8n-nodes-base/node-param-description-wrong-for-dynamic-multi-options
		description: 'The fields to add to the output',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['get'],
				output: ['fields'],
			},
		},
		routing: {
			send: {
				property: '$select',
				type: 'query',
				value: '={{ $value.concat("id").join(",") }}',
			},
		},
		typeOptions: {
			loadOptionsMethod: 'getUserProperties',
		},
		type: 'multiOptions',
	},
];

const getAllFields: INodeProperties[] = [
	{
		displayName: 'Return All',
		name: 'returnAll',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getAll'],
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
							qs: {
								$filter:
									'={{ !!$response.body?.["@odata.nextLink"] ? undefined : $request.qs?.$filter }}',
								$select:
									'={{ !!$response.body?.["@odata.nextLink"] ? undefined : $request.qs?.$select }}',
							},
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
				resource: ['user'],
				operation: ['getAll'],
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
	{
		displayName: 'Filter',
		name: 'filter',
		default: '',
		description:
			'<a href="https://docs.microsoft.com/en-us/graph/query-parameters#filter-parameter">Query parameter</a> to filter results by',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getAll'],
			},
		},
		placeholder: "e.g. startswith(displayName, 'a')",
		routing: {
			send: {
				property: '$filter',
				type: 'query',
				value: '={{ $value ? $value : undefined }}',
			},
		},
		type: 'string',
		validateType: 'string',
	},
	{
		displayName: 'Output',
		name: 'output',
		default: 'simple',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				name: 'Simplified',
				value: 'simple',
				routing: {
					send: {
						property: '$select',
						type: 'query',
						value:
							'id,createdDateTime,displayName,userPrincipalName,mail,mailNickname,securityIdentifier',
					},
				},
			},
			{
				name: 'Raw',
				value: 'raw',
				routing: {
					send: {
						property: '$select',
						type: 'query',
						value:
							'id,accountEnabled,ageGroup,assignedLicenses,assignedPlans,authorizationInfo,businessPhones,city,companyName,consentProvidedForMinor,country,createdDateTime,creationType,customSecurityAttributes,deletedDateTime,department,displayName,employeeHireDate,employeeId,employeeLeaveDateTime,employeeOrgData,employeeType,externalUserState,externalUserStateChangeDateTime,faxNumber,givenName,identities,imAddresses,isManagementRestricted,isResourceAccount,jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,mail,mailNickname,mobilePhone,officeLocation,onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,securityIdentifier,serviceProvisioningErrors,showInAddressList,signInSessionsValidFromDateTime,state,streetAddress,surname,usageLocation,userPrincipalName,userType',
					},
				},
			},
			{
				name: 'Selected Fields',
				value: 'fields',
			},
		],
		type: 'options',
	},
	{
		// eslint-disable-next-line n8n-nodes-base/node-param-display-name-wrong-for-dynamic-multi-options
		displayName: 'Fields',
		name: 'fields',
		default: [],
		// eslint-disable-next-line n8n-nodes-base/node-param-description-wrong-for-dynamic-multi-options
		description: 'The fields to add to the output',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getAll'],
				output: ['fields'],
			},
		},
		routing: {
			send: {
				property: '$select',
				type: 'query',
				value: '={{ $value.concat("id").join(",") }}',
			},
		},
		typeOptions: {
			loadOptionsMethod: 'getUserPropertiesGetAll',
		},
		type: 'multiOptions',
	},
];

const removeGroupFields: INodeProperties[] = [
	{
		displayName: 'Group',
		name: 'group',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['removeGroup'],
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
		displayName: 'User to Remove',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['removeGroup'],
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

const updateFields: INodeProperties[] = [
	{
		displayName: 'User to Update',
		name: 'user',
		default: {
			mode: 'list',
			value: '',
		},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['update'],
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
		displayName: 'Update Fields',
		name: 'updateFields',
		default: {},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['update'],
			},
		},
		options: [
			{
				displayName: 'About Me',
				name: 'aboutMe',
				default: '',
				description: 'A freeform text entry field for the user to describe themselves',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Account Enabled',
				name: 'accountEnabled',
				default: true,
				description: 'Whether the account is enabled',
				routing: {
					send: {
						property: 'accountEnabled',
						type: 'body',
					},
				},
				type: 'boolean',
				validateType: 'boolean',
			},
			{
				displayName: 'Age Group',
				name: 'ageGroup',
				default: 'Adult',
				description: 'Sets the age group of the user',
				options: [
					{
						name: 'Adult',
						value: 'Adult',
					},
					{
						name: 'Minor',
						value: 'Minor',
					},
					{
						name: 'Not Adult',
						value: 'NotAdult',
					},
				],
				routing: {
					send: {
						property: 'ageGroup',
						type: 'body',
					},
				},
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Birthday',
				name: 'birthday',
				default: '',
				description: 'The birthday of the user',
				type: 'dateTime',
				validateType: 'dateTime',
			},
			{
				displayName: 'Business Phone',
				name: 'businessPhones',
				default: '',
				description: 'The telephone number for the user',
				routing: {
					send: {
						property: 'businessPhones',
						type: 'body',
						value: '={{ $value ? [$value] : [] }}',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'City',
				name: 'city',
				default: '',
				description: 'The city in which the user is located',
				routing: {
					send: {
						property: 'city',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Company Name',
				name: 'companyName',
				default: '',
				description: 'The name of the company associated with the user',
				routing: {
					send: {
						property: 'companyName',
						type: 'body',
						preSend: [
							async function (
								this: IExecuteSingleFunctions,
								requestOptions: IHttpRequestOptions,
							): Promise<IHttpRequestOptions> {
								const companyName = this.getNodeParameter('updateFields.companyName') as string;
								if (companyName?.length > 64) {
									throw new NodeOperationError(
										this.getNode(),
										"'Company Name' should have a maximum length of 64",
									);
								}
								return requestOptions;
							},
						],
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Consent Provided',
				name: 'consentProvidedForMinor',
				default: 'Denied',
				description: 'Specifies if consent is provided for minors',
				options: [
					{
						name: 'Denied',
						value: 'Denied',
					},
					{
						name: 'Granted',
						value: 'Granted',
					},
					{
						name: 'Not Required',
						value: 'NotRequired',
					},
				],
				routing: {
					send: {
						property: 'consentProvidedForMinor',
						type: 'body',
					},
				},
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Country',
				name: 'country',
				default: '',
				description: 'The country/region of the user',
				placeholder: 'e.g. US',
				routing: {
					send: {
						property: 'country',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Department',
				name: 'department',
				default: '',
				description: 'The department name where the user works',
				routing: {
					send: {
						property: 'department',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Display Name',
				name: 'displayName',
				default: '',
				description: 'The name to display in the address book for the user',
				routing: {
					send: {
						property: 'displayName',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Employee Hire Date',
				name: 'employeeHireDate',
				default: '',
				description: 'The hire date of the user',
				placeholder: 'e.g. 2014-01-01T00:00:00Z',
				routing: {
					send: {
						property: 'employeeHireDate',
						type: 'body',
						value: '={{ $value?.toUTC().toISO() }}',
					},
				},
				type: 'dateTime',
				validateType: 'dateTime',
			},
			{
				displayName: 'Employee ID',
				name: 'employeeId',
				default: '',
				description: 'Employee identifier assigned by the organization',
				routing: {
					send: {
						property: 'employeeId',
						type: 'body',
						preSend: [
							async function (
								this: IExecuteSingleFunctions,
								requestOptions: IHttpRequestOptions,
							): Promise<IHttpRequestOptions> {
								const employeeId = this.getNodeParameter('updateFields.employeeId') as string;
								if (employeeId?.length > 16) {
									throw new NodeOperationError(
										this.getNode(),
										"'Employee ID' should have a maximum length of 16",
									);
								}
								return requestOptions;
							},
						],
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Employee Leave Date',
				name: 'employeeLeaveDateTime',
				default: '',
				description: 'The date and time when the user left or will leave the organization',
				placeholder: 'e.g. 2014-01-01T00:00:00Z',
				routing: {
					send: {
						property: 'employeeLeaveDateTime',
						type: 'body',
						value: '={{ $value?.toUTC().toISO() }}',
					},
				},
				type: 'dateTime',
				validateType: 'dateTime',
			},
			{
				displayName: 'Employee Organization Data',
				name: 'employeeOrgData',
				default: {},
				description:
					'Represents organization data (for example, division and costCenter) associated with a user',
				options: [
					{
						displayName: 'Employee Organization Data',
						name: 'employeeOrgValues',
						values: [
							{
								displayName: 'Cost Center',
								name: 'costCenter',
								description: 'The cost center associated with the user',
								routing: {
									send: {
										property: 'employeeOrgData.costCenter',
										type: 'body',
									},
								},
								type: 'string',
								default: '',
							},
							{
								displayName: 'Division',
								name: 'division',
								description: 'The name of the division in which the user works',
								routing: {
									send: {
										property: 'employeeOrgData.division',
										type: 'body',
									},
								},
								type: 'string',
								default: '',
							},
						],
					},
				],
				type: 'fixedCollection',
				validateType: 'string',
			},
			{
				displayName: 'Employee Type',
				name: 'employeeType',
				default: '',
				description: 'Defines enterprise worker type',
				placeholder: 'e.g. Contractor',
				routing: {
					send: {
						property: 'employeeType',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'First Name',
				name: 'givenName',
				default: '',
				description: 'The given name (first name) of the user',
				routing: {
					send: {
						property: 'givenName',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Force Change Password',
				name: 'forceChangePassword',
				default: 'forceChangePasswordNextSignIn',
				description: 'Whether the user must change their password on the next sign-in',
				options: [
					{
						name: 'Next Sign In',
						value: 'forceChangePasswordNextSignIn',
					},
					{
						name: 'Next Sign In with MFA',
						value: 'forceChangePasswordNextSignInWithMfa',
					},
				],
				routing: {
					send: {
						preSend: [
							async function (
								this: IExecuteSingleFunctions,
								requestOptions: IHttpRequestOptions,
							): Promise<IHttpRequestOptions> {
								const forceChangePassword = this.getNodeParameter(
									'updateFields.forceChangePassword',
								) as string;
								if (forceChangePassword === 'forceChangePasswordNextSignIn') {
									(requestOptions.body as IDataObject).passwordProfile ??= {};
									(
										(requestOptions.body as IDataObject).passwordProfile as IDataObject
									).forceChangePasswordNextSignIn = true;
								} else if (forceChangePassword === 'forceChangePasswordNextSignInWithMfa') {
									(
										(requestOptions.body as IDataObject).passwordProfile as IDataObject
									).forceChangePasswordNextSignInWithMfa = true;
								}
								return requestOptions;
							},
						],
					},
				},
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Interests',
				name: 'interests',
				default: [],
				description: 'A list for the user to describe their interests',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Job Title',
				name: 'jobTitle',
				default: '',
				description: "The user's job title",
				routing: {
					send: {
						property: 'jobTitle',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Last Name',
				name: 'surname',
				default: '',
				description: "The user's last name (family name)",
				routing: {
					send: {
						property: 'surname',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Mail',
				name: 'mail',
				default: '',
				description: 'The SMTP address for the user',
				placeholder: 'e.g. jeff@contoso.com',
				routing: {
					send: {
						property: 'mail',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Mail Nickname',
				name: 'mailNickname',
				default: '',
				description: 'The mail alias for the user',
				routing: {
					send: {
						property: 'mailNickname',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Mobile Phone',
				name: 'mobilePhone',
				default: '',
				description: 'The primary cellular telephone number for the user',
				routing: {
					send: {
						property: 'mobilePhone',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'My Site',
				name: 'mySite',
				default: '',
				description: "The URL for the user's personal site",
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Office Location',
				name: 'officeLocation',
				default: '',
				description: 'The office location for the user',
				routing: {
					send: {
						property: 'officeLocation',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'On Premises Immutable ID',
				name: 'onPremisesImmutableId',
				default: '',
				description:
					'This property is used to associate an on-premises Active Directory user account to their Microsoft Entra user object',
				routing: {
					send: {
						property: 'onPremisesImmutableId',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Other Emails',
				name: 'otherMails',
				default: [],
				description: 'Additional email addresses for the user',
				routing: {
					send: {
						property: 'otherMails',
						type: 'body',
					},
				},
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Password',
				name: 'password',
				default: '',
				description:
					'The password for the user. The password must satisfy minimum requirements as specified by the passwordPolicies property.',
				routing: {
					send: {
						property: 'passwordProfile.password',
						type: 'body',
					},
				},
				type: 'string',
				typeOptions: {
					password: true,
				},
				validateType: 'string',
			},
			{
				displayName: 'Password Policies',
				name: 'passwordPolicies',
				default: [],
				description: 'Specifies password policies',
				options: [
					{
						name: 'Disable Password Expiration',
						value: 'DisablePasswordExpiration',
					},
					{
						name: 'Disable Strong Password',
						value: 'DisableStrongPassword',
					},
				],
				routing: {
					send: {
						property: 'passwordPolicies',
						type: 'body',
						value: '={{ $value?.join(",") }}',
					},
				},
				type: 'multiOptions',
			},
			{
				displayName: 'Past Projects',
				name: 'pastProjects',
				default: [],
				description: 'A list of past projects the user has worked on',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Postal Code',
				name: 'postalCode',
				default: '',
				description: "The postal code for the user's address",
				routing: {
					send: {
						property: 'postalCode',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Preferred Language',
				name: 'preferredLanguage',
				default: '',
				description: "User's preferred language in ISO 639-1 code",
				placeholder: 'e.g. en-US',
				routing: {
					send: {
						property: 'preferredLanguage',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Responsibilities',
				name: 'responsibilities',
				default: [],
				description: 'A list of responsibilities the user has',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Schools Attended',
				name: 'schools',
				default: [],
				description: 'A list of schools the user attended',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'Skills',
				name: 'skills',
				default: [],
				description: 'A list of skills the user possesses',
				type: 'string',
				typeOptions: {
					multipleValues: true,
				},
				validateType: 'array',
			},
			{
				displayName: 'State',
				name: 'state',
				default: '',
				description: "The state or province of the user's address",
				routing: {
					send: {
						property: 'state',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Street Address',
				name: 'streetAddress',
				default: '',
				description: "The street address of the user's place of business",
				routing: {
					send: {
						property: 'streetAddress',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Usage Location',
				name: 'usageLocation',
				default: '',
				description: 'Two-letter country code where the user is located',
				placeholder: 'e.g. US',
				routing: {
					send: {
						property: 'usageLocation',
						type: 'body',
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'User Principal Name',
				name: 'userPrincipalName',
				default: '',
				description: 'The user principal name (UPN)',
				placeholder: 'e.g. AdeleV@contoso.com',
				routing: {
					send: {
						property: 'userPrincipalName',
						type: 'body',
						preSend: [
							async function (
								this: IExecuteSingleFunctions,
								requestOptions: IHttpRequestOptions,
							): Promise<IHttpRequestOptions> {
								const userPrincipalName = this.getNodeParameter(
									'updateFields.userPrincipalName',
								) as string;
								if (!/^[A-Za-z0-9'._\-!#^~@]+$/.test(userPrincipalName)) {
									throw new NodeOperationError(
										this.getNode(),
										"Only the following characters are allowed for 'User Principal Name': A-Z, a-z, 0-9, ' . - _ ! # ^ ~",
									);
								}
								return requestOptions;
							},
						],
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'User Type',
				name: 'userType',
				default: 'Guest',
				description: 'Classifies the user type',
				options: [
					{
						name: 'Guest',
						value: 'Guest',
					},
					{
						name: 'Member',
						value: 'Member',
					},
				],
				routing: {
					send: {
						property: 'userType',
						type: 'body',
					},
				},
				type: 'options',
				validateType: 'options',
			},
		],
		placeholder: 'Add Field',
		routing: {
			output: {
				postReceive: [
					async function (
						this: IExecuteSingleFunctions,
						items: INodeExecutionData[],
						_response: IN8nHttpFullResponse,
					): Promise<INodeExecutionData[]> {
						for (const item of items) {
							const userId = this.getNodeParameter('user.value', item.index) as string;
							const fields = this.getNodeParameter('updateFields', item.index) as IDataObject;
							// The rest of the update travels with the request the routing already built;
							// only the properties Graph insists on receiving alone are sent here.
							const separateBody = splitSeparateOnly(toGraphUserBody(fields));

							if (Object.keys(separateBody).length) {
								await microsoftApiRequest.call(this, 'PATCH', `/users/${userId}`, separateBody);
							}
						}
						return items;
					},
				],
			},
		},
		type: 'collection',
	},
];

/** The user picker, reused by the operations that target a single user. */
const userLocator = (operation: string, description: string): INodeProperties =>
	sharedUserLocator('user', operation, description);

const getGroupsFields: INodeProperties[] = [
	userLocator('getGroups', 'The user whose group memberships should be retrieved'),
	{
		displayName: 'Return All',
		name: 'returnAll',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getGroups'],
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
				resource: ['user'],
				operation: ['getGroups'],
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

const getManagerFields: INodeProperties[] = [
	userLocator('getManager', 'The user whose manager should be retrieved'),
];

const revokeSessionsFields: INodeProperties[] = [
	userLocator('revokeSessions', 'The user whose sessions should be revoked'),
];

const setManagerFields: INodeProperties[] = [
	userLocator('setManager', 'The user whose manager should be set'),
	{
		displayName: 'Manager',
		name: 'manager',
		default: {
			mode: 'list',
			value: '',
		},
		description: 'The user to set as manager',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['setManager'],
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
		routing: {
			send: {
				property: '@odata.id',
				propertyInDotNotation: false,
				type: 'body',
				value:
					'={{ ($credentials.graphApiBaseUrl || "https://graph.microsoft.com").replace(/\\/+$/, "") }}/v1.0/users/{{ $value }}',
			},
		},
		type: 'resourceLocator',
	},
];

export const userFields: INodeProperties[] = [
	...addGroupFields,
	...createFields,
	...deleteFields,
	...getFields,
	...getAllFields,
	...getGroupsFields,
	...getManagerFields,
	...removeGroupFields,
	...revokeSessionsFields,
	...setManagerFields,
	...updateFields,
];
