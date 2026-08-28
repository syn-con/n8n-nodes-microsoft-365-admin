import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { splitSeparateOnly, toGraphUserBody, validateUserFields } from '../../helpers/user';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

export const properties: INodeProperties[] = [
	userRLC('User to Update', 'The user to update'),
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		default: {},
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
				displayName: 'Display Name',
				name: 'displayName',
				default: '',
				description: 'The name to display in the address book for the user',
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
				displayName: 'Mail Nickname',
				name: 'mailNickname',
				default: '',
				description: 'The mail alias for the user',
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
				displayName: 'Password',
				name: 'password',
				default: '',
				description:
					'The password for the user. The password must satisfy minimum requirements as specified by the passwordPolicies property.',
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
				displayName: 'User Principal Name',
				name: 'userPrincipalName',
				default: '',
				description: 'The user principal name (UPN)',
				placeholder: 'e.g. AdeleV@contoso.com',
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
		type: 'collection',
	},
];

const displayOptions = {
	show: {
		resource: ['user'],
		operation: ['update'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const node = this.getNode();

	const user = assertPathSafe(
		node,
		this.getNodeParameter('user', index, '', { extractValue: true }),
		'user',
		index,
	);

	const fields = this.getNodeParameter('updateFields', index, {}) as IDataObject;
	validateUserFields(node, fields, index);

	const body = toGraphUserBody(fields);
	// Graph rejects these unless they arrive with nothing else alongside them.
	const separateBody = splitSeparateOnly(body);

	// An empty PATCH is what Graph rejects as "Empty Payload", so a no-op update simply
	// does not go out — declarative routing had to send it and swallow the 400.
	if (Object.keys(body).length > 0) {
		await microsoftApiRequest.call(this, 'PATCH', `/users/${user}`, body, { itemIndex: index });
	}

	if (Object.keys(separateBody).length > 0) {
		await microsoftApiRequest.call(this, 'PATCH', `/users/${user}`, separateBody, {
			itemIndex: index,
		});
	}

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ updated: true }), {
		itemData: { item: index },
	});
}
