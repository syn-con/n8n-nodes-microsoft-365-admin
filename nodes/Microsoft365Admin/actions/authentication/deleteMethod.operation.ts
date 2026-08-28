import {
	NodeOperationError,
	type IExecuteFunctions,
	type INodeExecutionData,
	type INodeProperties,
} from 'n8n-workflow';

import { authenticationMethodRLC, userRLC } from '../../helpers/descriptions';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

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

/** Guards `methodType` against anything an expression might supply. */
const DELETABLE_METHOD_VALUES = DELETABLE_METHOD_TYPES.map((type) => type.value);

export const properties: INodeProperties[] = [
	userRLC('User', 'The user to act on'),
	{
		displayName: 'Method Type',
		name: 'methodType',
		default: 'microsoftAuthenticatorMethods',
		description: 'Which kind of method to remove',
		options: DELETABLE_METHOD_TYPES,
		type: 'options',
	},
	authenticationMethodRLC(
		'Method',
		'The registered method to remove. The list holds the methods of the chosen type that this user has registered.',
	),
];

const displayOptions = {
	show: {
		resource: ['authentication'],
		operation: ['deleteMethod'],
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
	const method = assertPathSafe(
		node,
		this.getNodeParameter('method', index, '', { extractValue: true }),
		'method',
		index,
	);

	// The collection segment is part of the address: an ID alone does not identify a
	// method. See `helpers/authentication.ts`.
	const methodType = this.getNodeParameter('methodType', index, '') as string;

	// The parameter is a fixed list, but an expression can put anything here — and this
	// value goes straight into the request path.
	if (!DELETABLE_METHOD_VALUES.includes(methodType)) {
		throw new NodeOperationError(node, `'${methodType}' is not a deletable method type`, {
			itemIndex: index,
			description: `Choose one of: ${DELETABLE_METHOD_VALUES.join(', ')}. Graph has no delete for a password method.`,
		});
	}

	await microsoftApiRequest.call(
		this,
		'DELETE',
		`/users/${user}/authentication/${methodType}/${method}`,
		{},
		{ itemIndex: index },
	);

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ deleted: true }), {
		itemData: { item: index },
	});
}
