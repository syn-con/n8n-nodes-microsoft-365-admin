import { randomInt } from 'node:crypto';

import {
	NodeOperationError,
	type IDataObject,
	type IExecuteFunctions,
	type IExecuteSingleFunctions,
	type IHttpRequestOptions,
	type ILoadOptionsFunctions,
	type IN8nHttpFullResponse,
	type INodeExecutionData,
	type INodeListSearchItems,
	type INodeListSearchResult,
} from 'n8n-workflow';

import { microsoftApiRequest } from './GenericFunctions';
import { asNodeError, errorItem } from './utils';

/**
 * The authentication methods API keeps every method type in its own collection, so an ID
 * on its own does not address a method — the type decides the URL segment, and Graph never
 * reports that segment: `GET /authentication/methods` answers with an `@odata.type`. This
 * module bridges the two, so Get Many Methods can be piped straight into Delete Method,
 * and holds the one operation that cannot be declarative (Reset Password).
 */

/** URL segment for each method type, keyed by the `@odata.type` Graph reports. */
const METHOD_TYPE_BY_ODATA_TYPE: Record<string, string> = {
	emailAuthenticationMethod: 'emailMethods',
	externalAuthenticationMethod: 'externalAuthenticationMethods',
	fido2AuthenticationMethod: 'fido2Methods',
	microsoftAuthenticatorAuthenticationMethod: 'microsoftAuthenticatorMethods',
	passwordAuthenticationMethod: 'passwordMethods',
	phoneAuthenticationMethod: 'phoneMethods',
	platformCredentialAuthenticationMethod: 'platformCredentialMethods',
	softwareOathAuthenticationMethod: 'softwareOathMethods',
	temporaryAccessPassAuthenticationMethod: 'temporaryAccessPassMethods',
	windowsHelloForBusinessAuthenticationMethod: 'windowsHelloForBusinessMethods',
};

/** What to call a method that carries no name of its own. */
const METHOD_TYPE_LABELS: Record<string, string> = {
	emailMethods: 'Email',
	externalAuthenticationMethods: 'External provider',
	fido2Methods: 'FIDO2 security key',
	microsoftAuthenticatorMethods: 'Microsoft Authenticator',
	passwordMethods: 'Password',
	phoneMethods: 'Phone',
	platformCredentialMethods: 'Platform credential',
	softwareOathMethods: 'Software OATH token',
	temporaryAccessPassMethods: 'Temporary Access Pass',
	windowsHelloForBusinessMethods: 'Windows Hello for Business',
};

/** Graph offers no delete for a password; every other method type has one. */
const UNDELETABLE_METHOD_TYPES = ['passwordMethods'];

/** Characters a generated password is drawn from, without the easily confused ones. */
const PASSWORD_ALPHABET = {
	lower: 'abcdefghijkmnopqrstuvwxyz',
	upper: 'ABCDEFGHJKLMNPQRSTUVWXYZ',
	digits: '23456789',
	symbols: '!#$%&*+-=?@^_',
};

const DEFAULT_PASSWORD_LENGTH = 16;
const MIN_PASSWORD_LENGTH = 8;
const MAX_PASSWORD_LENGTH = 256;

/**
 * A user ID, a method ID and a method type all go into the request path, and all three
 * can come from workflow data. None of them legitimately contains any of these
 * characters — a Graph method ID is base64url — so a value carrying one is refused
 * rather than escaped: `alice/authentication/methods?x=` would otherwise turn a delete
 * of one method into a request against a different Graph endpoint.
 */
const UNSAFE_IN_TARGET = /[\\/?#%\s]/;

/** The parameters the declarative operations interpolate into the request path. */
const PATH_PARAMETERS: Record<string, string> = {
	user: 'user',
	method: 'method',
	methodType: 'method type',
};

function isFilled(value: unknown): value is string {
	return typeof value === 'string' && value.trim() !== '';
}

/** `#microsoft.graph.phoneAuthenticationMethod` → `phoneMethods`. */
export function methodTypeOf(odataType: unknown): string | undefined {
	if (typeof odataType !== 'string') {
		return undefined;
	}

	return METHOD_TYPE_BY_ODATA_TYPE[odataType.replace(/^#?microsoft\.graph\./, '')];
}

/**
 * Names a method the way an admin would recognize it.
 *
 * Which field carries the name depends on the type — a security key has a `displayName`, a
 * phone has only its number — and a Temporary Access Pass has none at all.
 */
export function describeMethod(method: IDataObject, methodType?: string): string {
	const named = [method.displayName, method.emailAddress, method.phoneNumber].find(isFilled);
	const fallback = methodType ? METHOD_TYPE_LABELS[methodType] : undefined;
	const name = named ?? fallback ?? String(method.id);

	const detail = [method.phoneType, method.model, method.createdDateTime].find(isFilled);

	return detail && detail !== name ? `${name} (${detail})` : name;
}

/**
 * Adds the two things a Delete Method step needs and Graph does not return: the URL
 * segment the method lives under, and whether it can be deleted at all.
 */
export async function annotateMethodsPostReceive(
	this: IExecuteSingleFunctions,
	data: INodeExecutionData[],
	_response: IN8nHttpFullResponse,
): Promise<INodeExecutionData[]> {
	return data.map((item) => {
		const methodType = methodTypeOf(item.json['@odata.type']);

		return {
			...item,
			json: {
				...item.json,
				methodType: methodType ?? null,
				methodName: describeMethod(item.json, methodType),
				deletable: methodType !== undefined && !UNDELETABLE_METHOD_TYPES.includes(methodType),
			},
		};
	});
}

/**
 * Refuses a path segment that could reshape the request URL.
 *
 * Attached to the User picker of this resource, so it guards every declarative operation on
 * it. Reset Password does its own checking, being a custom operation.
 */
export async function validatePathSegmentsPreSend(
	this: IExecuteSingleFunctions,
	requestOptions: IHttpRequestOptions,
): Promise<IHttpRequestOptions> {
	for (const [name, label] of Object.entries(PATH_PARAMETERS)) {
		// Read the resource locator's `value` directly rather than asking for the parameter
		// with `extractValue`: that resolves against the *displayed* properties and throws
		// "Could not find property" for a parameter this operation does not show — `method`
		// and `methodType` belong to Delete Method alone.
		const value = this.getNodeParameter(`${name}.value`, '') || this.getNodeParameter(name, '');

		// An empty value means the parameter does not apply here; whether it is required is
		// n8n's business, not this hook's.
		if (isFilled(value) && UNSAFE_IN_TARGET.test(value)) {
			throw new NodeOperationError(
				this.getNode(),
				`The ${label} value contains characters that are not allowed`,
				{
					description:
						'Pick the value from the list, or supply an ID, e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
				},
			);
		}
	}

	return requestOptions;
}

/** Reads the user whose picker is being opened, as a path-safe segment. */
function currentUser(this: ILoadOptionsFunctions): string {
	const user = this.getCurrentNodeParameter('user', { extractValue: true });

	if (!isFilled(user) || UNSAFE_IN_TARGET.test(user)) {
		throw new NodeOperationError(this.getNode(), 'Choose a user first', {
			description: 'The methods on offer are the ones registered to that user',
		});
	}

	return user;
}

/**
 * Lists the methods of the selected type registered to the selected user.
 *
 * These collections are small and support no `$filter`, so the search term is matched
 * against the labels here rather than sent to Graph.
 */
export async function getAuthenticationMethods(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const user = currentUser.call(this);
	const methodType = this.getCurrentNodeParameter('methodType');

	if (!isFilled(methodType) || METHOD_TYPE_LABELS[methodType] === undefined) {
		throw new NodeOperationError(this.getNode(), 'Choose a method type first');
	}

	const response = (await microsoftApiRequest.call(
		this,
		'GET',
		`/users/${user}/authentication/${methodType}`,
	)) as { value?: IDataObject[] };

	const term = filter?.toLowerCase();
	const results: INodeListSearchItems[] = (response.value ?? [])
		.map((method) => ({ name: describeMethod(method, methodType), value: String(method.id) }))
		.filter((item) => !term || item.name.toLowerCase().includes(term))
		.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

	return { results };
}

/**
 * Builds a password that satisfies Entra's complexity rule — three of the four character
 * classes — by taking one character from each and filling the rest at random.
 */
export function generatePassword(length: number): string {
	const classes = Object.values(PASSWORD_ALPHABET);
	const alphabet = classes.join('');
	const characters = classes.map((set) => set[randomInt(set.length)]);

	while (characters.length < length) {
		characters.push(alphabet[randomInt(alphabet.length)]);
	}

	// Without the shuffle the first four positions would always be a lowercase letter, an
	// uppercase letter, a digit and a symbol, in that order.
	for (let index = characters.length - 1; index > 0; index--) {
		const swap = randomInt(index + 1);
		[characters[index], characters[swap]] = [characters[swap], characters[index]];
	}

	return characters.join('');
}

interface ResetPasswordOptions {
	password?: string;
	passwordLength?: number;
	forceChangePassword?: string;
}

/**
 * Sets a new password on the user through `passwordProfile`.
 *
 * Graph does have a dedicated password reset —
 * `POST /users/{id}/authentication/methods/{id}/resetPassword` — but it supports delegated
 * access only, so an app-only credential like this node's can never call it. Updating
 * **passwordProfile** is the app-only equivalent: it writes to Entra ID and, where password
 * writeback is configured, on to on-premises AD.
 *
 * Programmatic rather than declarative so that a generated password can be handed back to
 * the workflow — a declarative PATCH answers 204 with nothing in it.
 */
export async function executeResetPassword(
	this: IExecuteFunctions,
): Promise<INodeExecutionData[][]> {
	const items = this.getInputData();
	const results: INodeExecutionData[] = [];

	for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
		try {
			const user = this.getNodeParameter('user', itemIndex, '', { extractValue: true }) as string;
			if (!isFilled(user)) {
				throw new NodeOperationError(this.getNode(), 'No user was given', { itemIndex });
			}
			if (UNSAFE_IN_TARGET.test(user)) {
				throw new NodeOperationError(
					this.getNode(),
					'The user ID contains characters that are not allowed',
					{
						itemIndex,
						description:
							'Use the object ID, e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315, or the userPrincipalName',
					},
				);
			}

			const options = this.getNodeParameter('options', itemIndex, {}) as ResetPasswordOptions;
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
				passwordProfile.forceChangePasswordNextSignIn =
					forceChange === 'forceChangePasswordNextSignIn';
			}

			// Serial: a password reset is a directory write, and Entra rejects writes that
			// overlap within a tenant.
			// eslint-disable-next-line no-await-in-loop
			await microsoftApiRequest.call(this, 'PATCH', `/users/${user}`, { passwordProfile });

			results.push({
				json: {
					id: user,
					passwordReset: true,
					// Carried out so the workflow can deliver it: a generated password exists
					// nowhere else, and Graph never reads one back.
					password,
					generated: supplied === undefined,
					forceChangePasswordNextSignIn: passwordProfile.forceChangePasswordNextSignIn === true,
					forceChangePasswordNextSignInWithMfa:
						passwordProfile.forceChangePasswordNextSignInWithMfa === true,
				},
				pairedItem: { item: itemIndex },
			});
		} catch (error) {
			if (!this.continueOnFail()) {
				throw asNodeError.call(this, error, itemIndex);
			}
			results.push(errorItem(error, itemIndex));
		}
	}

	return [results];
}
