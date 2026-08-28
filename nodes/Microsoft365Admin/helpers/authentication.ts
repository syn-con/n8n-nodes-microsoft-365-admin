import { randomInt } from 'node:crypto';

import type { IDataObject } from 'n8n-workflow';

import { isFilled } from './utils';

/**
 * The authentication methods API keeps every method type in its own collection, so an ID
 * on its own does not address a method — the type decides the URL segment, and Graph never
 * reports that segment: `GET /authentication/methods` answers with an `@odata.type`. This
 * module bridges the two, so Get Many Methods can be piped straight into Delete Method.
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
export const METHOD_TYPE_LABELS: Record<string, string> = {
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

export const DEFAULT_PASSWORD_LENGTH = 16;
export const MIN_PASSWORD_LENGTH = 8;
export const MAX_PASSWORD_LENGTH = 256;

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
export function annotateMethod(method: IDataObject): IDataObject {
	const methodType = methodTypeOf(method['@odata.type']);

	return {
		...method,
		methodType: methodType ?? null,
		methodName: describeMethod(method, methodType),
		deletable: methodType !== undefined && !UNDELETABLE_METHOD_TYPES.includes(methodType),
	};
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
