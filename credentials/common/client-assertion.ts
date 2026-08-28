import {
	createPrivateKey,
	createSign,
	randomUUID,
	X509Certificate,
	type KeyObject,
} from 'node:crypto';

import { OperationalError } from 'n8n-workflow';

import { formatPemBlock } from './format-pem-block';

// private_key_jwt (RFC 7521/7523): the client proves its identity with a JWT
// signed by its private key instead of a shared secret. The `x5t` header (SHA-1
// thumbprint of the certificate) tells the server which public key verifies it.
export const CLIENT_ASSERTION_TYPE = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer';

const ASSERTION_TTL_SECONDS = 300;

function base64url(input: Buffer | string): string {
	return Buffer.from(input).toString('base64url');
}

/**
 * Runs `parse` and reports the outcome instead of letting it throw.
 *
 * Callers raise their own error on the failure branch rather than from inside a
 * `catch`: this code runs during credential validation, outside any node execution,
 * so there is no node handle to build a `NodeApiError`/`NodeOperationError` from, and
 * throwing anything else from a catch block is not allowed.
 */
function tryParse<T>(parse: () => T): { ok: true; value: T } | { ok: false; error: unknown } {
	try {
		return { ok: true, value: parse() };
	} catch (error) {
		return { ok: false, error };
	}
}

function certificateThumbprint(certificate: string): string {
	const parsed = tryParse(() => new X509Certificate(formatPemBlock(certificate)));
	if (!parsed.ok) {
		throw new OperationalError(
			'The Certificate field must contain a PEM certificate (-----BEGIN CERTIFICATE-----).',
			{ cause: parsed.error },
		);
	}
	return Buffer.from(parsed.value.fingerprint.replace(/:/g, ''), 'hex').toString('base64url');
}

export interface BuildClientAssertionOptions {
	clientId: string;
	/** Token endpoint; used as the JWT `aud`. */
	accessTokenUri: string;
	/** RSA private key (PEM). Signing is RS256-only; EC/Ed25519 keys are not supported. */
	privateKey: string;
	certificate: string;
}

export function buildClientAssertion(options: BuildClientAssertionOptions): string {
	const now = Math.floor(Date.now() / 1000);
	const header = { alg: 'RS256', typ: 'JWT', x5t: certificateThumbprint(options.certificate) };
	const payload = {
		aud: options.accessTokenUri,
		iss: options.clientId,
		sub: options.clientId,
		jti: randomUUID(),
		iat: now,
		nbf: now,
		exp: now + ASSERTION_TTL_SECONDS,
	};

	const parsedKey = tryParse(() => createPrivateKey(formatPemBlock(options.privateKey)));
	if (!parsedKey.ok) {
		throw new OperationalError(
			'The Private Key field must contain a PEM private key (-----BEGIN PRIVATE KEY-----).',
			{ cause: parsedKey.error },
		);
	}
	const privateKey: KeyObject = parsedKey.value;

	// `createSign('RSA-SHA256')` also signs EC/Ed25519 keys, producing a signature
	// that contradicts the pinned `alg: RS256` header. Reject non-RSA keys up front.
	if (privateKey.asymmetricKeyType !== 'rsa') {
		throw new OperationalError('Certificate authentication requires an RSA private key');
	}

	const signingInput = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(payload))}`;
	const signature = createSign('RSA-SHA256').update(signingInput).sign(privateKey);
	return `${signingInput}.${base64url(signature)}`;
}
