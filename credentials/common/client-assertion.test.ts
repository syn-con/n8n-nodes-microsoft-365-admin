import { createVerify, generateKeyPairSync, X509Certificate } from 'node:crypto';
import { describe, expect, it } from 'vitest';

import { buildClientAssertion, CLIENT_ASSERTION_TYPE } from './client-assertion';
import { certificate as CERTIFICATE, privateKey as PRIVATE_KEY } from './client-assertion.fixtures';

const base = {
	clientId: 'a09519e2-abb5-4efa-a21f-41c86831d152',
	accessTokenUri: 'https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token',
	privateKey: PRIVATE_KEY,
	certificate: CERTIFICATE,
};

function decodeSegment(segment: string): Record<string, unknown> {
	return JSON.parse(Buffer.from(segment, 'base64url').toString()) as Record<string, unknown>;
}

describe('CLIENT_ASSERTION_TYPE', () => {
	it('is the RFC 7523 jwt-bearer identifier', () => {
		expect(CLIENT_ASSERTION_TYPE).toBe('urn:ietf:params:oauth:client-assertion-type:jwt-bearer');
	});
});

describe('buildClientAssertion', () => {
	it('produces a three-segment JWT', () => {
		expect(buildClientAssertion(base).split('.')).toHaveLength(3);
	});

	it('pins RS256 and carries the certificate thumbprint as x5t', () => {
		const [encodedHeader] = buildClientAssertion(base).split('.');
		const header = decodeSegment(encodedHeader);

		expect(header.alg).toBe('RS256');
		expect(header.typ).toBe('JWT');

		const expectedThumbprint = Buffer.from(
			new X509Certificate(CERTIFICATE).fingerprint.replace(/:/g, ''),
			'hex',
		).toString('base64url');
		expect(header.x5t).toBe(expectedThumbprint);
	});

	it('sets iss/sub to the client ID and aud to the token endpoint', () => {
		const payload = decodeSegment(buildClientAssertion(base).split('.')[1]);

		expect(payload.iss).toBe(base.clientId);
		expect(payload.sub).toBe(base.clientId);
		expect(payload.aud).toBe(base.accessTokenUri);
	});

	it('expires five minutes after issuance and is valid immediately', () => {
		const payload = decodeSegment(buildClientAssertion(base).split('.')[1]);

		expect(payload.exp as number).toBe((payload.iat as number) + 300);
		expect(payload.nbf).toBe(payload.iat);
	});

	it('uses a fresh jti per assertion so tokens cannot be replayed', () => {
		const first = decodeSegment(buildClientAssertion(base).split('.')[1]);
		const second = decodeSegment(buildClientAssertion(base).split('.')[1]);

		expect(first.jti).not.toBe(second.jti);
	});

	it('signs the header.payload input verifiably with the private key', () => {
		const assertion = buildClientAssertion(base);
		const [header, payload, signature] = assertion.split('.');

		const { publicKey } = new X509Certificate(CERTIFICATE);
		const verified = createVerify('RSA-SHA256')
			.update(`${header}.${payload}`)
			.verify(publicKey, Buffer.from(signature, 'base64url'));

		expect(verified).toBe(true);
	});

	it('accepts a flattened single-line private key', () => {
		const flattened = PRIVATE_KEY.replace(/\n/g, '\\n');
		expect(() => buildClientAssertion({ ...base, privateKey: flattened })).not.toThrow();
	});

	it('rejects a certificate that is not valid PEM', () => {
		expect(() => buildClientAssertion({ ...base, certificate: 'not-a-cert' })).toThrow(
			/must contain a PEM certificate/,
		);
	});

	it('rejects a private key that is not valid PEM', () => {
		expect(() => buildClientAssertion({ ...base, privateKey: 'not-a-key' })).toThrow(
			/must contain a PEM private key/,
		);
	});

	it('rejects a non-RSA private key, which would contradict the pinned RS256 header', () => {
		const { privateKey } = generateKeyPairSync('ec', { namedCurve: 'prime256v1' });
		const ecPem = privateKey.export({ type: 'pkcs8', format: 'pem' }).toString();

		expect(() => buildClientAssertion({ ...base, privateKey: ecPem })).toThrow(
			/requires an RSA private key/,
		);
	});
});
