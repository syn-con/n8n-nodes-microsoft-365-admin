import { createSign, generateKeyPairSync } from 'node:crypto';

/**
 * Generates a throwaway RSA key pair and a matching self-signed X.509 certificate
 * at test time, so no private key material is ever committed to the repository.
 *
 * `buildClientAssertion` parses the certificate with `node:crypto`'s `X509Certificate`
 * to derive the `x5t` thumbprint, so a structurally valid DER certificate is required —
 * a placeholder would make those tests pass without testing anything. Node cannot
 * generate certificates, and shelling out to `openssl` would not be portable across
 * CI images, so the certificate is assembled here with a minimal DER encoder.
 */

// --- Minimal DER encoding -------------------------------------------------

function derLength(length: number): Buffer {
	if (length < 0x80) {
		return Buffer.from([length]);
	}

	const bytes: number[] = [];
	let remaining = length;
	while (remaining > 0) {
		bytes.unshift(remaining & 0xff);
		remaining >>>= 8;
	}
	return Buffer.from([0x80 | bytes.length, ...bytes]);
}

function tlv(tag: number, ...content: Buffer[]): Buffer {
	const body = Buffer.concat(content);
	return Buffer.concat([Buffer.from([tag]), derLength(body.length), body]);
}

const sequence = (...content: Buffer[]) => tlv(0x30, ...content);
const set = (...content: Buffer[]) => tlv(0x31, ...content);
const objectIdentifier = (hex: string) => tlv(0x06, Buffer.from(hex, 'hex'));
const printableString = (value: string) => tlv(0x13, Buffer.from(value, 'ascii'));
const utcTime = (value: string) => tlv(0x17, Buffer.from(value, 'ascii'));
const nullValue = Buffer.from([0x05, 0x00]);

/** DER INTEGERs are signed, so a leading 0x00 is needed when the top bit is set. */
function integer(value: Buffer): Buffer {
	const padded = value[0] & 0x80 ? Buffer.concat([Buffer.from([0x00]), value]) : value;
	return tlv(0x02, padded);
}

/** BIT STRING with an explicit "0 unused bits" prefix. */
const bitString = (value: Buffer) => tlv(0x03, Buffer.concat([Buffer.from([0x00]), value]));

// sha256WithRSAEncryption (1.2.840.113549.1.1.11) and commonName (2.5.4.3)
const SHA256_WITH_RSA = '2a864886f70d01010b';
const COMMON_NAME = '550403';

function pem(label: string, der: Buffer): string {
	const body = der.toString('base64').match(/.{1,64}/g) ?? [];
	return `-----BEGIN ${label}-----\n${body.join('\n')}\n-----END ${label}-----`;
}

// --- Certificate ----------------------------------------------------------

function selfSign(): { privateKey: string; certificate: string } {
	const pair = generateKeyPairSync('rsa', { modulusLength: 2048 });

	const privateKeyPem = pair.privateKey.export({ type: 'pkcs8', format: 'pem' }).toString();
	// An SPKI export is already a complete SubjectPublicKeyInfo, so it drops straight in.
	const subjectPublicKeyInfo = pair.publicKey.export({ type: 'spki', format: 'der' }) as Buffer;

	const algorithm = sequence(objectIdentifier(SHA256_WITH_RSA), nullValue);
	const name = sequence(
		set(sequence(objectIdentifier(COMMON_NAME), printableString('m365-admin-test'))),
	);
	// Fixed dates keep the fixture deterministic; nothing validates the window.
	const validity = sequence(utcTime('240101000000Z'), utcTime('341231235959Z'));

	// v1 TBSCertificate: the version field is optional and defaults to v1.
	const tbsCertificate = sequence(
		integer(Buffer.from([0x2a])),
		algorithm,
		name,
		validity,
		name,
		subjectPublicKeyInfo,
	);

	const signature = createSign('RSA-SHA256').update(tbsCertificate).sign(pair.privateKey);
	const certificateDer = sequence(tbsCertificate, algorithm, bitString(signature));

	return { privateKey: privateKeyPem, certificate: pem('CERTIFICATE', certificateDer) };
}

export const { privateKey, certificate } = selfSign();
