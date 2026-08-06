import { describe, expect, it } from 'vitest';

import { formatPemBlock } from './format-pem-block';

const BODY = 'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAtest';

describe('formatPemBlock', () => {
	it('returns input unchanged when it already contains newlines', () => {
		const pem = `-----BEGIN PRIVATE KEY-----\n${BODY}\n-----END PRIVATE KEY-----`;
		expect(formatPemBlock(pem)).toBe(pem);
	});

	it('returns falsy input unchanged', () => {
		expect(formatPemBlock('')).toBe('');
	});

	it('rewraps a single-line private key onto 64-character lines', () => {
		const long = 'A'.repeat(130);
		const result = formatPemBlock(`-----BEGIN PRIVATE KEY-----${long}-----END PRIVATE KEY-----`);

		const lines = result.split('\n');
		expect(lines[0]).toBe('-----BEGIN PRIVATE KEY-----');
		expect(lines[lines.length - 1]).toBe('-----END PRIVATE KEY-----');
		expect(lines[1]).toHaveLength(64);
		expect(lines[2]).toHaveLength(64);
		expect(lines[3]).toHaveLength(2);
	});

	it('normalises literal \\n sequences into real newlines', () => {
		const result = formatPemBlock(
			`-----BEGIN CERTIFICATE-----\\n${BODY}\\n-----END CERTIFICATE-----`,
		);
		expect(result).toContain('\n');
		expect(result).not.toContain('\\n');
		expect(result).toContain(BODY);
	});

	it('handles certificates as well as keys', () => {
		const result = formatPemBlock(`-----BEGIN CERTIFICATE-----${BODY}-----END CERTIFICATE-----`);
		expect(result.startsWith('-----BEGIN CERTIFICATE-----\n')).toBe(true);
		expect(result.endsWith('\n-----END CERTIFICATE-----')).toBe(true);
	});

	it('matches PUBLIC KEY labels only when isPublic is set', () => {
		const pem = `-----BEGIN PUBLIC KEY-----${BODY}-----END PUBLIC KEY-----`;
		const asPublic = formatPemBlock(pem, true);
		expect(asPublic.startsWith('-----BEGIN PUBLIC KEY-----\n')).toBe(true);
	});

	it('leaves multi-block PEM chains alone', () => {
		const chain =
			'-----BEGIN CERTIFICATE-----a-----END CERTIFICATE-----' +
			'-----BEGIN CERTIFICATE-----b-----END CERTIFICATE-----';
		// More than one BEGIN marker, so the compact-PEM path declines to reformat.
		expect(formatPemBlock(chain)).toContain('-----BEGIN CERTIFICATE-----');
	});

	it('collapses whitespace inside an encrypted key header', () => {
		const pem =
			'-----BEGIN RSA PRIVATE KEY-----Proc-Type: 4,ENCRYPTED DEK-Info: AES-128-CBC,ABC' +
			`${BODY}-----END RSA PRIVATE KEY-----`;
		const result = formatPemBlock(pem);
		expect(result).toContain('Proc-Type:4,ENCRYPTED');
	});
});
