import { describe, expect, it } from 'vitest';

import { annotateMethod, describeMethod, generatePassword, methodTypeOf } from './authentication';

describe('method types', () => {
	it.each([
		['#microsoft.graph.phoneAuthenticationMethod', 'phoneMethods'],
		[
			'#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
			'microsoftAuthenticatorMethods',
		],
		['#microsoft.graph.fido2AuthenticationMethod', 'fido2Methods'],
		['#microsoft.graph.passwordAuthenticationMethod', 'passwordMethods'],
		// Graph is not always consistent about the leading hash.
		['microsoft.graph.emailAuthenticationMethod', 'emailMethods'],
	])('maps %s to the collection that addresses it', (odataType, expected) => {
		expect(methodTypeOf(odataType)).toBe(expected);
	});

	it.each([undefined, null, 42, '#microsoft.graph.somethingNewAuthenticationMethod'])(
		'has nothing to say about %s',
		(odataType) => {
			expect(methodTypeOf(odataType)).toBeUndefined();
		},
	);
});

describe('method labels', () => {
	it('uses the name a security key was registered under', () => {
		expect(describeMethod({ displayName: 'Red key', model: 'NFC key' }, 'fido2Methods')).toBe(
			'Red key (NFC key)',
		);
	});

	it('names a phone by its number and type', () => {
		expect(
			describeMethod({ phoneNumber: '+370 600 00000', phoneType: 'mobile' }, 'phoneMethods'),
		).toBe('+370 600 00000 (mobile)');
	});

	it('falls back to the method type when Graph gives no name at all', () => {
		expect(
			describeMethod(
				{ id: 'tap-1', createdDateTime: '2026-08-21T09:00:00Z' },
				'temporaryAccessPassMethods',
			),
		).toBe('Temporary Access Pass (2026-08-21T09:00:00Z)');
	});

	it('falls back to the ID when even the type is unknown', () => {
		expect(describeMethod({ id: 'method-1' })).toBe('method-1');
	});
});

describe('annotating a method', () => {
	const phone = {
		'@odata.type': '#microsoft.graph.phoneAuthenticationMethod',
		id: 'p1',
		phoneNumber: '+1',
	};
	const password = { '@odata.type': '#microsoft.graph.passwordAuthenticationMethod', id: 'pw' };
	const unknown = { '@odata.type': '#microsoft.graph.newFangledAuthenticationMethod', id: 'x' };

	it('adds the collection a Delete Method step needs', () => {
		expect(annotateMethod(phone).methodType).toBe('phoneMethods');
	});

	it('marks the password method as one that cannot be deleted', () => {
		expect(annotateMethod(password)).toMatchObject({
			methodType: 'passwordMethods',
			deletable: false,
		});
		expect(annotateMethod(phone).deletable).toBe(true);
	});

	it('reports an unrecognized method type rather than guessing one', () => {
		expect(annotateMethod(unknown)).toMatchObject({ methodType: null, deletable: false });
	});

	it('keeps everything Graph returned', () => {
		expect(annotateMethod(phone)).toMatchObject({ id: 'p1', phoneNumber: '+1' });
	});
});

describe('generated passwords', () => {
	it('is as long as asked', () => {
		expect(generatePassword(24)).toHaveLength(24);
	});

	it('carries all four character classes, so Entra accepts it', () => {
		for (let attempt = 0; attempt < 50; attempt++) {
			const password = generatePassword(12);
			expect(password).toMatch(/[a-z]/);
			expect(password).toMatch(/[A-Z]/);
			expect(password).toMatch(/[0-9]/);
			expect(password).toMatch(/[!#$%&*+\-=?@^_]/);
		}
	});

	it('leaves out the characters that get misread', () => {
		for (let attempt = 0; attempt < 50; attempt++) {
			expect(generatePassword(32)).not.toMatch(/[lIO01]/);
		}
	});

	it('does not repeat itself', () => {
		const passwords = new Set(Array.from({ length: 20 }, () => generatePassword(16)));
		expect(passwords.size).toBe(20);
	});
});
