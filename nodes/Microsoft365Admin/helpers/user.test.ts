import { NodeOperationError, type IDataObject, type INode } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import {
	splitSeparateOnly,
	toGraphUserBody,
	validateUserFields,
	validateUserPrincipalName,
} from './user';

const NODE = { name: 'Microsoft 365 Admin', type: 'microsoft365Admin' } as unknown as INode;

/** Stands in for a luxon DateTime; only these two methods are used. */
const dateTime = (iso: string) => ({ toUTC: () => ({ toISO: () => iso }) });

describe('toGraphUserBody', () => {
	it('renders the date fields as UTC ISO strings', () => {
		const body = toGraphUserBody({
			birthday: dateTime('2000-01-01T00:00:00.000Z'),
			employeeHireDate: dateTime('2020-06-01T00:00:00.000Z'),
			employeeLeaveDateTime: dateTime('2026-01-01T00:00:00.000Z'),
		} as unknown as IDataObject);

		expect(body).toEqual({
			birthday: '2000-01-01T00:00:00.000Z',
			employeeHireDate: '2020-06-01T00:00:00.000Z',
			employeeLeaveDateTime: '2026-01-01T00:00:00.000Z',
		});
	});

	it('wraps a single business phone in the array Graph expects', () => {
		expect(toGraphUserBody({ businessPhones: '+370 600 00000' })).toEqual({
			businessPhones: ['+370 600 00000'],
		});
	});

	it('unwraps the fixed-collection around employeeOrgData', () => {
		expect(
			toGraphUserBody({ employeeOrgData: { employeeOrgValues: { division: 'R&D' } } }),
		).toEqual({ employeeOrgData: { division: 'R&D' } });
	});

	it('joins password policies into the comma-separated string Graph takes', () => {
		expect(
			toGraphUserBody({ passwordPolicies: ['DisablePasswordExpiration', 'DisableStrongPassword'] }),
		).toEqual({ passwordPolicies: 'DisablePasswordExpiration,DisableStrongPassword' });
	});

	it.each(['forceChangePasswordNextSignIn', 'forceChangePasswordNextSignInWithMfa'])(
		'folds %s into the password profile',
		(flag) => {
			expect(toGraphUserBody({ forceChangePassword: flag })).toEqual({
				passwordProfile: { [flag]: true },
			});
		},
	);

	it('ignores a force-change value it does not recognise', () => {
		// Otherwise an expression could name an arbitrary property of the password profile.
		expect(toGraphUserBody({ forceChangePassword: 'somethingElse' })).toEqual({});
	});

	it('leaves anything else untouched', () => {
		expect(toGraphUserBody({ displayName: 'Ada', jobTitle: 'Engineer' })).toEqual({
			displayName: 'Ada',
			jobTitle: 'Engineer',
		});
	});

	it('does not mutate the fields it was given', () => {
		const fields = { businessPhones: '+1' };
		toGraphUserBody(fields);
		expect(fields).toEqual({ businessPhones: '+1' });
	});
});

describe('splitSeparateOnly', () => {
	it('moves the properties Graph insists on receiving alone', () => {
		const body: IDataObject = { displayName: 'Ada', aboutMe: 'hi', skills: ['x'] };
		const separate = splitSeparateOnly(body);

		expect(separate).toEqual({ aboutMe: 'hi', skills: ['x'] });
		expect(body).toEqual({ displayName: 'Ada' });
	});

	it('returns nothing when no such property is present', () => {
		const body: IDataObject = { displayName: 'Ada' };
		expect(splitSeparateOnly(body)).toEqual({});
		expect(body).toEqual({ displayName: 'Ada' });
	});
});

describe('validateUserFields', () => {
	it('accepts fields within Graph’s limits', () => {
		expect(() =>
			validateUserFields(NODE, { companyName: 'Acme', employeeId: 'E1' }, 0),
		).not.toThrow();
	});

	it('rejects a company name over 64 characters', () => {
		expect(() => validateUserFields(NODE, { companyName: 'a'.repeat(65) }, 0)).toThrow(
			/maximum length of 64/,
		);
	});

	it('rejects an employee ID over 16 characters', () => {
		expect(() => validateUserFields(NODE, { employeeId: 'a'.repeat(17) }, 0)).toThrow(
			/maximum length of 16/,
		);
	});

	it('rejects a UPN with characters Entra will not take', () => {
		expect(() => validateUserFields(NODE, { userPrincipalName: 'a b@contoso.com' }, 0)).toThrow(
			/User Principal Name/,
		);
	});

	it('carries the item index on the error', () => {
		try {
			validateUserFields(NODE, { employeeId: 'a'.repeat(17) }, 4);
			expect.unreachable('should have thrown');
		} catch (error) {
			expect((error as NodeOperationError).context.itemIndex).toBe(4);
		}
	});
});

describe('validateUserPrincipalName', () => {
	it.each(['NathanSmith@contoso.com', "o'brien.a@contoso.com", 'a_b-c!#^~@contoso.com'])(
		'accepts %s',
		(upn) => {
			expect(() => validateUserPrincipalName(NODE, upn, 0)).not.toThrow();
		},
	);

	it.each(['has space@contoso.com', 'slash/@contoso.com', 'quote"@contoso.com'])(
		'rejects %s',
		(upn) => {
			expect(() => validateUserPrincipalName(NODE, upn, 0)).toThrow(/User Principal Name/);
		},
	);
});
