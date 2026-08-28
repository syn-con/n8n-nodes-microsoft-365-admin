import type { DateTime } from 'luxon';
import { NodeOperationError, type IDataObject, type INode } from 'n8n-workflow';

/** Only these characters are accepted in a UPN. */
const UPN_ALLOWED = /^[A-Za-z0-9'._\-!#^~@]+$/;

/**
 * Graph refuses these unless they travel in a PATCH of their own, with no other property
 * alongside them.
 */
const SEPARATE_ONLY_PROPERTIES = [
	'aboutMe',
	'birthday',
	'interests',
	'mySite',
	'pastProjects',
	'responsibilities',
	'schools',
	'skills',
];

/** Rewrites the collected fields into the shapes Graph accepts. */
export function toGraphUserBody(fields: IDataObject): IDataObject {
	const body: IDataObject = { ...fields };

	for (const key of ['birthday', 'employeeHireDate', 'employeeLeaveDateTime']) {
		if (body[key]) {
			body[key] = (body[key] as DateTime).toUTC().toISO();
		}
	}
	if (body.businessPhones) {
		body.businessPhones = [body.businessPhones as string];
	}
	if (body.employeeOrgData) {
		body.employeeOrgData = (body.employeeOrgData as IDataObject).employeeOrgValues;
	}
	if (body.passwordPolicies) {
		body.passwordPolicies = (body.passwordPolicies as string[]).join(',');
	}
	// forceChangePasswordNextSignInWithMfa doesn't seem to take effect when providing it in
	// the initial create request, so it is applied here instead. Only the two known flags are
	// honoured, so an expression cannot name an arbitrary property of the password profile.
	if (body.forceChangePassword) {
		const flag = body.forceChangePassword as string;
		if (['forceChangePasswordNextSignIn', 'forceChangePasswordNextSignInWithMfa'].includes(flag)) {
			body.passwordProfile ??= {};
			(body.passwordProfile as IDataObject)[flag] = true;
		}
		delete body.forceChangePassword;
	}

	return body;
}

/** Moves the properties Graph only accepts alone out of `body` and into their own object. */
export function splitSeparateOnly(body: IDataObject): IDataObject {
	const separateBody: IDataObject = {};

	for (const key of SEPARATE_ONLY_PROPERTIES) {
		if (key in body) {
			separateBody[key] = body[key];
			delete body[key];
		}
	}

	return separateBody;
}

/**
 * Checks the fields Graph would otherwise reject with a message that names neither the
 * parameter nor the limit it broke.
 */
export function validateUserFields(node: INode, fields: IDataObject, itemIndex: number): void {
	const companyName = (fields.companyName ?? '') as string;
	const employeeId = (fields.employeeId ?? '') as string;
	const userPrincipalName = (fields.userPrincipalName ?? '') as string;

	if (companyName.length > 64) {
		throw new NodeOperationError(node, "'Company Name' should have a maximum length of 64", {
			itemIndex,
		});
	}
	if (employeeId.length > 16) {
		throw new NodeOperationError(node, "'Employee ID' should have a maximum length of 16", {
			itemIndex,
		});
	}
	if (userPrincipalName && !UPN_ALLOWED.test(userPrincipalName)) {
		throw new NodeOperationError(
			node,
			"Only the following characters are allowed for 'User Principal Name': A-Z, a-z, 0-9, ' . - _ ! # ^ ~",
			{ itemIndex },
		);
	}
}

/** The same check for the required UPN on Create, which is not inside a collection. */
export function validateUserPrincipalName(
	node: INode,
	userPrincipalName: string,
	itemIndex: number,
): void {
	if (!UPN_ALLOWED.test(userPrincipalName)) {
		throw new NodeOperationError(
			node,
			"Only the following characters are allowed for 'User Principal Name': A-Z, a-z, 0-9, ' . - _ ! # ^ ~",
			{ itemIndex },
		);
	}
}
