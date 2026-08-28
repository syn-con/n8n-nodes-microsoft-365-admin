import { NodeApiError, type INode } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import { graphError, resolveGraphError, throwGraphError, type GraphErrorContext } from './errors';

const NODE = { name: 'Microsoft 365 Admin', type: 'microsoft365Admin' } as unknown as INode;

/**
 * The rules used to run as a declarative `postReceive` and needed a whole execution
 * context; they now take the three facts they actually read, so a case is one object.
 */
function context(
	resource: string,
	operation: string,
	params: Record<string, string> = {},
): GraphErrorContext {
	return {
		resource,
		operation,
		getParameter: (name, fallback) => params[name] ?? fallback ?? '',
	};
}

function body(error: unknown) {
	return { error };
}

const NOT_FOUND = { code: 'Request_ResourceNotFound', message: 'not found' };

describe('graphError', () => {
	it('reads the code and message out of the envelope', () => {
		expect(graphError(body({ code: 'X', message: 'm' }))).toEqual({
			code: 'X',
			message: 'm',
			details: undefined,
		});
	});

	it('copes with a gateway failure that carries no envelope', () => {
		expect(graphError('<html>502</html>')).toEqual({ code: '', message: '', details: undefined });
		expect(graphError(undefined)).toEqual({ code: '', message: '', details: undefined });
	});
});

describe('resolveGraphError', () => {
	it.each([
		['group', 'delete', /Group to Delete/],
		['group', 'get', /Group to Get/],
		['group', 'update', /Group to Update/],
		['user', 'delete', /User to Delete/],
		['user', 'get', /User to Get/],
		['user', 'update', /User to Update/],
	])('names the parameter at fault for %s.%s not-found', (resource, operation, match) => {
		const resolution = resolveGraphError(body(NOT_FOUND), context(resource, operation));
		expect(resolution?.description).toMatch(match);
	});

	it('explains that a user is already a member when adding a duplicate', () => {
		const resolution = resolveGraphError(
			body({
				code: 'Request_BadRequest',
				message:
					"One or more added object references already exist for the following modified properties: 'members'.",
			}),
			context('user', 'addGroup'),
		);
		expect(resolution?.message).toMatch(/already in the group/);
	});

	it('explains that a user is not a member when removing a non-member', () => {
		const resolution = resolveGraphError(body(NOT_FOUND), context('user', 'removeGroup'));
		expect(resolution?.message).toMatch(/not in the group/);
	});

	it('blames the group when an addGroup miss mentions the group ID', () => {
		const resolution = resolveGraphError(
			body({ code: 'Request_ResourceNotFound', message: 'no group-1 here' }),
			context('user', 'addGroup', { 'group.value': 'group-1' }),
		);
		expect(resolution?.description).toMatch(/'Group'/);
	});

	it('blames the user when an addGroup miss does not mention the group ID', () => {
		const resolution = resolveGraphError(
			body({ code: 'Request_ResourceNotFound', message: 'missing directory object' }),
			context('user', 'addGroup', { 'group.value': 'group-1' }),
		);
		expect(resolution?.description).toMatch(/'User to Add'/);
	});

	it('explains an unusable member reference on removeGroup', () => {
		const resolution = resolveGraphError(
			body({
				code: 'Request_UnsupportedQuery',
				message: "Unsupported referenced-object resource identifier for link property 'members'.",
			}),
			context('user', 'removeGroup'),
		);
		expect(resolution?.message).toMatch(/user ID is invalid/);
	});

	it.each([
		['group', 'The group ID is invalid'],
		['user', 'The user ID is invalid'],
	])('reports an invalid object identifier for %s', (resource, message) => {
		const resolution = resolveGraphError(
			body({ code: 'Request_BadRequest', message: 'Invalid object identifier "abc"' }),
			context(resource, 'get', { 'group.value': 'group-1' }),
		);
		expect(resolution?.message).toBe(message);
	});

	it.each(['ObjectConflict', 'ConflictingObjects'])(
		'reports an already-existing resource for %s',
		(code) => {
			const resolution = resolveGraphError(
				body({
					code: 'Request_BadRequest',
					message: 'conflict',
					details: [{ code, message: 'conflict' }],
				}),
				context('group', 'create'),
			);
			expect(resolution?.message).toMatch(/group already exists/);
		},
	);

	it('leaves an unrecognised failure to Graph’s own wording', () => {
		expect(
			resolveGraphError(
				body({ code: 'Authorization_RequestDenied', message: 'denied' }),
				context('license', 'assign'),
			),
		).toBeUndefined();
	});

	it('no longer treats the empty-payload 400 specially', () => {
		// The operations skip a no-op PATCH instead of sending it and swallowing the error,
		// so this is now an ordinary failure rather than a rule.
		expect(
			resolveGraphError(
				body({ code: 'BadRequest', message: 'Empty Payload. JSON content expected.' }),
				context('user', 'update'),
			),
		).toBeUndefined();
	});
});

describe('throwGraphError', () => {
	it('always raises a NodeApiError', () => {
		expect(() =>
			throwGraphError(NODE, { statusCode: 404, body: body(NOT_FOUND) }, context('user', 'get'), 3),
		).toThrow(NodeApiError);
	});

	it('carries the reworded message when a rule matched', () => {
		expect(() =>
			throwGraphError(NODE, { statusCode: 404, body: body(NOT_FOUND) }, context('user', 'get')),
		).toThrow(/user doesn't match/);
	});
});
