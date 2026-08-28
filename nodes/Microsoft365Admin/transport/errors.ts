import { NodeApiError, type INode, type JsonObject } from 'n8n-workflow';

import type { GraphErrorBody } from '../helpers/interfaces';

/**
 * Graph reports a failure the same way whatever went wrong: a 4xx or 5xx carrying an error
 * code and a sentence written for whoever wrote the request, not for whoever is looking at
 * the workflow. This module turns the ones an operator can act on into a message that names
 * the parameter to fix, and lets the rest through as Graph phrased them.
 */

/** The ID format quoted back whenever an identifier is unusable. */
const ID_FORMAT = 'The ID should be in the format e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315';

/**
 * What a rule needs to know about the call that failed.
 *
 * The rules used to read this off `IExecuteSingleFunctions` while running as a declarative
 * `postReceive`. Programmatic operations resolve their parameters per item, so the caller
 * passes the values in rather than the whole execution context — which also makes the rule
 * table straightforward to test.
 */
export interface GraphErrorContext {
	resource: string;
	operation: string;
	/** Reads a node parameter for the failing item; used by rules that name the parameter at fault. */
	getParameter: (name: string, fallback?: string) => string;
}

/** A message to raise in place of Graph's own. */
export type ErrorResolution = { message: string; description: string };

interface ErrorRule {
	/** The Graph error code this rule answers. */
	code: string;
	/** Narrows a code that Graph uses for more than one thing. */
	when?: (message: string) => boolean;
	resolve: ErrorResolution | ((message: string, context: GraphErrorContext) => ErrorResolution);
}

/** The only thing that varies between the not-found messages is the parameter to blame. */
function notFound(resource: 'group' | 'user', parameterName: string): ErrorResolution {
	return {
		message: `The required ${resource} doesn't match any existing one`,
		description: `Double-check the value in the parameter '${parameterName}' and try again`,
	};
}

/** Keyed by `resource.operation`; the first matching rule wins. */
const OPERATION_RULES: Record<string, ErrorRule[]> = {
	'authentication.deleteMethod': [
		{
			code: 'Request_ResourceNotFound',
			resolve: {
				message: "The authentication method doesn't match any existing one",
				description:
					'It may have been removed already, or belong to another user. Run Get Many Methods for the user to see what is registered, and check that Method Type matches the method.',
			},
		},
	],
	'group.delete': [
		{ code: 'Request_ResourceNotFound', resolve: notFound('group', 'Group to Delete') },
	],
	'group.get': [{ code: 'Request_ResourceNotFound', resolve: notFound('group', 'Group to Get') }],
	'group.update': [
		{ code: 'Request_ResourceNotFound', resolve: notFound('group', 'Group to Update') },
	],
	'user.addGroup': [
		{
			code: 'Request_BadRequest',
			when: (message) =>
				message ===
				"One or more added object references already exist for the following modified properties: 'members'.",
			resolve: {
				message: 'The user is already in the group',
				description:
					'The specified user cannot be added to the group because they are already a member',
			},
		},
		{
			code: 'Request_ResourceNotFound',
			// Graph names whichever object it could not find, so the message decides the blame.
			resolve(message, context) {
				const group = context.getParameter('group.value');
				return message.includes(group)
					? notFound('group', 'Group')
					: notFound('user', 'User to Add');
			},
		},
	],
	'user.delete': [
		{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Delete') },
	],
	'user.get': [{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Get') }],
	'user.removeGroup': [
		{
			code: 'Request_ResourceNotFound',
			resolve: {
				message: 'The user is not in the group',
				description:
					'The specified user cannot be removed from the group because they are not a member of the group',
			},
		},
		{
			code: 'Request_UnsupportedQuery',
			when: (message) =>
				message ===
				"Unsupported referenced-object resource identifier for link property 'members'.",
			resolve: { message: 'The user ID is invalid', description: ID_FORMAT },
		},
	],
	'user.update': [
		{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Update') },
	],
};

/** Tried after the operation's own rules, for failures any operation can hit. */
const GENERIC_RULES: ErrorRule[] = [
	{
		code: 'Request_BadRequest',
		when: (message) => message.startsWith('Invalid object identifier'),
		resolve(message, context) {
			const group = context.getParameter('group.value', '');
			const parameterResource =
				context.resource === 'group' || (group !== '' && message.includes(group))
					? 'group'
					: 'user';

			return { message: `The ${parameterResource} ID is invalid`, description: ID_FORMAT };
		},
	},
];

/** Reads Graph's error envelope, which is absent on some gateway failures. */
export function graphError(body: unknown): GraphErrorBody {
	const error = (body as { error?: { code?: string; message?: string; details?: [] } })?.error;

	return { code: error?.code ?? '', message: error?.message ?? '', details: error?.details };
}

/** The operation's own rules are tried before the generic ones; the first match wins. */
function findRule(
	resource: string,
	operation: string,
	code: string,
	message: string,
): ErrorRule | undefined {
	const rules = [...(OPERATION_RULES[`${resource}.${operation}`] ?? []), ...GENERIC_RULES];

	return rules.find((rule) => rule.code === code && (!rule.when || rule.when(message)));
}

/**
 * Decides what a failed Graph response should surface as.
 *
 * `undefined` means no rule recognised it, so Graph's own wording stands.
 */
export function resolveGraphError(
	body: unknown,
	context: GraphErrorContext,
): ErrorResolution | undefined {
	const { code, message, details } = graphError(body);
	const rule = findRule(context.resource, context.operation, code, message);

	if (rule) {
		return typeof rule.resolve === 'function' ? rule.resolve(message, context) : rule.resolve;
	}

	if (details?.some((detail) => ['ObjectConflict', 'ConflictingObjects'].includes(detail.code))) {
		return {
			message: `The ${context.resource} already exists`,
			description: message,
		};
	}

	return undefined;
}

/** True when the status line is a Graph failure rather than a success. */
export function isErrorStatus(statusCode: number): boolean {
	return statusCode >= 400;
}

/**
 * Raises a failed Graph response as a node error, reworded where a rule recognises it.
 *
 * @throws always — the return type is `never` so callers need no unreachable fallthrough.
 */
export function throwGraphError(
	node: INode,
	response: { statusCode: number; body: unknown },
	context: GraphErrorContext,
	itemIndex?: number,
): never {
	const resolution = resolveGraphError(response.body, context);

	throw new NodeApiError(node, response as unknown as JsonObject, {
		...(resolution ?? {}),
		...(itemIndex === undefined ? {} : { itemIndex }),
	});
}
