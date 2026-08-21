import {
	NodeApiError,
	type IExecuteSingleFunctions,
	type IN8nHttpFullResponse,
	type INodeExecutionData,
	type JsonObject,
} from 'n8n-workflow';

/**
 * Graph reports a failure the same way whatever went wrong: a 4xx or 5xx carrying an error
 * code and a sentence written for whoever wrote the request, not for whoever is looking at
 * the workflow. This module turns the ones an operator can act on into a message that names
 * the parameter to fix, and lets the rest through as Graph phrased them.
 */

/** The ID format quoted back whenever an identifier is unusable. */
const ID_FORMAT = 'The ID should be in the format e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315';

/** n8n strips an empty body object, so a no-op update reaches Graph with no payload. */
const EMPTY_PAYLOAD = 'Empty Payload. JSON content expected.';

/** A message to raise in place of Graph's, or `ignore` to let the response through. */
type ErrorResolution = { message: string; description: string } | 'ignore';

interface ErrorRule {
	/** The Graph error code this rule answers. */
	code: string;
	/** Narrows a code that Graph uses for more than one thing. */
	when?: (message: string) => boolean;
	resolve:
		| ErrorResolution
		| ((this: IExecuteSingleFunctions, message: string, resource: string) => ErrorResolution);
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
		{ code: 'BadRequest', when: (message) => message === EMPTY_PAYLOAD, resolve: 'ignore' },
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
			resolve(message) {
				const group = this.getNodeParameter('group.value') as string;
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
		{ code: 'BadRequest', when: (message) => message === EMPTY_PAYLOAD, resolve: 'ignore' },
		{ code: 'Request_ResourceNotFound', resolve: notFound('user', 'User to Update') },
	],
};

/** Tried after the operation's own rules, for failures any operation can hit. */
const GENERIC_RULES: ErrorRule[] = [
	{
		code: 'Request_BadRequest',
		when: (message) => message.startsWith('Invalid object identifier'),
		resolve(message, resource) {
			const group = this.getNodeParameter('group.value', '') as string;
			const parameterResource = resource === 'group' || message.includes(group) ? 'group' : 'user';

			return { message: `The ${parameterResource} ID is invalid`, description: ID_FORMAT };
		},
	},
];

/** Reads Graph's error envelope, which is absent on some gateway failures. */
export function graphError(body: unknown): {
	code: string;
	message: string;
	details?: Array<{ code: string; message: string }>;
} {
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

export async function handleErrorPostReceive(
	this: IExecuteSingleFunctions,
	data: INodeExecutionData[],
	response: IN8nHttpFullResponse,
): Promise<INodeExecutionData[]> {
	const statusCode = String(response.statusCode);
	if (!statusCode.startsWith('4') && !statusCode.startsWith('5')) {
		return data;
	}

	const resource = this.getNodeParameter('resource') as string;
	const operation = this.getNodeParameter('operation') as string;
	const { code, message, details } = graphError(response.body);
	const rule = findRule(resource, operation, code, message);

	if (rule) {
		const resolution =
			typeof rule.resolve === 'function'
				? rule.resolve.call(this, message, resource)
				: rule.resolve;

		// The empty-payload case is n8n's doing, not the user's, so the items pass through.
		if (resolution !== 'ignore') {
			throw new NodeApiError(this.getNode(), response as unknown as JsonObject, resolution);
		}

		return data;
	}

	if (details?.some((detail) => ['ObjectConflict', 'ConflictingObjects'].includes(detail.code))) {
		throw new NodeApiError(this.getNode(), response as unknown as JsonObject, {
			message: `The ${resource} already exists`,
			description: message,
		});
	}

	throw new NodeApiError(this.getNode(), response as unknown as JsonObject);
}
