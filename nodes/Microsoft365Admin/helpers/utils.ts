import {
	NodeApiError,
	NodeOperationError,
	type IDataObject,
	type IExecuteFunctions,
	type INode,
	type INodeExecutionData,
	type INodeProperties,
} from 'n8n-workflow';

/**
 * A user ID, a method ID and a method type all go into the request path, and all three
 * can come from workflow data. None of them legitimately contains any of these
 * characters — a Graph method ID is base64url — so a value carrying one is refused
 * rather than escaped: `alice/authentication/methods?x=` would otherwise turn a delete
 * of one method into a request against a different Graph endpoint.
 */
const UNSAFE_IN_PATH = /[\\/?#%\s]/;

export function isFilled(value: unknown): value is string {
	return typeof value === 'string' && value.trim() !== '';
}

/**
 * Returns `value` if it is safe to interpolate into a request path, and throws otherwise.
 *
 * Declarative routing ran this once as a `preSend` hook over every path parameter; each
 * operation now guards the segments it actually builds its URL from.
 */
export function assertPathSafe(
	node: INode,
	value: unknown,
	label: string,
	itemIndex?: number,
): string {
	if (!isFilled(value)) {
		throw new NodeOperationError(node, `No ${label} was given`, { itemIndex });
	}

	if (UNSAFE_IN_PATH.test(value)) {
		throw new NodeOperationError(
			node,
			`The ${label} value contains characters that are not allowed`,
			{
				itemIndex,
				description:
					'Pick the value from the list, or supply an ID, e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
			},
		);
	}

	return value;
}

/** True when a value could be interpolated into a path without reshaping the URL. */
export function isPathSafe(value: unknown): value is string {
	return isFilled(value) && !UNSAFE_IN_PATH.test(value);
}

/** What an item carries out when the run continues on failure. */
export function errorItem(error: unknown, itemIndex: number): INodeExecutionData {
	return { json: { error: (error as Error).message }, pairedItem: { item: itemIndex } };
}

/** Anything that is not already a node error would otherwise reach n8n unwrapped. */
export function asNodeError(this: IExecuteFunctions, error: unknown, itemIndex: number): Error {
	if (error instanceof NodeApiError || error instanceof NodeOperationError) {
		return error;
	}

	return new NodeOperationError(this.getNode(), error as Error, { itemIndex });
}

function isPlainObject(value: unknown): value is IDataObject {
	return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Recursively merges `source` into `target`, replacing lodash's `merge`.
 *
 * Community node packages may not ship runtime dependencies, and the few call
 * sites here only ever merge plain JSON objects.
 */
export function deepMerge(target: IDataObject, source: IDataObject): IDataObject {
	for (const [key, value] of Object.entries(source)) {
		const existing = target[key];
		if (isPlainObject(existing) && isPlainObject(value)) {
			deepMerge(existing, value);
		} else {
			target[key] = value;
		}
	}
	return target;
}

/**
 * Stamps `resource`/`operation` display options onto every property of an operation.
 *
 * Operation files declare their parameters without repeating which resource and
 * operation they belong to; this folds that in once at export. nodes-base has the
 * same helper behind its `@utils/utilities` alias, which a community package cannot
 * import, so it is reimplemented here.
 */
export function updateDisplayOptions(
	displayOptions: INodeProperties['displayOptions'],
	properties: INodeProperties[],
): INodeProperties[] {
	return properties.map((property) => ({
		...property,
		displayOptions: {
			...property.displayOptions,
			show: {
				...displayOptions?.show,
				...property.displayOptions?.show,
			},
		},
	}));
}

/**
 * Reads property names for the given entity types out of Graph's `$metadata` document.
 *
 * The document is machine-generated CSDL with a stable shape and we need only the
 * `Name` attribute of `<Property>` elements inside particular `<EntityType>` blocks,
 * so a targeted scan avoids taking on an XML-parser dependency — which community
 * node packages are not permitted to bundle.
 *
 * `<NavigationProperty>` elements are not matched: the pattern anchors on `<Property`
 * immediately after the angle bracket.
 */
export function extractEntityProperties(metadata: string, entityNames: string[]): string[] {
	const schema = /<Schema[^>]*Namespace="microsoft\.graph"[^>]*>([\s\S]*?)<\/Schema>/.exec(
		metadata,
	);
	const scope = schema?.[1] ?? metadata;

	const properties: string[] = [];
	const entityPattern = /<EntityType[^>]*\sName="([^"]+)"[^>]*>([\s\S]*?)<\/EntityType>/g;
	const propertyPattern = /<Property\s[^>]*\bName="([^"]+)"/g;

	for (const entity of scope.matchAll(entityPattern)) {
		const [, entityName, body] = entity;
		if (!entityNames.includes(entityName)) {
			continue;
		}

		for (const property of body.matchAll(propertyPattern)) {
			properties.push(property[1]);
		}
	}

	return properties;
}
