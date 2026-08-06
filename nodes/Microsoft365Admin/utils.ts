import type { IDataObject } from 'n8n-workflow';

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
	const schema = /<Schema[^>]*Namespace="microsoft\.graph"[^>]*>([\s\S]*?)<\/Schema>/.exec(metadata);
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

/** Shape of a Graph collection response used by the resource-locator searches. */
export interface DirectoryListResponse {
	value?: Array<{ id: string; displayName: string }>;
	'@odata.nextLink'?: string;
}
