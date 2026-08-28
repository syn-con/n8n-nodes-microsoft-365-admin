import { NodeOperationError, type IDataObject, type INode } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import {
	assertPathSafe,
	deepMerge,
	errorItem,
	extractEntityProperties,
	isPathSafe,
	updateDisplayOptions,
} from './utils';

const NODE = { name: 'Microsoft 365 Admin', type: 'microsoft365Admin' } as unknown as INode;
const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';

describe('deepMerge', () => {
	it('copies new keys onto the target', () => {
		expect(deepMerge({ a: 1 }, { b: 2 })).toEqual({ a: 1, b: 2 });
	});

	it('overwrites scalars', () => {
		expect(deepMerge({ a: 1 }, { a: 2 })).toEqual({ a: 2 });
	});

	it('merges nested objects rather than replacing them', () => {
		const result = deepMerge(
			{ nested: { keep: 1, replace: 1 } },
			{ nested: { replace: 2, add: 3 } },
		);
		expect(result).toEqual({ nested: { keep: 1, replace: 2, add: 3 } });
	});

	it('replaces arrays wholesale instead of merging by index', () => {
		expect(deepMerge({ list: [1, 2, 3] }, { list: [9] })).toEqual({ list: [9] });
	});

	it('replaces an object with a scalar when types differ', () => {
		expect(deepMerge({ a: { b: 1 } }, { a: 'scalar' })).toEqual({ a: 'scalar' });
	});

	it('treats null as a scalar rather than recursing into it', () => {
		expect(deepMerge({ a: { b: 1 } }, { a: null })).toEqual({ a: null });
	});

	it('mutates and returns the target, matching the previous lodash behaviour', () => {
		const target: IDataObject = { a: 1 };
		expect(deepMerge(target, { b: 2 })).toBe(target);
		expect(target).toEqual({ a: 1, b: 2 });
	});
});

describe('path safety', () => {
	it('lets an object ID and a userPrincipalName through', () => {
		expect(assertPathSafe(NODE, USER, 'user')).toBe(USER);
		expect(assertPathSafe(NODE, "o'brien.a@contoso.com", 'user')).toBe("o'brien.a@contoso.com");
	});

	it('lets a base64url method ID through', () => {
		expect(assertPathSafe(NODE, '-2_GRUg2-HYz6_1YG4YRAQ2', 'method')).toBe(
			'-2_GRUg2-HYz6_1YG4YRAQ2',
		);
	});

	// Left unchecked, these would move the request to a different Graph endpoint and leave
	// the intended path in the query string.
	it.each([
		['user', 'alice/revokeSignInSessions?x='],
		['method', '../../../users/bob'],
		['method type', 'methods%2F..'],
	])('refuses an unsafe %s', (label, value) => {
		expect(() => assertPathSafe(NODE, value, label)).toThrow(/characters that are not allowed/);
	});

	it('refuses an empty value, naming what was missing', () => {
		expect(() => assertPathSafe(NODE, '', 'user')).toThrow(/No user was given/);
		expect(() => assertPathSafe(NODE, undefined, 'group')).toThrow(/No group was given/);
	});

	it('raises a node error carrying the item index', () => {
		try {
			assertPathSafe(NODE, 'a/b', 'user', 7);
			expect.unreachable('should have thrown');
		} catch (error) {
			expect(error).toBeInstanceOf(NodeOperationError);
			expect((error as NodeOperationError).context.itemIndex).toBe(7);
		}
	});

	it('isPathSafe answers the same question without throwing', () => {
		expect(isPathSafe(USER)).toBe(true);
		expect(isPathSafe('alice/x?y=')).toBe(false);
		expect(isPathSafe('')).toBe(false);
		expect(isPathSafe(undefined)).toBe(false);
	});
});

describe('updateDisplayOptions', () => {
	it('stamps the resource and operation onto every property', () => {
		const [property] = updateDisplayOptions({ show: { resource: ['user'], operation: ['get'] } }, [
			{ displayName: 'User', name: 'user', type: 'string', default: '' },
		]);

		expect(property.displayOptions?.show).toEqual({ resource: ['user'], operation: ['get'] });
	});

	it('keeps a sibling-parameter condition the property already had', () => {
		const [property] = updateDisplayOptions(
			{ show: { resource: ['group'], operation: ['create'] } },
			[
				{
					displayName: 'Mail Enabled',
					name: 'mailEnabled',
					type: 'boolean',
					default: false,
					displayOptions: { show: { groupType: ['Unified'] } },
				},
			],
		);

		expect(property.displayOptions?.show).toEqual({
			resource: ['group'],
			operation: ['create'],
			groupType: ['Unified'],
		});
	});

	it('leaves the source properties untouched', () => {
		const properties = [
			{ displayName: 'User', name: 'user', type: 'string' as const, default: '' },
		];
		updateDisplayOptions({ show: { resource: ['user'] } }, properties);
		expect(properties[0].displayOptions).toBeUndefined();
	});
});

describe('extractEntityProperties', () => {
	const metadata = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx><edmx:DataServices>
<Schema Namespace="microsoft.graph">
<EntityType Name="user">
<Property Name="displayName" Type="Edm.String"/>
<NavigationProperty Name="manager" Type="graph.directoryObject"/>
</EntityType>
<EntityType Name="device"><Property Name="deviceId" Type="Edm.String"/></EntityType>
</Schema>
<Schema Namespace="microsoft.graph.callRecords">
<EntityType Name="user"><Property Name="fromOtherNamespace" Type="Edm.String"/></EntityType>
</Schema>
</edmx:DataServices></edmx:Edmx>`;

	it('reads the properties of the requested entity only', () => {
		expect(extractEntityProperties(metadata, ['user'])).toEqual(['displayName']);
	});

	it('ignores NavigationProperty elements', () => {
		expect(extractEntityProperties(metadata, ['user'])).not.toContain('manager');
	});

	it('ignores schemas outside the microsoft.graph namespace', () => {
		expect(extractEntityProperties(metadata, ['user'])).not.toContain('fromOtherNamespace');
	});
});

describe('errorItem', () => {
	it('carries the message and pairs the item back to its input', () => {
		expect(errorItem(new Error('boom'), 3)).toEqual({
			json: { error: 'boom' },
			pairedItem: { item: 3 },
		});
	});
});
