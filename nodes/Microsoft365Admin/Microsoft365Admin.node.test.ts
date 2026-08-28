import type { INodeProperties, INodePropertyOptions } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import { Microsoft365Admin } from './Microsoft365Admin.node';

const node = new Microsoft365Admin();
const { description } = node;

function propertyNamed(name: string): INodeProperties {
	const property = description.properties.find((p) => p.name === name);
	if (!property) {
		throw new Error(`no property named ${name}`);
	}
	return property;
}

function operationsFor(resource: string): string[] {
	const property = description.properties.find(
		(p) => p.name === 'operation' && p.displayOptions?.show?.resource?.includes(resource),
	);
	if (!property) {
		throw new Error(`no operations for ${resource}`);
	}
	return (property.options as INodePropertyOptions[]).map((o) => String(o.value));
}

describe('node identity', () => {
	it('is named distinctly from the built-in Microsoft Entra ID node', () => {
		expect(description.name).toBe('microsoft365Admin');
		expect(description.displayName).toBe('Microsoft 365 Admin');
	});

	it('requires the service principal credential, not OAuth2', () => {
		expect(description.credentials).toEqual([
			{ name: 'microsoft365AdminServicePrincipalApi', required: true },
		]);
	});

	it('is usable as an AI tool', () => {
		expect(description.usableAsTool).toBe(true);
	});

	it('declares no requestDefaults, because the transport builds every URL', () => {
		expect(description.requestDefaults).toBeUndefined();
	});

	it('executes through the router rather than declarative routing', () => {
		expect(typeof node.execute).toBe('function');
		expect(node).not.toHaveProperty('customOperations');
	});
});

describe('resources and operations', () => {
	it('exposes authentication, group, license and user', () => {
		const values = (propertyNamed('resource').options as INodePropertyOptions[]).map((o) =>
			String(o.value),
		);
		expect(values).toEqual(['authentication', 'group', 'license', 'user']);
	});

	it.each([
		[
			'authentication',
			[
				'createTemporaryAccessPass',
				'deleteMethod',
				'getAllMethods',
				'getPasswordMethod',
				'resetPassword',
			],
		],
		[
			'group',
			[
				'addOwner',
				'create',
				'delete',
				'get',
				'getAll',
				'getMembers',
				'getOwners',
				'removeOwner',
				'update',
			],
		],
		['license', ['assign', 'assignGroup', 'queryHolders', 'queryTenant', 'queryUser', 'unassign']],
		[
			'user',
			[
				'addGroup',
				'create',
				'delete',
				'get',
				'getGroups',
				'getManager',
				'getAll',
				'removeGroup',
				'revokeSessions',
				'setManager',
				'update',
			],
		],
	])('offers the expected %s operations', (resource, expected) => {
		expect(operationsFor(resource)).toEqual(expected);
	});

	it('gives every operation an action label', () => {
		for (const property of description.properties.filter((p) => p.name === 'operation')) {
			for (const option of property.options as INodePropertyOptions[]) {
				expect((option as { action?: string }).action, String(option.value)).toBeTruthy();
			}
		}
	});
});

describe('parameter scoping', () => {
	it('scopes every parameter to a resource and an operation', () => {
		for (const property of description.properties) {
			if (property.name === 'resource') {
				continue;
			}
			const show = property.displayOptions?.show;
			expect(show?.resource, property.name).toBeDefined();
			if (property.name !== 'operation') {
				expect(show?.operation, property.name).toBeDefined();
			}
		}
	});

	it('never repeats a parameter name within one operation', () => {
		const seen = new Set<string>();
		for (const property of description.properties) {
			const show = property.displayOptions?.show;
			const key = JSON.stringify([show?.resource, show?.operation ?? null, property.name]);
			expect(seen.has(key), key).toBe(false);
			seen.add(key);
		}
	});

	it('keeps a sibling condition alongside the resource and operation scope', () => {
		expect(propertyNamed('mailEnabled').displayOptions?.show).toEqual({
			resource: ['group'],
			operation: ['create'],
			groupType: ['Unified'],
		});
	});
});

describe('expression capability', () => {
	it('blocks expressions only on resource and operation', () => {
		const blocked = description.properties.filter((p) => p.noDataExpression).map((p) => p.name);
		expect(new Set(blocked)).toEqual(new Set(['resource', 'operation']));
	});

	it('gives every resource locator a By ID mode that accepts an expression', () => {
		const locators = description.properties.filter((p) => p.type === 'resourceLocator');
		expect(locators.length).toBeGreaterThan(0);

		for (const locator of locators) {
			expect(
				(locator.modes ?? []).map((m) => m.name),
				locator.name,
			).toContain('id');
		}
	});

	it('leaves every dropdown expression-capable', () => {
		const loaders = description.properties.filter((p) => p.typeOptions?.loadOptionsMethod);
		expect(loaders.length).toBeGreaterThan(0);

		for (const loader of loaders) {
			expect(loader.noDataExpression, loader.name).toBeFalsy();
		}
	});

	it('flags ID-valued dropdowns with the Name or ID convention and an expression hint', () => {
		const idLoaders = description.properties.filter((p) => p.name === 'skuId');
		expect(idLoaders.length).toBeGreaterThan(0);

		for (const loader of idLoaders) {
			expect(loader.displayName, loader.name).toMatch(/Names? or IDs?$/);
			expect(loader.description, loader.name).toContain('expression');
		}
	});
});

describe('loadOptions and listSearch wiring', () => {
	it('registers every searchListMethod referenced by a resource locator', () => {
		const referenced = new Set(
			description.properties
				.filter((p) => p.type === 'resourceLocator')
				.flatMap((p) => p.modes ?? [])
				.map((m) => m.typeOptions?.searchListMethod)
				.filter(Boolean) as string[],
		);

		expect(referenced.size).toBeGreaterThan(0);
		for (const method of referenced) {
			expect(node.methods.listSearch, method).toHaveProperty(method);
		}
	});

	it('registers every loadOptionsMethod referenced by a dropdown', () => {
		const referenced = new Set(
			description.properties
				.map((p) => p.typeOptions?.loadOptionsMethod)
				.filter(Boolean) as string[],
		);

		expect(referenced.size).toBeGreaterThan(0);
		for (const method of referenced) {
			expect(node.methods.loadOptions, method).toHaveProperty(method);
		}
	});
});
