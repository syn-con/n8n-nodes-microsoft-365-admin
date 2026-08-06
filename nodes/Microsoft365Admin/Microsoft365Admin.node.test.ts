import type { INodeProperties, INodePropertyOptions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { Microsoft365Admin } from './Microsoft365Admin.node';

const node = new Microsoft365Admin();
const { description } = node;

function propertyNamed(name: string): INodeProperties {
	const property = description.properties.find((p) => p.name === name);
	if (!property) throw new Error(`no property named ${name}`);
	return property;
}

function operationsFor(resource: string): INodePropertyOptions[] {
	const property = description.properties.find(
		(p) =>
			p.name === 'operation' &&
			(p.displayOptions?.show?.resource as string[] | undefined)?.includes(resource),
	);
	if (!property) throw new Error(`no operation dropdown for ${resource}`);
	return property.options as INodePropertyOptions[];
}

function fieldsFor(resource: string, operation: string): INodeProperties[] {
	return description.properties.filter((p) => {
		const show = p.displayOptions?.show;
		return (
			(show?.resource as string[] | undefined)?.includes(resource) &&
			(show?.operation as string[] | undefined)?.includes(operation)
		);
	});
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

	it('resolves the Graph base URL from the credential, tolerating a trailing slash', () => {
		expect(description.requestDefaults?.baseURL).toContain('graphApiBaseUrl');
		expect(description.requestDefaults?.baseURL).toContain('/v1.0');
	});
});

describe('resources', () => {
	it('exposes group, license and user', () => {
		const values = (propertyNamed('resource').options as INodePropertyOptions[]).map(
			(o) => o.value,
		);
		expect(values).toEqual(['group', 'license', 'user']);
	});
});

describe('operations', () => {
	it.each([
		['user', ['addGroup', 'create', 'delete', 'get', 'getGroups', 'getManager', 'getAll', 'removeGroup', 'revokeSessions', 'setManager', 'update']],
		['group', ['addOwner', 'create', 'delete', 'get', 'getAll', 'getMembers', 'getOwners', 'removeOwner', 'update']],
		['license', ['assign', 'assignGroup', 'queryHolders', 'queryTenant', 'queryUser', 'unassign']],
	])('%s exposes the expected operations', (resource, expected) => {
		expect(operationsFor(resource).map((o) => o.value)).toEqual(expected);
	});

	it('gives every operation a routing request and an action label', () => {
		for (const resource of ['user', 'group', 'license']) {
			for (const operation of operationsFor(resource)) {
				expect(operation.routing?.request, `${resource}.${operation.value}`).toBeDefined();
				expect(operation.action, `${resource}.${operation.value}`).toBeTruthy();
			}
		}
	});

	it('lists operations alphabetically by display name', () => {
		for (const resource of ['user', 'group', 'license']) {
			const names = operationsFor(resource).map((o) => o.name);
			expect(names, resource).toEqual([...names].sort((a, b) => a.localeCompare(b)));
		}
	});
});

describe('license routing', () => {
	function routingFor(operation: string) {
		const found = operationsFor('license').find((o) => o.value === operation);
		return found?.routing?.request;
	}

	it('reads tenant SKUs from /subscribedSkus', () => {
		expect(routingFor('queryTenant')).toMatchObject({ method: 'GET', url: '/subscribedSkus' });
	});

	it('reads a user’s licenses from /licenseDetails', () => {
		expect(routingFor('queryUser')?.url).toContain('/licenseDetails');
	});

	it('assigns via POST to assignLicense on the user', () => {
		expect(routingFor('assign')).toMatchObject({ method: 'POST' });
		expect(routingFor('assign')?.url).toContain('/users/');
		expect(routingFor('assign')?.url).toContain('/assignLicense');
	});

	it('assigns to a group via the group assignLicense endpoint', () => {
		expect(routingFor('assignGroup')?.url).toContain('/groups/');
		expect(routingFor('assignGroup')?.url).toContain('/assignLicense');
	});

	it('unassigns through the same assignLicense endpoint', () => {
		expect(routingFor('unassign')?.url).toContain('/assignLicense');
	});

	it('sends the eventual-consistency header and $count required for the holders filter', () => {
		const request = routingFor('queryHolders');
		expect(request?.headers).toMatchObject({ ConsistencyLevel: 'eventual' });
		expect(request?.qs).toMatchObject({ $count: 'true' });
	});

	it('switches the holders URL between users and groups', () => {
		expect(routingFor('queryHolders')?.url).toContain('holderType');
	});
});

describe('license request bodies', () => {
	function sendValue(operation: string, field: string): string {
		const property = fieldsFor('license', operation).find((p) => p.name === field);
		return property?.routing?.send?.value as string;
	}

	it('assign builds an addLicenses array from the chosen SKU', () => {
		expect(sendValue('assign', 'skuId')).toContain('skuId: $value');
		expect(sendValue('assign', 'skuId')).toContain('disabledPlans');
	});

	it('assign always sends an empty removeLicenses, which Graph requires', () => {
		expect(sendValue('assign', 'removeLicenses')).toBe('={{ [] }}');
	});

	it('unassign sends the SKU under removeLicenses and an empty addLicenses', () => {
		expect(sendValue('unassign', 'skuId')).toBe('={{ [$value] }}');
		expect(sendValue('unassign', 'addLicenses')).toBe('={{ [] }}');
	});

	it('holders filters on assignedLicenses with an unquoted GUID', () => {
		const filter = sendValue('queryHolders', 'skuId');
		expect(filter).toContain('assignedLicenses/any');
		expect(filter).not.toContain("eq '");
	});

	it('holders selects the fields that reveal group-inherited assignments', () => {
		const select = sendValue('queryHolders', 'select');
		expect(select).toContain('licenseAssignmentStates');
		expect(select).toContain('licenseProcessingState');
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
			const modes = (locator.modes ?? []).map((m) => m.name);
			expect(modes, locator.name).toContain('id');
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
		// Applies to fields whose value is an opaque ID the user may want to supply from
		// earlier data. The upstream property selectors are excluded: their value is the
		// property name itself, so there is no ID to substitute.
		const idLoaders = description.properties.filter((p) => p.name === 'skuId');
		expect(idLoaders.length).toBeGreaterThan(0);

		for (const loader of idLoaders) {
			expect(loader.displayName, loader.name).toMatch(/Name or ID$/);
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

		for (const method of referenced) {
			expect(node.methods.loadOptions, method).toHaveProperty(method);
		}
	});

	// Drives the real metadata chain rather than stubbing the loader, so the filtering
	// is exercised against the same parsing the node uses at runtime.
	function metadataContext(entity: string, properties: string[]) {
		const declarations = properties.map((p) => `<Property Name="${p}" Type="Edm.String"/>`).join('');
		const metadata =
			'<Schema Namespace="microsoft.graph">' +
			`<EntityType Name="${entity}">${declarations}</EntityType>` +
			'</Schema>';

		return {
			getCredentials: vi.fn(async () => ({})),
			helpers: { httpRequestWithAuthentication: vi.fn(async () => metadata) },
		} as never;
	}

	it('filters list-unsupported properties out of the group getAll selector', async () => {
		const context = metadataContext('group', ['displayName', 'allowExternalSenders', 'unseenCount']);

		const all = await node.methods.loadOptions.getGroupProperties.call(context);
		const forGetAll = await node.methods.loadOptions.getGroupPropertiesGetAll.call(context);

		expect(all.map((o) => o.value)).toContain('allowExternalSenders');
		expect(forGetAll.map((o) => o.value)).toEqual(['displayName']);
	});

	it('filters list-unsupported properties out of the user getAll selector', async () => {
		const context = metadataContext('user', ['displayName', 'aboutMe', 'skills']);

		const all = await node.methods.loadOptions.getUserProperties.call(context);
		const forGetAll = await node.methods.loadOptions.getUserPropertiesGetAll.call(context);

		expect(all.map((o) => o.value)).toContain('aboutMe');
		expect(forGetAll.map((o) => o.value)).toEqual(['displayName']);
	});
});
