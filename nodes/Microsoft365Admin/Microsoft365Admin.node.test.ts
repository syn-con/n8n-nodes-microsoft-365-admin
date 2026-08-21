import type { IExecuteFunctions, INodeProperties, INodePropertyOptions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

vi.mock('./LicenseFunctions', () => ({ executeLicenseWrite: vi.fn(async () => [[]]) }));
// Partial, so the descriptions keep the real hooks they reference.
vi.mock('./AuthenticationFunctions', async (importOriginal) => ({
	...(await importOriginal<typeof import('./AuthenticationFunctions')>()),
	executeResetPassword: vi.fn(async () => [[]]),
}));

import { executeResetPassword } from './AuthenticationFunctions';
import { executeLicenseWrite } from './LicenseFunctions';
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
	it('exposes authentication, group, license and user', () => {
		const values = (propertyNamed('resource').options as INodePropertyOptions[]).map(
			(o) => o.value,
		);
		expect(values).toEqual(['authentication', 'group', 'license', 'user']);
	});
});

describe('operations', () => {
	it.each([
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
			'authentication',
			[
				'createTemporaryAccessPass',
				'deleteMethod',
				'getAllMethods',
				'getPasswordMethod',
				'resetPassword',
			],
		],
	])('%s exposes the expected operations', (resource, expected) => {
		expect(operationsFor(resource).map((o) => o.value)).toEqual(expected);
	});

	it('gives every operation an action label and a way to execute', () => {
		const custom = node.customOperations as Record<string, Record<string, unknown>>;

		for (const resource of ['user', 'group', 'license', 'authentication']) {
			for (const operation of operationsFor(resource)) {
				const label = `${resource}.${operation.value}`;
				// Either declarative routing or a custom handler — never both, since a custom
				// handler silently wins and would leave the routing as a decoy.
				const routed = operation.routing?.request !== undefined;
				const handled = custom[resource]?.[operation.value as string] !== undefined;

				expect(routed !== handled, label).toBe(true);
				expect(operation.action, label).toBeTruthy();
			}
		}
	});

	it('lists operations alphabetically by display name', () => {
		for (const resource of ['user', 'group', 'license', 'authentication']) {
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

	it('sends the eventual-consistency header and $count required for the holders filter', () => {
		const request = routingFor('queryHolders');
		expect(request?.headers).toMatchObject({ ConsistencyLevel: 'eventual' });
		expect(request?.qs).toMatchObject({ $count: 'true' });
	});

	it('switches the holders URL between users and groups', () => {
		expect(routingFor('queryHolders')?.url).toContain('holderType');
	});
});

describe('license writes', () => {
	// The bodies these operations send are built in LicenseFunctions.ts, so what matters
	// here is only that the node hands the operations over to it.
	it.each(['assign', 'assignGroup', 'unassign'] as const)(
		'routes %s through a custom operation',
		async (op) => {
			expect(operationsFor('license').find((o) => o.value === op)?.routing).toBeUndefined();

			const context = { getInputData: () => [] } as unknown as IExecuteFunctions;
			await node.customOperations.license[op].call(context);

			expect(executeLicenseWrite).toHaveBeenCalledWith(op);
			expect(vi.mocked(executeLicenseWrite).mock.contexts.at(-1)).toBe(context);
		},
	);

	it('lets every write pick more than one SKU at a time', () => {
		for (const operation of ['assign', 'assignGroup', 'unassign']) {
			const sku = fieldsFor('license', operation).find((p) => p.name === 'skuId');
			expect(sku?.type, operation).toBe('multiOptions');
		}
	});

	it('offers the same batching and retry options to all three writes', () => {
		for (const operation of ['assign', 'assignGroup', 'unassign']) {
			const options = fieldsFor('license', operation).find((p) => p.name === 'options');
			const names = (options?.options ?? []).map((o) => (o as INodeProperties).name);

			expect(names, operation).toEqual(
				expect.arrayContaining(['combineItems', 'maxRetries', 'waitBetweenRequests']),
			);
		}
	});

	it('combines items for the same target unless told otherwise', () => {
		const options = fieldsFor('license', 'assign').find((p) => p.name === 'options');
		const combine = (options?.options ?? []).find(
			(o) => (o as INodeProperties).name === 'combineItems',
		) as INodeProperties;

		expect(combine.default).toBe(true);
	});
});

describe('license query bodies', () => {
	function sendValue(operation: string, field: string): string {
		const property = fieldsFor('license', operation).find((p) => p.name === field);
		return property?.routing?.send?.value as string;
	}

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

describe('authentication methods', () => {
	function routingFor(operation: string) {
		return operationsFor('authentication').find((o) => o.value === operation)?.routing?.request;
	}

	it('lists a user’s methods, and reads the password method from its own collection', () => {
		expect(routingFor('getAllMethods')?.url).toContain('/authentication/methods');
		expect(routingFor('getPasswordMethod')?.url).toContain('/authentication/passwordMethods');
	});

	it('addresses a method by its collection as well as its ID', () => {
		// An ID alone is not enough: each method type lives under its own segment.
		const url = routingFor('deleteMethod')?.url ?? '';
		expect(routingFor('deleteMethod')?.method).toBe('DELETE');
		expect(url).toContain('methodType');
		expect(url.indexOf('methodType')).toBeLessThan(url.indexOf('$parameter["method"]'));
	});

	it('offers only method types that Graph can delete in v1.0', () => {
		const types = fieldsFor('authentication', 'deleteMethod').find((p) => p.name === 'methodType');
		const values = (types?.options ?? []).map((o) => (o as INodePropertyOptions).value);

		expect(values).toContain('microsoftAuthenticatorMethods');
		expect(values).toContain('phoneMethods');
		expect(values).toContain('fido2Methods');
		// Graph has no delete for a password, and beta-only types are out of reach here.
		expect(values).not.toContain('passwordMethods');
	});

	it('posts a temporary access pass with the policy defaults left out', () => {
		const routing = routingFor('createTemporaryAccessPass');
		expect(routing?.url).toContain(
			'/authentication/temporaryAccessPassMethods',
		);
		expect(routing?.body).toBe('={{ $parameter["options"] || {} }}');

		const options = fieldsFor('authentication', 'createTemporaryAccessPass').find(
			(p) => p.name === 'options',
		);
		const sent = (options?.options ?? []).map(
			(o) => (o as INodeProperties).routing?.send?.property,
		);

		expect(sent).toEqual(['isUsableOnce', 'lifetimeInMinutes', 'startDateTime']);
	});

	it('keeps the pass lifetime inside the range Graph accepts', () => {
		const options = fieldsFor('authentication', 'createTemporaryAccessPass').find(
			(p) => p.name === 'options',
		);
		const lifetime = (options?.options ?? []).find(
			(o) => (o as INodeProperties).name === 'lifetimeInMinutes',
		) as INodeProperties;

		expect(lifetime.typeOptions).toMatchObject({ minValue: 10, maxValue: 43200 });
	});

	it('routes the password reset through a custom operation', async () => {
		// The authentication methods API cannot reset a password app-only, so this one
		// writes passwordProfile instead — and has to return the password it generated.
		expect(routingFor('resetPassword')).toBeUndefined();

		const context = { getInputData: () => [] } as unknown as IExecuteFunctions;
		await node.customOperations.authentication.resetPassword.call(context);

		expect(executeResetPassword).toHaveBeenCalled();
		expect(vi.mocked(executeResetPassword).mock.contexts.at(-1)).toBe(context);
	});

	it('guards the request path of every declarative operation', () => {
		// A user or method ID arriving from workflow data goes into the URL, so each of these
		// operations has to carry the check. Reset Password does its own, being programmatic.
		for (const operation of operationsFor('authentication')) {
			const value = operation.value as string;
			if (value === 'resetPassword') continue;

			const guarded = fieldsFor('authentication', value).some(
				(p) => (p.routing?.send?.preSend ?? []).length > 0,
			);
			expect(guarded, value).toBe(true);
		}
	});

	it('asks for the user on every operation', () => {
		for (const operation of operationsFor('authentication')) {
			const fields = fieldsFor('authentication', operation.value as string);
			expect(
				fields.some((p) => p.name === 'user' && p.type === 'resourceLocator'),
				operation.value as string,
			).toBe(true);
		}
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
			// Plural for the multi-selects on the write operations.
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
		const declarations = properties
			.map((p) => `<Property Name="${p}" Type="Edm.String"/>`)
			.join('');
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
		const context = metadataContext('group', [
			'displayName',
			'allowExternalSenders',
			'unseenCount',
		]);

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
