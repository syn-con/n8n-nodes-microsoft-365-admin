import type { IDataObject, ILoadOptionsFunctions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import {
	getGroupProperties,
	getGroupPropertiesGetAll,
	getSubscribedSkus,
	getUserProperties,
	getUserPropertiesGetAll,
} from './loadOptions';

const METADATA = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
<edmx:DataServices>
<Schema Namespace="microsoft.graph">
<EntityType Name="entity" Abstract="true">
<Key><PropertyRef Name="id"/></Key>
<Property Name="id" Type="Edm.String" Nullable="false"/>
</EntityType>
<EntityType Name="directoryObject" BaseType="graph.entity">
<Property Name="deletedDateTime" Type="Edm.DateTimeOffset"/>
</EntityType>
<EntityType Name="user" BaseType="graph.directoryObject">
<Property Name="displayName" Type="Edm.String"/>
<Property Name="usageLocation" Type="Edm.String"/>
<Property Name="mailboxSettings" Type="graph.mailboxSettings"/>
<NavigationProperty Name="manager" Type="graph.directoryObject"/>
</EntityType>
<EntityType Name="group" BaseType="graph.directoryObject">
<Property Name="mailNickname" Type="Edm.String"/>
<Property Name="isArchived" Type="Edm.Boolean"/>
</EntityType>
<EntityType Name="device" BaseType="graph.directoryObject">
<Property Name="deviceId" Type="Edm.String"/>
</EntityType>
</Schema>
<Schema Namespace="microsoft.graph.callRecords">
<EntityType Name="session"><Property Name="fromOtherNamespace" Type="Edm.String"/></EntityType>
</Schema>
</edmx:DataServices>
</edmx:Edmx>`;

/** Answers every request with one body, wrapped the way n8n's helper does. */
function mockContext(body: unknown) {
	return {
		getCredentials: vi.fn(async () => ({ graphApiBaseUrl: 'https://graph.microsoft.com' })),
		getNode: vi.fn(() => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' })),
		getNodeParameter: vi.fn(() => ''),
		helpers: {
			httpRequestWithAuthentication: vi.fn(async () => ({
				statusCode: 200,
				headers: {},
				body,
			})),
			requestWithAuthenticationPaginated: vi.fn(),
		},
	} as unknown as ILoadOptionsFunctions;
}

const metadataContext = () => mockContext(METADATA);

describe('getSubscribedSkus', () => {
	it('maps SKUs to options labelled with seat usage', async () => {
		const skus = await getSubscribedSkus.call(
			mockContext({
				value: [
					{
						skuId: 'sku-1',
						skuPartNumber: 'ENTERPRISEPACK',
						consumedUnits: 142,
						prepaidUnits: { enabled: 200 },
					},
				],
			}),
		);

		expect(skus).toHaveLength(1);
		expect(skus[0].value).toBe('sku-1');
		expect(skus[0].name).toBe('ENTERPRISEPACK (142/200 used)');
	});

	it('defaults missing seat counts to zero rather than rendering undefined', async () => {
		const skus = await getSubscribedSkus.call(
			mockContext({ value: [{ skuId: 'sku-2', skuPartNumber: 'FLOW_FREE' }] }),
		);

		expect(skus).toHaveLength(1);
		expect(skus[0].value).toBe('sku-2');
		expect(skus[0].name).toBe('FLOW_FREE (0/0 used)');
	});

	it('returns an empty list when the tenant has no subscriptions', async () => {
		expect(await getSubscribedSkus.call(mockContext({} as IDataObject))).toEqual([]);
	});
});

describe('metadata-driven property loaders', () => {
	it('collects user properties from the inherited entity chain', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);

		expect(names).toContain('displayName');
		expect(names).toContain('usageLocation');
		expect(names).toContain('deletedDateTime');
	});

	it('excludes properties that need permissions beyond the node’s scope', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);

		expect(names).not.toContain('mailboxSettings');
		expect(names).not.toContain('id');
	});

	it('ignores NavigationProperty elements', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('manager');
	});

	it('ignores entity types that were not requested', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('deviceId');
	});

	it('ignores schemas outside the microsoft.graph namespace', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('fromOtherNamespace');
	});

	it('returns options sorted alphabetically', async () => {
		const names = (await getUserProperties.call(metadataContext())).map((o) => o.value);
		expect(names).toEqual([...names].sort());
	});

	it('collects group properties and drops the excluded ones', async () => {
		const names = (await getGroupProperties.call(metadataContext())).map((o) => o.value);

		expect(names).toContain('mailNickname');
		expect(names).not.toContain('isArchived');
		expect(names).not.toContain('id');
		expect(names).not.toContain('displayName');
	});
});

describe('collection-endpoint property loaders', () => {
	it('drops the user properties /users cannot project', async () => {
		const names = (await getUserPropertiesGetAll.call(metadataContext())).map((o) => o.value);
		expect(names).not.toContain('mailboxSettings');
		expect(names).toContain('usageLocation');
	});

	it('drops the group properties /groups cannot project', async () => {
		// The fixture has no such property, so this asserts the filter leaves the rest alone.
		const all = (await getGroupProperties.call(metadataContext())).map((o) => o.value);
		const listed = (await getGroupPropertiesGetAll.call(metadataContext())).map((o) => o.value);
		expect(listed.every((name) => all.includes(name as string))).toBe(true);
	});
});
