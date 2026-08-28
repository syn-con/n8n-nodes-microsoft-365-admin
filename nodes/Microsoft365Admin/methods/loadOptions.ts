import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import { extractEntityProperties } from '../helpers/utils';
import { microsoftApiRequest } from '../transport';

/** Properties Graph advertises but will not return on a directory read. */
const UNREADABLE_GROUP_PROPERTIES = ['id', 'isArchived', 'hasMembersWithLicenseErrors'];

/**
 * `signInActivity` needs AuditLog.Read.All and `mailboxSettings` needs MailboxSettings.Read,
 * neither of which this node asks for, so offering them would only produce 403s.
 */
const UNREADABLE_USER_PROPERTIES = [
	'id',
	'deviceEnrollmentLimit',
	'mailboxSettings',
	'print',
	'signInActivity',
];

/** Properties the `/groups` collection endpoint cannot project, unlike a single-group read. */
const GROUP_PROPERTIES_NOT_ON_LIST = [
	'allowExternalSenders',
	'autoSubscribeNewMembers',
	'hideFromAddressLists',
	'hideFromOutlookClients',
	'isSubscribedByMail',
	'unseenCount',
];

/** Properties the `/users` collection endpoint cannot project, unlike a single-user read. */
const USER_PROPERTIES_NOT_ON_LIST = [
	'aboutMe',
	'birthday',
	'hireDate',
	'interests',
	'mySite',
	'pastProjects',
	'preferredName',
	'responsibilities',
	'schools',
	'skills',
	'mailboxSettings',
];

function toOptions(properties: string[]): INodePropertyOptions[] {
	return properties.sort().map((property) => ({ name: property, value: property }));
}

export async function getGroupProperties(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const response = await microsoftApiRequest.call(this, 'GET', '/$metadata#groups');

	const properties = extractEntityProperties(response as string, [
		'entity',
		'directoryObject',
		'group',
	]).filter((property) => !UNREADABLE_GROUP_PROPERTIES.includes(property));

	return toOptions(properties);
}

export async function getGroupPropertiesGetAll(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	return (await getGroupProperties.call(this)).filter(
		(option) => !GROUP_PROPERTIES_NOT_ON_LIST.includes(option.value as string),
	);
}

export async function getUserProperties(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const response = await microsoftApiRequest.call(this, 'GET', '/$metadata#users');

	const properties = extractEntityProperties(response as string, [
		'entity',
		'directoryObject',
		'user',
	]).filter((property) => !UNREADABLE_USER_PROPERTIES.includes(property));

	return toOptions(properties);
}

export async function getUserPropertiesGetAll(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	return (await getUserProperties.call(this)).filter(
		(option) => !USER_PROPERTIES_NOT_ON_LIST.includes(option.value as string),
	);
}

/**
 * Lists the tenant's subscribed license SKUs for the License operations.
 *
 * `/subscribedSkus` returns the full collection in one response — it supports no
 * paging or `$filter`, so there is nothing to paginate over here.
 */
export async function getSubscribedSkus(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const response = (await microsoftApiRequest.call(this, 'GET', '/subscribedSkus')) as {
		value?: Array<{
			skuId: string;
			skuPartNumber: string;
			consumedUnits?: number;
			prepaidUnits?: { enabled?: number };
		}>;
	};

	return (response.value ?? []).map((sku) => {
		const consumed = sku.consumedUnits ?? 0;
		const enabled = sku.prepaidUnits?.enabled ?? 0;
		return {
			// Surfacing seat usage makes it obvious when a SKU has nothing left to assign.
			name: `${sku.skuPartNumber} (${consumed}/${enabled} used)`,
			value: sku.skuId,
		};
	});
}
