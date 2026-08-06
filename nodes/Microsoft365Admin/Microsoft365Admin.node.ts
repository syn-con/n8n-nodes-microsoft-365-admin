import {
	NodeConnectionTypes,
	type ILoadOptionsFunctions,
	type INodePropertyOptions,
	type INodeType,
	type INodeTypeDescription,
} from 'n8n-workflow';

import {
	groupFields,
	groupOperations,
	licenseFields,
	licenseOperations,
	userFields,
	userOperations,
} from './descriptions';
import {
	getGroupProperties,
	getGroups,
	getSubscribedSkus,
	getUserProperties,
	getUsers,
} from './GenericFunctions';

export class Microsoft365Admin implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft 365 Admin',
		name: 'microsoft365Admin',
		// Single asset: the four-square mark has no background fill, so it reads
		// correctly on both the light and dark canvas.
		icon: 'file:microsoft365Admin.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Interact with Microsoft Entra ID API',
		defaults: {
			name: 'Microsoft 365 Admin',
		},
		usableAsTool: true,
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'microsoft365AdminServicePrincipalApi',
				required: true,
			},
		],
		requestDefaults: {
			headers: {
				'Content-Type': 'application/json',
			},
			baseURL:
				'={{ ($credentials.graphApiBaseUrl || "https://graph.microsoft.com").replace(/\\/+$/, "") }}/v1.0',
		},
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Group',
						value: 'group',
					},
					{
						name: 'License',
						value: 'license',
					},
					{
						name: 'User',
						value: 'user',
					},
				],
				default: 'user',
			},

			...groupOperations,
			...groupFields,
			...licenseOperations,
			...licenseFields,
			...userOperations,
			...userFields,
		],
	};

	methods = {
		loadOptions: {
			getGroupProperties,

			async getGroupPropertiesGetAll(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				// Filter items not supported for list endpoint
				return (await getGroupProperties.call(this)).filter(
					(x) =>
						![
							'allowExternalSenders',
							'autoSubscribeNewMembers',
							'hideFromAddressLists',
							'hideFromOutlookClients',
							'isSubscribedByMail',
							'unseenCount',
						].includes(x.value as string),
				);
			},

			getSubscribedSkus,

			getUserProperties,

			async getUserPropertiesGetAll(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				// Filter items not supported for list endpoint
				return (await getUserProperties.call(this)).filter(
					(x) =>
						![
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
						].includes(x.value as string),
				);
			},
		},

		listSearch: {
			getGroups,

			getUsers,
		},
	};
}
