import {
	NodeConnectionTypes,
	type IExecuteFunctions,
	type ILoadOptionsFunctions,
	type INodeExecutionData,
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
import { executeLicenseWrite } from './LicenseFunctions';

export class Microsoft365Admin implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft 365 Admin',
		name: 'microsoft365Admin',
		// The mark carries no background fill; the dark variant only lifts each fill's
		// lightness so it holds the same contrast against a dark canvas.
		icon: { light: 'file:microsoft365Admin.svg', dark: 'file:microsoft365Admin.dark.svg' },
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

	/**
	 * The license writes are the one part of this node that cannot be declarative: Entra ID
	 * applies one license change per tenant at a time and rejects the rest, while
	 * declarative routing fires every input item's request at once. See LicenseFunctions.ts.
	 */
	customOperations = {
		license: {
			async assign(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
				return executeLicenseWrite.call(this, 'assign');
			},

			async assignGroup(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
				return executeLicenseWrite.call(this, 'assignGroup');
			},

			async unassign(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
				return executeLicenseWrite.call(this, 'unassign');
			},
		},
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
