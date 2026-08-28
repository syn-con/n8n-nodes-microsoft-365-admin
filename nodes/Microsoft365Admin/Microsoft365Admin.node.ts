import {
	NodeConnectionTypes,
	type IExecuteFunctions,
	type INodeType,
	type INodeTypeDescription,
} from 'n8n-workflow';

import * as authentication from './actions/authentication';
import * as group from './actions/group';
import * as license from './actions/license';
import { router } from './actions/router';
import * as user from './actions/user';
import { listSearch, loadOptions } from './methods';

/**
 * The node is assembled from `actions/<resource>/<operation>.operation.ts`: each operation
 * owns its own parameters and its own `execute`, and `actions/router.ts` dispatches to them.
 *
 * The description lives here rather than in `actions/` because this node has a single
 * version — n8n's own linter expects the node class description to sit in the file named
 * after the node, and only a versioned node splits it into a separate base description.
 */
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
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Authentication',
						value: 'authentication',
					},
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

			...authentication.description,
			...group.description,
			...license.description,
			...user.description,
		],
	};

	methods = { loadOptions, listSearch };

	async execute(this: IExecuteFunctions) {
		return router.call(this);
	}
}
