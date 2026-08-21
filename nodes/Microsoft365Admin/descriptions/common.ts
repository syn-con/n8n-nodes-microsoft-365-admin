import type { INodeProperties } from 'n8n-workflow';

export const ignoreHttpStatusErrorsConfig = {
	ignore: true as const,
	// 401 responses must be passed to requestWithAuthentication so expired OAuth2 tokens can refresh.
	except: [401],
};

/**
 * The User picker, which several resources need and which only ever differs in the
 * operations it belongs to and the wording of its hint.
 */
export const userLocator = (
	resource: string,
	operation: string | string[],
	description: string,
): INodeProperties => ({
	displayName: 'User',
	name: 'user',
	default: {
		mode: 'list',
		value: '',
	},
	description,
	displayOptions: {
		show: {
			resource: [resource],
			operation: Array.isArray(operation) ? operation : [operation],
		},
	},
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: {
				searchListMethod: 'getUsers',
				searchable: true,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			placeholder: 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315',
			type: 'string',
		},
	],
	required: true,
	type: 'resourceLocator',
});
