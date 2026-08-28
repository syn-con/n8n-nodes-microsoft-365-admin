import type { INodeProperties } from 'n8n-workflow';

const ID_PLACEHOLDER = 'e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315';

/**
 * Shared parameter shapes.
 *
 * Operation files own their own parameters, but the directory pickers are the same
 * control everywhere — only the label and the hint change. Declaring them once here
 * keeps the `searchListMethod` names in a single place, so renaming a list search
 * cannot leave one operation pointing at a method that no longer exists.
 *
 * None of these carry `resource`/`operation` display options: the operation file folds
 * those in through `updateDisplayOptions`.
 */

/** The User picker. `name` varies because some operations take two different users. */
export function userRLC(displayName: string, description: string, name = 'user'): INodeProperties {
	return {
		displayName,
		name,
		default: { mode: 'list', value: '' },
		description,
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
				placeholder: ID_PLACEHOLDER,
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	};
}

/** The Group picker. */
export function groupRLC(
	displayName: string,
	description?: string,
	name = 'group',
): INodeProperties {
	return {
		displayName,
		name,
		default: { mode: 'list', value: '' },
		...(description ? { description } : {}),
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getGroups',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				placeholder: ID_PLACEHOLDER,
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	};
}

/** The registered-authentication-method picker, scoped to the chosen user and type. */
export function authenticationMethodRLC(displayName: string, description: string): INodeProperties {
	return {
		displayName,
		name: 'method',
		default: { mode: 'list', value: '' },
		description,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'getAuthenticationMethods',
					searchable: true,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				// Phone methods are the exception: their IDs are three fixed GUIDs, one per
				// phone type, the same in every tenant.
				placeholder: 'e.g. 3179e48a-750b-4051-897c-87b9720928f7 (mobile phone)',
				type: 'string',
			},
		],
		required: true,
		type: 'resourceLocator',
	};
}

/**
 * The Return All / Limit pair every collection operation offers.
 *
 * Declarative routing drove paging through `send.paginate`; the operations now call
 * `fetchCollection` (see `helpers/pagination.ts`), which reads these two parameters.
 */
export function paginationProperties(): INodeProperties[] {
	return [
		{
			displayName: 'Return All',
			name: 'returnAll',
			default: false,
			description: 'Whether to return all results or only up to a given limit',
			type: 'boolean',
		},
		{
			displayName: 'Limit',
			name: 'limit',
			default: 50,
			description: 'Max number of results to return',
			displayOptions: {
				show: {
					returnAll: [false],
				},
			},
			type: 'number',
			typeOptions: {
				minValue: 1,
			},
			validateType: 'number',
		},
	];
}

/**
 * The Output / Field Names pair that decides the Graph `$select` projection.
 *
 * `loadOptionsMethod` differs between a single read and a collection read, because the
 * collection endpoints cannot project every property a single read can.
 */
export function outputProperties(loadOptionsMethod: string): INodeProperties[] {
	return [
		{
			displayName: 'Output',
			name: 'output',
			default: 'simple',
			options: [
				{
					name: 'Simplified',
					value: 'simple',
				},
				{
					name: 'Raw',
					value: 'raw',
				},
				{
					name: 'Selected Fields',
					value: 'fields',
				},
			],
			type: 'options',
		},
		{
			displayName: 'Field Names or IDs',
			name: 'fields',
			default: [],
			description:
				'The fields to add to the output. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			displayOptions: {
				show: {
					output: ['fields'],
				},
			},
			typeOptions: {
				loadOptionsMethod,
			},
			type: 'multiOptions',
		},
	];
}
