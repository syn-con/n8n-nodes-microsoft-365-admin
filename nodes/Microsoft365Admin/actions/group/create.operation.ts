import {
	NodeApiError,
	NodeOperationError,
	type IDataObject,
	type IExecuteFunctions,
	type INodeExecutionData,
	type INodeProperties,
	type JsonObject,
} from 'n8n-workflow';

import { validateGroupNames } from '../../helpers/group';
import { deepMerge, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

/**
 * Additional fields Graph accepts on the create call itself. Everything else in the
 * collection has to be applied by a follow-up PATCH — see `execute`.
 */
const CREATE_TIME_ADDITIONAL_FIELDS = [
	'isAssignableToRole',
	'membershipRule',
	'membershipRuleProcessingState',
];

export const properties: INodeProperties[] = [
	{
		displayName: 'Group Type',
		name: 'groupType',
		default: '',
		options: [
			{
				name: 'Microsoft 365',
				value: 'Unified',
			},
			{
				name: 'Security',
				value: '',
			},
		],
		type: 'options',
	},
	{
		displayName: 'Group Name',
		name: 'displayName',
		default: '',
		description: 'The name to display in the address book for the group',
		required: true,
		type: 'string',
		validateType: 'string',
	},
	{
		displayName: 'Group Email Address',
		name: 'mailNickname',
		default: '',
		description: 'The mail alias for the group. Only enter the local-part without the domain.',
		placeholder: 'e.g. alias',
		required: true,
		type: 'string',
		validateType: 'string',
	},
	{
		displayName: 'Mail Enabled',
		name: 'mailEnabled',
		default: false,
		description: 'Whether the group is mail-enabled',
		displayOptions: {
			show: {
				groupType: ['Unified'],
			},
		},
		required: true,
		type: 'boolean',
		validateType: 'boolean',
	},
	{
		displayName: 'Membership Type',
		name: 'membershipType',
		default: '',
		options: [
			{
				name: 'Assigned',
				value: '',
				description:
					'Lets you add specific users as members of a group and have unique permissions',
			},
			{
				name: 'Dynamic',
				value: 'DynamicMembership',
				description: 'Lets you use rules to automatically add and remove users as members',
			},
		],
		type: 'options',
	},
	{
		displayName: 'Security Enabled',
		name: 'securityEnabled',
		default: true,
		description: 'Whether the group is a security group',
		displayOptions: {
			show: {
				groupType: ['Unified'],
			},
		},
		type: 'boolean',
		validateType: 'boolean',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		default: {},
		options: [
			{
				displayName: 'Assignable to Role',
				name: 'isAssignableToRole',
				default: false,
				description: 'Whether Microsoft Entra roles can be assigned to the group',
				displayOptions: {
					hide: {
						'/membershipType': ['DynamicMembership'],
					},
				},
				type: 'boolean',
				validateType: 'boolean',
			},
			{
				displayName: 'Description',
				name: 'description',
				default: '',
				description: 'Description for the group',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Membership Rule',
				name: 'membershipRule',
				default: '',
				description:
					'The <a href="https://learn.microsoft.com/en-us/entra/identity/users/groups-dynamic-membership">dynamic membership rules</a>',
				displayOptions: {
					show: {
						'/membershipType': ['DynamicMembership'],
					},
				},
				placeholder: 'e.g. user.department -eq "Marketing"',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Membership Rule Processing State',
				name: 'membershipRuleProcessingState',
				default: 'On',
				description: 'Indicates whether the dynamic membership processing is on or paused',
				displayOptions: {
					show: {
						'/membershipType': ['DynamicMembership'],
					},
				},
				options: [
					{
						name: 'On',
						value: 'On',
					},
					{
						name: 'Paused',
						value: 'Paused',
					},
				],
				type: 'options',
				validateType: 'options',
			},
			{
				displayName: 'Preferred Data Location',
				name: 'preferredDataLocation',
				default: '',
				description:
					'A property set for the group that Office 365 services use to provision the corresponding data-at-rest resources (mailbox, OneDrive, groups sites, and so on)',
				displayOptions: {
					show: {
						'/groupType': ['Unified'],
					},
				},
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Unique Name',
				name: 'uniqueName',
				default: '',
				description:
					'The unique identifier for the group, can only be updated if null, and is immutable once set',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Visibility',
				name: 'visibility',
				default: 'Public',
				description: 'Specifies the visibility of the group',
				options: [
					{
						name: 'Private',
						value: 'Private',
					},
					{
						name: 'Public',
						value: 'Public',
					},
				],
				type: 'options',
				validateType: 'options',
			},
		],
		placeholder: 'Add Field',
		type: 'collection',
	},
];

const displayOptions = {
	show: {
		resource: ['group'],
		operation: ['create'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

/** Graph enforces these combinations too, but its wording names neither parameter. */
function validateAssignableToRole(
	this: IExecuteFunctions,
	index: number,
	groupType: string,
	additionalFields: IDataObject,
): void {
	if (!additionalFields.isAssignableToRole) {
		return;
	}

	const securityEnabled = this.getNodeParameter('securityEnabled', index, true) as boolean;
	const mailEnabled = this.getNodeParameter('mailEnabled', index, false) as boolean;
	const visibility = (additionalFields.visibility ?? '') as string;

	if (!securityEnabled) {
		throw new NodeOperationError(
			this.getNode(),
			"'Security Enabled' must be set to true if 'Assignable to Role' is set",
			{ itemIndex: index },
		);
	}
	if (visibility !== 'Private') {
		throw new NodeOperationError(
			this.getNode(),
			"'Visibility' must be set to 'Private' if 'Assignable to Role' is set",
			{ itemIndex: index },
		);
	}
	if (groupType === 'Unified' && !mailEnabled) {
		throw new NodeOperationError(
			this.getNode(),
			"'Mail Enabled' must be set to true if 'Assignable to Role' is set",
			{ itemIndex: index },
		);
	}
}

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const node = this.getNode();
	const groupType = this.getNodeParameter('groupType', index, '') as string;
	const displayName = this.getNodeParameter('displayName', index) as string;
	const mailNickname = this.getNodeParameter('mailNickname', index) as string;
	const membershipType = this.getNodeParameter('membershipType', index, '') as string;
	const additionalFields = this.getNodeParameter('additionalFields', index, {}) as IDataObject;

	validateGroupNames(node, { displayName, mailNickname }, index);
	validateAssignableToRole.call(this, index, groupType, additionalFields);

	const body: IDataObject = { displayName, mailNickname };

	if (groupType) {
		body.groupTypes = [groupType];
		body.mailEnabled = this.getNodeParameter('mailEnabled', index, false) as boolean;
		body.securityEnabled = this.getNodeParameter('securityEnabled', index, true) as boolean;
	} else {
		// mailEnabled and securityEnabled are hidden for a Security group but still required.
		body.mailEnabled = false;
		body.securityEnabled = true;
	}

	if (membershipType) {
		body.groupTypes = [...((body.groupTypes as string[]) ?? []), membershipType];
	}

	for (const field of CREATE_TIME_ADDITIONAL_FIELDS) {
		if (additionalFields[field] !== undefined) {
			body[field] = additionalFields[field];
		}
	}

	const group = (await microsoftApiRequest.call(this, 'POST', '/groups', body, {
		itemIndex: index,
	})) as IDataObject;

	// The rest of Additional Fields cannot go on the create call, so they are applied
	// straight after. A failure here would otherwise leave a half-configured group behind,
	// so the group is deleted again before the error is raised.
	const patch: IDataObject = { ...additionalFields };
	for (const field of CREATE_TIME_ADDITIONAL_FIELDS) {
		delete patch[field];
	}

	if (Object.keys(patch).length > 0) {
		if (patch.assignedLabels) {
			patch.assignedLabels = [(patch.assignedLabels as IDataObject).labelValues];
		}

		try {
			await microsoftApiRequest.call(this, 'PATCH', `/groups/${group.id as string}`, patch, {
				itemIndex: index,
			});
			deepMerge(group, patch);
		} catch (error) {
			try {
				await microsoftApiRequest.call(
					this,
					'DELETE',
					`/groups/${group.id as string}`,
					{},
					{ itemIndex: index },
				);
			} catch {
				// The group could not be rolled back; the original failure is the useful one.
			}
			// Keep an already-typed node error as-is; wrap anything else so the workflow
			// always surfaces a NodeApiError rather than a bare Error.
			throw error instanceof NodeApiError
				? error
				: new NodeApiError(this.getNode(), error as JsonObject, { itemIndex: index });
		}
	}

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray(group), {
		itemData: { item: index },
	});
}
