import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
} from 'n8n-workflow';

import { groupRLC } from '../../helpers/descriptions';
import { validateGroupNames } from '../../helpers/group';
import { assertPathSafe, updateDisplayOptions } from '../../helpers/utils';
import { microsoftApiRequest } from '../../transport';

/**
 * Graph refuses these in a PATCH that also carries anything else, so they go out in a
 * second request of their own.
 */
const SEPARATE_PATCH_FIELDS = ['allowExternalSenders', 'autoSubscribeNewMembers'];

export const properties: INodeProperties[] = [
	groupRLC('Group to Update'),
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		default: {},
		options: [
			{
				displayName: 'Allow External Senders',
				name: 'allowExternalSenders',
				default: false,
				description:
					'Whether people external to the organization can send messages to the group. Wait a few seconds before editing this field in a newly created group.',
				type: 'boolean',
				validateType: 'boolean',
			},
			{
				displayName: 'Auto Subscribe New Members',
				name: 'autoSubscribeNewMembers',
				default: false,
				description:
					'Whether new members added to the group will be auto-subscribed to receive email notifications. Wait a few seconds before editing this field in a newly created group.',
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
				displayName: 'Group Email Address',
				name: 'mailNickname',
				default: '',
				description: 'The mail alias for the group. Only enter the local-part without the domain.',
				placeholder: 'e.g. alias',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Group Name',
				name: 'displayName',
				default: '',
				description: 'The name to display in the address book for the group',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Membership Rule',
				name: 'membershipRule',
				default: '',
				description:
					'The <a href="https://learn.microsoft.com/en-us/entra/identity/users/groups-dynamic-membership">dynamic membership rules</a>',
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Membership Rule Processing State',
				name: 'membershipRuleProcessingState',
				default: 'On',
				description: 'Indicates whether the dynamic membership processing is on or paused',
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
				type: 'string',
				validateType: 'string',
			},
			{
				displayName: 'Security Enabled',
				name: 'securityEnabled',
				default: true,
				description: 'Whether the group is a security group',
				type: 'boolean',
				validateType: 'boolean',
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
		operation: ['update'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

export async function execute(
	this: IExecuteFunctions,
	index: number,
): Promise<INodeExecutionData[]> {
	const group = assertPathSafe(
		this.getNode(),
		this.getNodeParameter('group', index, '', { extractValue: true }),
		'group',
		index,
	);
	const fields = this.getNodeParameter('updateFields', index, {}) as IDataObject;

	validateGroupNames(
		this.getNode(),
		{ displayName: fields.displayName as string, mailNickname: fields.mailNickname as string },
		index,
	);

	const body: IDataObject = {};
	const separate: IDataObject = {};
	for (const [key, value] of Object.entries(fields)) {
		if (SEPARATE_PATCH_FIELDS.includes(key)) {
			separate[key] = value;
		} else {
			body[key] = value;
		}
	}

	// An empty PATCH is what Graph rejects as "Empty Payload", so a no-op update simply
	// does not go out — declarative routing had to send it and swallow the 400.
	if (Object.keys(body).length > 0) {
		await microsoftApiRequest.call(this, 'PATCH', `/groups/${group}`, body, {
			itemIndex: index,
		});
	}

	if (Object.keys(separate).length > 0) {
		await microsoftApiRequest.call(this, 'PATCH', `/groups/${group}`, separate, {
			itemIndex: index,
		});
	}

	return this.helpers.constructExecutionMetaData(this.helpers.returnJsonArray({ updated: true }), {
		itemData: { item: index },
	});
}
