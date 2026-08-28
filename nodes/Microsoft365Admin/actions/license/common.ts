import type { INodeProperties } from 'n8n-workflow';

/**
 * Shared by the three write operations, which all post to `assignLicense` and all have to
 * cope with Entra ID processing one license change per tenant at a time.
 *
 * `showForAssign` hides the two parameters that only make sense when licenses are being
 * added, so Unassign does not offer them.
 */
export function licenseWriteOptions(showForAssign: boolean): INodeProperties {
	const assignOnly: INodeProperties[] = [
		{
			displayName: 'Disabled Plans',
			name: 'disabledPlans',
			default: '',
			description:
				'Comma-separated service plan IDs to leave switched off. Each plan is applied to whichever selected SKU contains it. Leave empty to enable every plan.',
			placeholder: 'e.g. 8c7d2df8-86f0-4902-b2ed-a0458298f3b3',
			type: 'string',
		},
		{
			displayName: 'License SKU Names or IDs to Remove',
			name: 'removeSkuIds',
			default: [],
			description:
				'Licenses to remove in the same request that assigns the ones above — a swap costs one round of tenant processing instead of two. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			type: 'multiOptions',
			typeOptions: {
				loadOptionsMethod: 'getSubscribedSkus',
			},
		},
	];

	return {
		displayName: 'Options',
		name: 'options',
		default: {},
		options: [
			{
				displayName: 'Combine Items for the Same Target',
				name: 'combineItems',
				default: true,
				description:
					'Whether to merge every input item aimed at the same user or group into one Graph request. A single request can add and remove any number of licenses at no extra cost, so leaving this on is what makes a run over a long list of license changes finish in minutes rather than hours. Turn it off to send one request per item.',
				type: 'boolean',
			},
			...(showForAssign ? assignOnly : []),
			{
				displayName: 'Max Retries',
				name: 'maxRetries',
				default: 5,
				description:
					'How many times to retry a write that Entra rejected because the tenant was busy with another license change. Each retry waits longer than the last, up to a minute.',
				type: 'number',
				typeOptions: {
					minValue: 0,
				},
				validateType: 'number',
			},
			{
				displayName: 'Wait Between Requests',
				name: 'waitBetweenRequests',
				default: 0,
				description:
					'Milliseconds to pause between requests. Only useful when something outside this workflow is also changing licenses in the tenant and the retries are not keeping up.',
				type: 'number',
				typeOptions: {
					minValue: 0,
				},
				validateType: 'number',
			},
		],
		placeholder: 'Add option',
		type: 'collection',
	};
}

/** The SKU multi-select the write operations share. */
export function skuMultiOptions(description: string): INodeProperties {
	return {
		// A multi-select, because Graph applies every SKU in one `assignLicense` call —
		// and Entra ID spends the same tenant-wide processing time whether that call
		// carries one license or ten.
		displayName: 'License SKU Names or IDs',
		name: 'skuId',
		default: [],
		description: `${description} Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>`,
		required: true,
		type: 'multiOptions',
		typeOptions: {
			loadOptionsMethod: 'getSubscribedSkus',
		},
	};
}
