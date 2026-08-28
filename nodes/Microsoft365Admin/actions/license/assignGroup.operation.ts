import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { groupRLC } from '../../helpers/descriptions';
import { executeLicenseWrite } from '../../helpers/licenseWrite';
import { updateDisplayOptions } from '../../helpers/utils';
import { licenseWriteOptions, skuMultiOptions } from './common';

export const properties: INodeProperties[] = [
	groupRLC('Group', 'The group to assign the license to'),
	skuMultiOptions('The licenses to assign.'),
	licenseWriteOptions(true),
];

const displayOptions = {
	show: {
		resource: ['license'],
		operation: ['assignGroup'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

/** Runs once for the whole input; see `assign.operation.ts`. */
export async function executeAll(this: IExecuteFunctions): Promise<INodeExecutionData[]> {
	const [results] = await executeLicenseWrite.call(this, 'assignGroup');
	return results;
}
