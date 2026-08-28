import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { executeLicenseWrite } from '../../helpers/licenseWrite';
import { updateDisplayOptions } from '../../helpers/utils';
import { licenseWriteOptions, skuMultiOptions } from './common';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user to remove the license from'),
	skuMultiOptions('The licenses to remove.'),
	licenseWriteOptions(false),
];

const displayOptions = {
	show: {
		resource: ['license'],
		operation: ['unassign'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

/** Runs once for the whole input; see `assign.operation.ts`. */
export async function executeAll(this: IExecuteFunctions): Promise<INodeExecutionData[]> {
	const [results] = await executeLicenseWrite.call(this, 'unassign');
	return results;
}
