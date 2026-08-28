import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';

import { userRLC } from '../../helpers/descriptions';
import { executeLicenseWrite } from '../../helpers/licenseWrite';
import { updateDisplayOptions } from '../../helpers/utils';
import { licenseWriteOptions, skuMultiOptions } from './common';

export const properties: INodeProperties[] = [
	userRLC('User', 'The user to assign the license to'),
	skuMultiOptions('The licenses to assign.'),
	licenseWriteOptions(true),
];

const displayOptions = {
	show: {
		resource: ['license'],
		operation: ['assign'],
	},
};

export const description = updateDisplayOptions(displayOptions, properties);

/**
 * Runs once for the whole input rather than per item: the writes have to be serialized and
 * are folded together per target. See `helpers/licenseWrite.ts`.
 */
export async function executeAll(this: IExecuteFunctions): Promise<INodeExecutionData[]> {
	const [results] = await executeLicenseWrite.call(this, 'assign');
	return results;
}
