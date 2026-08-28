import { NodeOperationError, type IExecuteFunctions, type INodeExecutionData } from 'n8n-workflow';

import { asNodeError, errorItem } from '../helpers/utils';
import * as authentication from './authentication';
import * as group from './group';
import * as license from './license';
import type { Microsoft365Admin } from './node.type';
import * as user from './user';

/**
 * License writes consume the whole input at once instead of running item by item.
 *
 * Entra ID applies one license change per tenant at a time and rejects a second write
 * arriving while the first is still being applied, so these operations serialize their
 * requests and fold every item aimed at the same target into one call. Running them
 * through the per-item loop below would undo that. See `helpers/licenseWrite.ts`.
 */
const WHOLE_RUN_LICENSE_OPERATIONS = ['assign', 'assignGroup', 'unassign'] as const;

type WholeRunLicenseOperation = (typeof WHOLE_RUN_LICENSE_OPERATIONS)[number];

function isWholeRunLicenseOperation(operation: string): operation is WholeRunLicenseOperation {
	return (WHOLE_RUN_LICENSE_OPERATIONS as readonly string[]).includes(operation);
}

export async function router(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
	const items = this.getInputData();
	const returnData: INodeExecutionData[] = [];

	const resource = this.getNodeParameter<Microsoft365Admin>('resource', 0) as string;
	const operation = this.getNodeParameter('operation', 0) as string;

	const microsoft365Admin = { resource, operation } as Microsoft365Admin;

	if (microsoft365Admin.resource === 'license' && isWholeRunLicenseOperation(operation)) {
		// Reports its own per-item errors, including under "continue on fail".
		return [await license[operation].executeAll.call(this)];
	}

	let responseData: INodeExecutionData[];

	for (let i = 0; i < items.length; i++) {
		try {
			switch (microsoft365Admin.resource) {
				case 'authentication':
					responseData = await authentication[microsoft365Admin.operation].execute.call(this, i);
					break;
				case 'group':
					responseData = await group[microsoft365Admin.operation].execute.call(this, i);
					break;
				case 'license': {
					const readOperation = microsoft365Admin.operation;
					if (isWholeRunLicenseOperation(readOperation)) {
						// Returned above; narrowing the union is what leaves `execute` reachable.
						throw new NodeOperationError(
							this.getNode(),
							`The operation "${readOperation}" runs over the whole input, not per item`,
						);
					}
					responseData = await license[readOperation].execute.call(this, i);
					break;
				}
				case 'user':
					responseData = await user[microsoft365Admin.operation].execute.call(this, i);
					break;
				default:
					throw new NodeOperationError(this.getNode(), `The resource "${resource}" is not known`);
			}

			returnData.push(...responseData);
		} catch (error) {
			if (this.continueOnFail()) {
				returnData.push(errorItem(error, i));
				continue;
			}
			// A node error raised by the transport may be missing the item index.
			throw asNodeError.call(this, error, i);
		}
	}

	return [returnData];
}
