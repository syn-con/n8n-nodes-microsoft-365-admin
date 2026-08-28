import {
	NodeApiError,
	NodeOperationError,
	sleep,
	type IDataObject,
	type IExecuteFunctions,
	type IN8nHttpFullResponse,
	type INodeExecutionData,
	type JsonObject,
} from 'n8n-workflow';

import { graphError, IGNORE_STATUS_ERRORS, microsoftApiRequest } from '../transport';
import { asNodeError, errorItem, isPathSafe } from './utils';

/**
 * Assign, Assign to Group and Unassign all post to `assignLicense`, which Entra ID
 * processes one write at a time per tenant: a second write arriving while the first is
 * still being applied is rejected with `Directory_ConcurrencyViolation` ("concurrent
 * requests being made to the tenant"), and the tenant can stay busy for the better part
 * of a minute.
 *
 * Declarative routing cannot survive that. n8n's RoutingNode builds one request per input
 * item and hands the whole array to `Promise.allSettled`, so a run over 20 users fires 20
 * simultaneous writes and all but one of them fail. These three operations are therefore
 * registered as `customOperations` (requires n8n 1.81 or newer) so that they can:
 *
 *   - send one request at a time,
 *   - fold every item aimed at the same user or group into a single request, since one
 *     `assignLicense` call can add and remove any number of SKUs for the same 30 seconds
 *     of tenant processing,
 *   - retry a write the tenant rejected as concurrent instead of failing the run.
 */

/* eslint-disable no-await-in-loop -- awaiting inside these loops is the point: the writes
   have to leave one at a time, and a rejected one has to wait before being resent. */

/** Entra's code for "another directory write for this tenant is still in flight". */
const TENANT_CONFLICT_CODE = 'Directory_ConcurrencyViolation';

/** Statuses that say nothing about the request itself, so resending it is safe. */
const RETRYABLE_STATUS_CODES = [429, 502, 503, 504];

const DEFAULT_MAX_RETRIES = 5;
const FIRST_RETRY_DELAY_MS = 5_000;
const MAX_RETRY_DELAY_MS = 60_000;

/** Ceiling on Max Retries, so a mistyped or expression-driven value cannot loop for hours. */
const RETRY_LIMIT = 20;

export type LicenseWriteOperation = 'assign' | 'assignGroup' | 'unassign';

/** One entry of the `addLicenses` array Graph expects. */
interface LicenseAssignment {
	skuId: string;
	disabledPlans: string[];
}

interface LicenseWrite {
	endpoint: string;
	addLicenses: LicenseAssignment[];
	removeLicenses: string[];
	/** Input items answered by this request; each one gets the response back. */
	itemIndexes: number[];
}

interface WriteOptions {
	combineItems: boolean;
	maxRetries: number;
	waitBetweenRequests: number;
}

/**
 * Reads the ID list out of a parameter.
 *
 * The SKU pickers are multi-selects, but the same parameters used to be single-selects and
 * can still be driven by an expression, so a bare string — including a comma-separated one
 * — is accepted just as readily as an array.
 */
function idList(value: unknown): string[] {
	const entries = Array.isArray(value) ? value : [value];

	const ids = entries
		.flatMap((entry) => (typeof entry === 'string' ? entry.split(',') : [entry]))
		.map((entry) => (typeof entry === 'string' || typeof entry === 'number' ? String(entry) : ''))
		.map((entry) => entry.trim())
		.filter((entry) => entry.length > 0);

	return [...new Set(ids)];
}

/** Maps every subscribed SKU to the service plans it contains. */
async function fetchServicePlans(this: IExecuteFunctions): Promise<Map<string, Set<string>>> {
	const response = (await microsoftApiRequest.call(this, 'GET', '/subscribedSkus')) as {
		value?: Array<{ skuId: string; servicePlans?: Array<{ servicePlanId: string }> }>;
	};

	return new Map(
		(response.value ?? []).map((sku) => [
			sku.skuId,
			new Set((sku.servicePlans ?? []).map((plan) => plan.servicePlanId)),
		]),
	);
}

/**
 * Splits the requested disabled plans across the selected SKUs.
 *
 * A plan ID only means something to the one SKU that contains it, so sending the whole
 * list with every SKU would have Graph reject the request as soon as more than one SKU is
 * selected. A plan that belongs to none of them is a typo worth reporting.
 */
function disabledPlansPerSku(
	this: IExecuteFunctions,
	skuIds: string[],
	disabledPlans: string[],
	plansBySku: Map<string, Set<string>>,
	itemIndex: number,
): LicenseAssignment[] {
	const assignments = skuIds.map((skuId) => ({
		skuId,
		disabledPlans: disabledPlans.filter((plan) => plansBySku.get(skuId)?.has(plan)),
	}));

	const unmatched = disabledPlans.filter(
		(plan) => !assignments.some((assignment) => assignment.disabledPlans.includes(plan)),
	);

	if (unmatched.length > 0) {
		throw new NodeOperationError(
			this.getNode(),
			`Service plan ${unmatched[0]} is not part of any selected license SKU`,
			{
				itemIndex,
				description:
					'Disabled Plans takes service plan IDs from the SKUs being assigned. Run License → Query Tenant Licenses and read `servicePlans` of the SKU to find them.',
			},
		);
	}

	return assignments;
}

/** Graph reports the same SKU twice, or in both lists, as a bad request. */
function canMerge(existing: LicenseWrite, candidate: LicenseWrite): boolean {
	for (const assignment of candidate.addLicenses) {
		if (existing.removeLicenses.includes(assignment.skuId)) {
			return false;
		}

		const clash = existing.addLicenses.find((entry) => entry.skuId === assignment.skuId);
		if (
			clash &&
			(clash.disabledPlans.length !== assignment.disabledPlans.length ||
				!clash.disabledPlans.every((plan) => assignment.disabledPlans.includes(plan)))
		) {
			return false;
		}
	}

	return !candidate.removeLicenses.some((skuId) =>
		existing.addLicenses.some((assignment) => assignment.skuId === skuId),
	);
}

/** Queues a write, folding it into an earlier one for the same target where possible. */
function queueWrite(writes: LicenseWrite[], candidate: LicenseWrite, combineItems: boolean): void {
	const existing = combineItems
		? writes.find((write) => write.endpoint === candidate.endpoint && canMerge(write, candidate))
		: undefined;

	if (!existing) {
		writes.push(candidate);
		return;
	}

	for (const assignment of candidate.addLicenses) {
		if (!existing.addLicenses.some((entry) => entry.skuId === assignment.skuId)) {
			existing.addLicenses.push(assignment);
		}
	}
	for (const skuId of candidate.removeLicenses) {
		if (!existing.removeLicenses.includes(skuId)) {
			existing.removeLicenses.push(skuId);
		}
	}
	existing.itemIndexes.push(...candidate.itemIndexes);
}

/**
 * Two licences cannot both switch on the same service plan for one user, and Entra assigns
 * what it can before refusing the rest — so this failure usually arrives with part of the
 * request already applied. Retrying it never helps.
 */
const PLAN_CONFLICT = /conflicts with service plan|mutually exclusive/i;

function isPlanConflict(response: IN8nHttpFullResponse): boolean {
	const { message, details } = graphError(response.body);

	return (
		details?.some((detail) => detail.code === 'MutuallyExclusiveViolation') === true ||
		PLAN_CONFLICT.test(message)
	);
}

/** True when the failure is about the tenant being busy rather than about the request. */
function isRetryable(response: IN8nHttpFullResponse): boolean {
	const statusCode = Number(response.statusCode);
	if (RETRYABLE_STATUS_CODES.includes(statusCode)) {
		return true;
	}
	if (statusCode !== 400 && statusCode !== 409) {
		return false;
	}

	// Entra answers a concurrent license write with a plain 400 whose code is not always
	// filled in, so the message is checked too.
	const { code, message } = graphError(response.body);
	return code === TENANT_CONFLICT_CODE || /concurrent request/i.test(message);
}

/** Graph sends `Retry-After` in whole seconds. */
function retryAfterMs(headers: IDataObject | undefined): number | undefined {
	const header = headers?.['retry-after'] ?? headers?.['Retry-After'];
	const value = Array.isArray(header) ? header[0] : header;
	if (value === undefined || value === null || value === '') {
		return undefined;
	}

	const seconds = Number(value);

	return Number.isFinite(seconds) && seconds >= 0 ? seconds * 1_000 : undefined;
}

/**
 * Exponential backoff, jittered so that two n8n workers rejected by the same tenant do
 * not line their retries back up.
 */
function backoffMs(attempt: number): number {
	const delay = Math.min(MAX_RETRY_DELAY_MS, FIRST_RETRY_DELAY_MS * 2 ** attempt);

	return Math.round(delay * (0.75 + Math.random() * 0.25));
}

function writeError(
	this: IExecuteFunctions,
	response: IN8nHttpFullResponse,
	itemIndex: number,
	attempts: number,
): NodeApiError {
	const body = response as unknown as JsonObject;

	if (isPlanConflict(response)) {
		return new NodeApiError(this.getNode(), body, {
			itemIndex,
			message: 'Microsoft refused this combination of licenses',
			description: `${graphError(response.body).message} Two SKUs cannot both enable the same service plan for one user, and Entra assigns what it can before refusing the rest — so a license it managed to apply stays applied, which is why part of the change appears to have worked. Switch the overlapping plan off on one of the SKUs under Options → Disabled Plans, or assign only one of them. License → Query Tenant Licenses lists the servicePlans each SKU contains.`,
		});
	}

	if (!isRetryable(response)) {
		return new NodeApiError(this.getNode(), body, { itemIndex });
	}

	return new NodeApiError(this.getNode(), body, {
		itemIndex,
		message: 'Microsoft Entra is still busy with another license change',
		description: `Gave up after ${attempts} ${attempts === 1 ? 'attempt' : 'attempts'}. Entra ID applies one license change at a time per tenant, so a busy tenant — another workflow, an admin in the portal, or group-based licensing catching up — blocks this one. Raise Max Retries in Options, or license through a group (Assign to Group) so that Entra fans the change out itself.`,
	});
}

/** Posts one `assignLicense` write, retrying while the tenant reports itself busy. */
async function licenseWriteRequest(
	this: IExecuteFunctions,
	write: LicenseWrite,
	maxRetries: number,
): Promise<IDataObject> {
	const body = { addLicenses: write.addLicenses, removeLicenses: write.removeLicenses };
	const itemIndex = write.itemIndexes[0];

	for (let attempt = 0; ; attempt++) {
		const response = (await microsoftApiRequest.call(this, 'POST', write.endpoint, body, {
			returnFullResponse: true,
			// Inspecting the body is what makes a concurrency rejection distinguishable from a
			// genuine bad request. 401 still has to reach the auth helper so an expired token
			// gets refreshed rather than retried as a failure.
			ignoreHttpStatusErrors: IGNORE_STATUS_ERRORS,
		})) as IN8nHttpFullResponse;

		if (Number(response.statusCode) < 400) {
			// Group assignments answer 202 Accepted, and an empty body arrives as `''` rather
			// than as an object — which would leave the item's `json` set to a string.
			const { body } = response;

			return (body && typeof body === 'object' ? body : {}) as IDataObject;
		}

		if (attempt >= maxRetries || !isRetryable(response)) {
			throw writeError.call(this, response, itemIndex, attempt + 1);
		}

		await sleep(retryAfterMs(response.headers) ?? backoffMs(attempt));
	}
}

/**
 * The knobs that govern the whole run rather than a single item, so they are read from the
 * first item the way n8n reads its own batching settings.
 */
function writeOptions(this: IExecuteFunctions): WriteOptions {
	const options = this.getNodeParameter('options', 0, {}) as {
		combineItems?: boolean;
		maxRetries?: number;
		waitBetweenRequests?: number;
	};

	return {
		combineItems: options.combineItems ?? true,
		maxRetries: Math.min(RETRY_LIMIT, Math.max(0, options.maxRetries ?? DEFAULT_MAX_RETRIES)),
		waitBetweenRequests: Math.max(0, options.waitBetweenRequests ?? 0),
	};
}

/** Carries the run-wide lookups that only some items need. */
interface RunState {
	plansBySku?: Map<string, Set<string>>;
}

/** Turns one input item into the request it asks for. */
async function buildWrite(
	this: IExecuteFunctions,
	operation: LicenseWriteOperation,
	itemIndex: number,
	state: RunState,
): Promise<LicenseWrite> {
	const targetsGroup = operation === 'assignGroup';
	const target = this.getNodeParameter(targetsGroup ? 'group' : 'user', itemIndex, '', {
		extractValue: true,
	}) as string;

	const targetName = targetsGroup ? 'group' : 'user';
	if (!target) {
		throw new NodeOperationError(this.getNode(), `No ${targetName} was given`, { itemIndex });
	}
	if (!isPathSafe(target)) {
		throw new NodeOperationError(
			this.getNode(),
			`The ${targetName} ID contains characters that are not allowed`,
			{
				itemIndex,
				description: `Use the object ID, e.g. 02bd9fd6-8f93-4758-87c3-1fb73740a315${
					targetsGroup ? '' : ', or the userPrincipalName'
				}`,
			},
		);
	}

	const skuIds = idList(this.getNodeParameter('skuId', itemIndex, []));
	if (skuIds.length === 0) {
		throw new NodeOperationError(this.getNode(), 'No license SKU was selected', {
			itemIndex,
			description: 'Pick at least one SKU, or supply its ID with an expression',
		});
	}

	const write: LicenseWrite = {
		endpoint: `/${targetsGroup ? 'groups' : 'users'}/${target}/assignLicense`,
		addLicenses: [],
		removeLicenses: [],
		itemIndexes: [itemIndex],
	};

	if (operation === 'unassign') {
		write.removeLicenses = skuIds;
		return write;
	}

	const options = this.getNodeParameter('options', itemIndex, {}) as {
		disabledPlans?: string;
		removeSkuIds?: string | string[];
	};
	const disabledPlans = idList(options.disabledPlans ?? '');

	if (disabledPlans.length === 0) {
		write.addLicenses = skuIds.map((skuId) => ({ skuId, disabledPlans: [] }));
	} else {
		// Fetched at most once per run, and only when Disabled Plans is in play.
		state.plansBySku ??= await fetchServicePlans.call(this);
		write.addLicenses = disabledPlansPerSku.call(
			this,
			skuIds,
			disabledPlans,
			state.plansBySku,
			itemIndex,
		);
	}

	// Swapping one SKU for another in a single call spares the tenant a second round of
	// license processing.
	write.removeLicenses = idList(options.removeSkuIds ?? []).filter(
		(skuId) => !skuIds.includes(skuId),
	);

	return write;
}

export async function executeLicenseWrite(
	this: IExecuteFunctions,
	operation: LicenseWriteOperation,
): Promise<INodeExecutionData[][]> {
	const items = this.getInputData();
	if (items.length === 0) {
		return [[]];
	}

	const results: INodeExecutionData[] = [];
	const writes: LicenseWrite[] = [];
	const { combineItems, maxRetries, waitBetweenRequests } = writeOptions.call(this);
	const state: RunState = {};

	for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
		try {
			const write = await buildWrite.call(this, operation, itemIndex, state);
			queueWrite(writes, write, combineItems);
		} catch (error) {
			if (!this.continueOnFail()) {
				throw asNodeError.call(this, error, itemIndex);
			}
			results[itemIndex] = errorItem(error, itemIndex);
		}
	}

	for (const [position, write] of writes.entries()) {
		if (position > 0 && waitBetweenRequests > 0) {
			await sleep(waitBetweenRequests);
		}

		try {
			const response = await licenseWriteRequest.call(this, write, maxRetries);
			for (const itemIndex of write.itemIndexes) {
				// A merged request answers several items, and items must not share one object:
				// a downstream node that writes to one would otherwise change its siblings.
				results[itemIndex] = { json: structuredClone(response), pairedItem: { item: itemIndex } };
			}
		} catch (error) {
			if (!this.continueOnFail()) {
				throw asNodeError.call(this, error, write.itemIndexes[0]);
			}
			for (const itemIndex of write.itemIndexes) {
				results[itemIndex] = errorItem(error, itemIndex);
			}
		}
	}

	// Sparse only if an item both failed and was dropped, which cannot happen — every
	// index is filled either with a response or with its error.
	return [results.filter((result) => result !== undefined)];
}
