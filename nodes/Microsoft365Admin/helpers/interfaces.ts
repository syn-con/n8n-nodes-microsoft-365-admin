import type { IDataObject, IHttpRequestOptions } from 'n8n-workflow';

/** Shape of a Graph collection response used by the resource-locator searches. */
export interface DirectoryListResponse {
	value?: Array<{ id: string; displayName: string }>;
	'@odata.nextLink'?: string;
}

/**
 * The parts of a Graph request that most call sites leave alone, kept out of the
 * parameter list so the common `(method, endpoint)` and `(method, endpoint, body)`
 * calls stay short.
 */
export interface GraphRequestExtras {
	qs?: IDataObject;
	headers?: IDataObject;
	/** An absolute URL, e.g. an `@odata.nextLink`, used instead of building one. */
	url?: string;
	/** Returns `{ body, headers, statusCode }` instead of just the body. */
	returnFullResponse?: boolean;
	/** Set to inspect an error response yourself rather than having it thrown. */
	ignoreHttpStatusErrors?: IHttpRequestOptions['ignoreHttpStatusErrors'];
}

/**
 * The request fields `requestWithAuthenticationPaginated` actually reads.
 *
 * That helper is current, but n8n still types its first parameter as the deprecated
 * `IRequestOptions` and ships no paginated equivalent taking `IHttpRequestOptions`.
 * Declaring the handful of fields locally keeps the deprecated type out of this
 * package while staying structurally assignable to the helper's parameter.
 */
export interface PaginatedRequestOptions {
	method: NonNullable<IHttpRequestOptions['method']>;
	/** `IHttpRequestOptions` has no `uri`, which this helper requires. */
	uri: string;
	json: boolean;
	headers?: IDataObject;
	body?: IDataObject;
	qs?: IDataObject;
}

/** A Graph error envelope, which is absent on some gateway failures. */
export interface GraphErrorBody {
	code: string;
	message: string;
	details?: Array<{ code: string; message: string }>;
}
