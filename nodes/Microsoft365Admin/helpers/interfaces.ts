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

/** A Graph error envelope, which is absent on some gateway failures. */
export interface GraphErrorBody {
	code: string;
	message: string;
	details?: Array<{ code: string; message: string }>;
}
