import type {
	IDataObject,
	IExecuteFunctions,
	IHookFunctions,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
	IPollFunctions,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError, sleep } from 'n8n-workflow';

export const CREDENTIAL_NAME = 'oneNoteOAuth2Api';
export const GRAPH_URL = 'https://graph.microsoft.com/v1.0';

type Ctx = IExecuteFunctions | IHookFunctions | ILoadOptionsFunctions | IPollFunctions;

export interface GraphRequest {
	method: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
	/** Path relative to the Graph root (e.g. `/me/onenote/pages`) or a full URL (nextLink) */
	resource: string;
	body?: IDataObject | IDataObject[] | Buffer;
	qs?: IDataObject;
	headers?: IDataObject;
	/** Return raw bytes (Buffer) instead of parsed JSON */
	binary?: boolean;
	/** Return the response body as text (page HTML) */
	text?: boolean;
	/** Parse the response as JSON. Default true unless `binary`/`text`. */
	json?: boolean;
	/** Return `{ body, headers, statusCode }` instead of only the body */
	fullResponse?: boolean;
}

const MAX_RETRIES = 4;

interface HttpError {
	httpCode?: string;
	statusCode?: number;
	response?: { status?: number; headers?: Record<string, string> };
	message?: string;
	description?: string;
	cause?: { response?: { data?: unknown } };
	context?: { data?: unknown };
}

function statusOf(error: HttpError): number {
	return Number(error.response?.status ?? error.statusCode ?? error.httpCode ?? 0);
}

/**
 * Perform an authenticated Microsoft Graph request. Retries on throttling (429)
 * and transient 5xx errors, honouring `Retry-After`.
 */
// The Graph response shape depends on the endpoint (JSON object, text or Buffer), callers narrow it.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function graphRequest(this: Ctx, req: GraphRequest): Promise<any> {
	const options: IHttpRequestOptions = {
		method: req.method,
		url: req.resource.startsWith('http') ? req.resource : `${GRAPH_URL}${req.resource}`,
		headers: { Accept: 'application/json', ...(req.headers ?? {}) },
		returnFullResponse: true,
		json: !req.binary && !req.text && req.json !== false,
	};
	if (req.qs && Object.keys(req.qs).length > 0) options.qs = req.qs;
	if (req.body !== undefined) options.body = req.body;
	if (req.binary) {
		options.encoding = 'arraybuffer';
		options.headers = { ...options.headers, Accept: '*/*' };
	} else if (req.text) {
		options.encoding = 'text';
		options.headers = { ...options.headers, Accept: 'text/html' };
	}

	for (let attempt = 0; ; attempt++) {
		try {
			const response = await this.helpers.httpRequestWithAuthentication.call(
				this,
				CREDENTIAL_NAME,
				options,
			);
			if (req.binary && response.body && !Buffer.isBuffer(response.body)) {
				response.body = Buffer.from(response.body as ArrayBuffer);
			}
			return req.fullResponse ? response : response.body;
		} catch (error) {
			const status = statusOf(error as HttpError);
			const retryable = status === 429 || status === 502 || status === 503 || status === 504;
			if (retryable && attempt < MAX_RETRIES) {
				const seconds = Number((error as HttpError).response?.headers?.['retry-after']);
				await sleep((Number.isFinite(seconds) && seconds > 0 ? seconds : 2 ** attempt) * 1000);
				continue;
			}
			throw new NodeApiError(this.getNode(), error as JsonObject, {
				message: graphErrorMessage(error as HttpError),
			});
		}
	}
}

function graphErrorMessage(error: HttpError): string {
	const data = (error.cause?.response?.data ?? error.context?.data) as IDataObject | undefined;
	const inner = data?.error as IDataObject | undefined;
	if (inner?.message) return inner.code ? `${String(inner.message)} (${String(inner.code)})` : String(inner.message);
	return String(error.description ?? error.message ?? 'Unknown Microsoft Graph error');
}

/** Follow `@odata.nextLink` until `limit` items are collected (0/undefined = all). */
export async function graphRequestAll(
	this: Ctx,
	resource: string,
	qs: IDataObject = {},
	limit?: number,
): Promise<IDataObject[]> {
	const results: IDataObject[] = [];
	let next: string | undefined = resource;
	let first = true;
	do {
		const page: IDataObject = await graphRequest.call(this, {
			method: 'GET',
			resource: next,
			qs: first ? { $top: limit && limit < 100 ? limit : 100, ...qs } : undefined,
		});
		first = false;
		const values = (page.value as IDataObject[]) ?? [];
		results.push(...values);
		next = page['@odata.nextLink'] as string | undefined;
		if (limit && results.length >= limit) return results.slice(0, limit);
	} while (next);
	return results;
}

/** Escape a value for use inside an OData single-quoted string literal. */
export function odataString(value: string): string {
	return value.replace(/'/g, "''");
}
