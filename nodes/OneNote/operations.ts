import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError, sleep } from 'n8n-workflow';
import { graphRequest } from './transport';
import {
	attachmentHtml,
	buildMultipart,
	buildPageHtml,
	extractResources,
	escapeHtml,
	htmlToText,
	isVideo,
	makeBoundary,
	markdownToHtml,
	textToHtml,
	type MultipartPart,
	type PageAttachment,
} from './utils/html';

export type VideoHandling = 'onedrive' | 'attachment';

export interface AttachmentOptions {
	videoHandling: VideoHandling;
	oneDriveFolder: string;
}

const ONEDRIVE_SIMPLE_UPLOAD_LIMIT = 250 * 1024 * 1024;

export function param<T = string>(ctx: IExecuteFunctions, name: string, i: number, fallback?: T): T {
	return ctx.getNodeParameter(name, i, (fallback ?? '') as string, { extractValue: true }) as T;
}

/** Add convenient flat fields (web URL, parent ids) to a page/section/notebook object. */
export function simplify(item: IDataObject): IDataObject {
	const links = item.links as IDataObject | undefined;
	const web = (links?.oneNoteWebUrl as IDataObject | undefined)?.href;
	const client = (links?.oneNoteClientUrl as IDataObject | undefined)?.href;
	const out: IDataObject = { ...item };
	if (web) out.webUrl = web;
	if (client) out.clientUrl = client;
	const nb = item.parentNotebook as IDataObject | undefined;
	const sec = item.parentSection as IDataObject | undefined;
	if (nb?.id) out.notebookId = nb.id;
	if (nb?.displayName) out.notebookName = nb.displayName;
	if (sec?.id) out.sectionId = sec.id;
	if (sec?.displayName) out.sectionName = sec.displayName;
	return out;
}

export function bodyToHtml(kind: string, content: string): string {
	if (kind === 'markdown') return markdownToHtml(content);
	if (kind === 'html') return content;
	return textToHtml(content);
}

/** Upload a video to OneDrive and return a shareable link. */
async function uploadVideoToOneDrive(
	ctx: IExecuteFunctions,
	fileName: string,
	data: Buffer,
	mimeType: string,
	folder: string,
): Promise<string> {
	if (data.length > ONEDRIVE_SIMPLE_UPLOAD_LIMIT) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Video "${fileName}" is larger than 250 MB. Upload it to OneDrive manually and add the link as text/HTML.`,
		);
	}
	const cleanFolder = folder.replace(/^\/+|\/+$/g, '');
	const path = `${cleanFolder ? cleanFolder + '/' : ''}${fileName}`
		.split('/')
		.map(encodeURIComponent)
		.join('/');
	const item: IDataObject = await graphRequest.call(ctx, {
		method: 'PUT',
		resource: `/me/drive/root:/${path}:/content`,
		qs: { '@microsoft.graph.conflictBehavior': 'rename' },
		headers: { 'Content-Type': mimeType || 'application/octet-stream' },
		body: data,
		json: false,
	}).then((r: unknown) => (typeof r === 'string' ? (JSON.parse(r) as IDataObject) : (r as IDataObject)));
	try {
		const link: IDataObject = await graphRequest.call(ctx, {
			method: 'POST',
			resource: `/me/drive/items/${String(item.id)}/createLink`,
			body: { type: 'view', scope: 'anonymous' },
		});
		return String((link.link as IDataObject).webUrl);
	} catch {
		return String(item.webUrl);
	}
}

/**
 * Turn binary properties of the current item into page attachments (images and files)
 * and link snippets for videos uploaded to OneDrive.
 */
export async function prepareBinaries(
	ctx: IExecuteFunctions,
	i: number,
	binaryNames: string,
	options: AttachmentOptions,
	items: INodeExecutionData[],
): Promise<{ attachments: PageAttachment[]; html: string }> {
	const attachments: PageAttachment[] = [];
	const htmlParts: string[] = [];
	const binary = items[i].binary ?? {};
	const names =
		binaryNames.trim() === '*'
			? Object.keys(binary)
			: binaryNames
					.split(',')
					.map((n) => n.trim())
					.filter(Boolean);

	for (const name of names) {
		const meta = binary[name];
		if (!meta) {
			throw new NodeOperationError(ctx.getNode(), `Item has no binary property "${name}"`, {
				itemIndex: i,
			});
		}
		const data = await ctx.helpers.getBinaryDataBuffer(i, name);
		const mimeType = meta.mimeType || 'application/octet-stream';
		const fileName = meta.fileName || `${name}.${meta.fileExtension ?? 'bin'}`;
		if (isVideo(mimeType, fileName) && options.videoHandling === 'onedrive') {
			const url = await uploadVideoToOneDrive(ctx, fileName, data, mimeType, options.oneDriveFolder);
			htmlParts.push(
				`<p>Video: <a href="${escapeHtml(url)}">${escapeHtml(fileName)}</a></p>`,
			);
		} else {
			const att: PageAttachment = {
				partName: `file${attachments.length}x${Math.random().toString(36).slice(2, 7)}`,
				fileName,
				mimeType,
				data,
			};
			attachments.push(att);
			htmlParts.push(attachmentHtml(att));
		}
	}
	return { attachments, html: htmlParts.join('\n') };
}

export function attachmentParts(attachments: PageAttachment[]): MultipartPart[] {
	return attachments.map((a) => ({
		name: a.partName,
		contentType: a.mimeType,
		data: a.data,
		fileName: a.fileName,
	}));
}

export async function createPage(
	ctx: IExecuteFunctions,
	sectionId: string,
	title: string,
	bodyHtml: string,
	attachments: PageAttachment[],
	created?: string,
): Promise<IDataObject> {
	const html = buildPageHtml(title, bodyHtml, created);
	const resource = `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`;
	if (attachments.length === 0) {
		return graphRequest.call(ctx, {
			method: 'POST',
			resource,
			headers: { 'Content-Type': 'text/html' },
			body: Buffer.from(html, 'utf8'),
		});
	}
	const boundary = makeBoundary();
	const body = buildMultipart(
		[{ name: 'Presentation', contentType: 'text/html', data: html }, ...attachmentParts(attachments)],
		boundary,
	);
	return graphRequest.call(ctx, {
		method: 'POST',
		resource,
		headers: { 'Content-Type': `multipart/form-data; boundary=${boundary}` },
		body,
	});
}

export interface UpdateCommand {
	target: string;
	action: 'append' | 'prepend' | 'replace' | 'insert';
	position?: 'before' | 'after';
	content: string;
}

export async function updatePage(
	ctx: IExecuteFunctions,
	pageId: string,
	commands: UpdateCommand[],
	attachments: PageAttachment[],
): Promise<void> {
	const resource = `/me/onenote/pages/${encodeURIComponent(pageId)}/content`;
	if (attachments.length === 0) {
		await graphRequest.call(ctx, { method: 'PATCH', resource, body: commands as unknown as IDataObject[] });
		return;
	}
	const boundary = makeBoundary();
	const body = buildMultipart(
		[
			{ name: 'Commands', contentType: 'application/json', data: JSON.stringify(commands) },
			...attachmentParts(attachments),
		],
		boundary,
	);
	await graphRequest.call(ctx, {
		method: 'PATCH',
		resource,
		headers: { 'Content-Type': `multipart/form-data; boundary=${boundary}` },
		body,
	});
}

export async function getPageHtml(
	ctx: IExecuteFunctions,
	pageId: string,
	includeIds = false,
): Promise<string> {
	const html = await graphRequest.call(ctx, {
		method: 'GET',
		resource: `/me/onenote/pages/${encodeURIComponent(pageId)}/content`,
		qs: includeIds ? { includeIDs: 'true' } : undefined,
		text: true,
	});
	return String(html);
}

const MIME_EXT: Record<string, string> = {
	'image/png': 'png',
	'image/jpeg': 'jpg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/webp': 'webp',
	'audio/mpeg': 'mp3',
	'audio/mp4': 'm4a',
	'audio/wav': 'wav',
	'application/pdf': 'pdf',
};

/** Download all OneNote-hosted images/attachments of a page into binary properties of `item`. */
export async function downloadPageMedia(
	ctx: IExecuteFunctions,
	pageId: string,
	item: INodeExecutionData,
	prefix: string,
	kind: 'all' | 'image' | 'attachment',
): Promise<{ count: number; files: IDataObject[] }> {
	const html = await getPageHtml(ctx, pageId);
	const refs = extractResources(html).filter((r) => kind === 'all' || r.kind === kind);
	item.binary = item.binary ?? {};
	const files: IDataObject[] = [];
	let n = 0;
	for (const ref of refs) {
		const res = await graphRequest.call(ctx, {
			method: 'GET',
			resource: ref.url,
			binary: true,
			fullResponse: true,
		});
		const mime = String(res.headers?.['content-type'] ?? ref.mimeType ?? 'application/octet-stream')
			.split(';')[0]
			.trim();
		const ext = MIME_EXT[mime] ?? (mime.split('/')[1] || 'bin');
		const fileName =
			ref.kind === 'attachment' && ref.fileName
				? ref.fileName
				: `${ref.fileName && ref.fileName.length < 60 && /\.\w+$/.test(ref.fileName) ? ref.fileName.replace(/\.\w+$/, '') : `image-${n + 1}`}.${ext}`;
		const key = `${prefix}${n}`;
		item.binary[key] = await ctx.helpers.prepareBinaryData(Buffer.from(res.body), fileName, mime);
		files.push({ binaryProperty: key, fileName, mimeType: mime, kind: ref.kind, resourceId: ref.resourceId });
		n++;
	}
	return { count: n, files };
}

/** Wait for an asynchronous OneNote operation (copy) and return the created page. */
export async function waitForOperation(ctx: IExecuteFunctions, operation: IDataObject): Promise<IDataObject> {
	const id = operation.id as string | undefined;
	if (!id) return operation;
	for (let attempt = 0; attempt < 30; attempt++) {
		const op: IDataObject = await graphRequest.call(ctx, {
			method: 'GET',
			resource: `/me/onenote/operations/${encodeURIComponent(id)}`,
		});
		if (op.status === 'completed') {
			if (op.resourceId) {
				return graphRequest.call(ctx, {
					method: 'GET',
					resource: `/me/onenote/pages/${String(op.resourceId)}`,
					qs: { $expand: 'parentNotebook($select=id,displayName),parentSection($select=id,displayName)' },
				});
			}
			return op;
		}
		if (op.status === 'failed') {
			throw new NodeOperationError(ctx.getNode(), `OneNote operation failed: ${JSON.stringify(op.error)}`);
		}
		await sleep(1000);
	}
	throw new NodeOperationError(ctx.getNode(), 'OneNote operation timed out');
}

export { htmlToText };
