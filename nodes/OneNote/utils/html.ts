/**
 * Pure helpers (no n8n dependencies) to convert between plain text, Markdown
 * and the HTML subset understood by the OneNote API.
 */

export function escapeHtml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

export function decodeEntities(value: string): string {
	return value
		.replace(/&nbsp;/g, ' ')
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&#39;|&apos;/g, "'")
		.replace(/&#(\d+);/g, (_, n: string) => String.fromCodePoint(Number(n)))
		.replace(/&#x([0-9a-f]+);/gi, (_, n: string) => String.fromCodePoint(parseInt(n, 16)))
		.replace(/&amp;/g, '&');
}

/** Plain text -> HTML. Blank lines separate paragraphs, single newlines become <br />. */
export function textToHtml(text: string): string {
	return text
		.replace(/\r\n?/g, '\n')
		.split(/\n{2,}/)
		.map((p) => p.trim())
		.filter((p) => p.length > 0)
		.map((p) => `<p>${escapeHtml(p).replace(/\n/g, '<br />')}</p>`)
		.join('\n');
}

function safeUrl(url: string): string {
	const trimmed = url.trim();
	if (/^(javascript|data|vbscript):/i.test(trimmed)) return '#';
	return trimmed;
}

function inlineMarkdown(text: string): string {
	const codes: string[] = [];
	let out = text.replace(/`([^`]+)`/g, (_, code: string) => {
		codes.push(`<code>${escapeHtml(code)}</code>`);
		return `\u0000${codes.length - 1}\u0000`;
	});
	out = escapeHtml(out);
	// images ![alt](url) are only allowed for http(s) sources
	out = out.replace(
		/!\[([^\]]*)\]\((https?:\/\/[^)\s]+)\)/g,
		(_, alt: string, url: string) => `<img src="${safeUrl(url)}" alt="${alt}" />`,
	);
	out = out.replace(
		/\[([^\]]+)\]\(([^)\s]+)\)/g,
		(_, label: string, url: string) => `<a href="${safeUrl(url)}">${label}</a>`,
	);
	out = out.replace(/\*\*([^*]+)\*\*|__([^_]+)__/g, (_, a: string, b: string) => `<b>${a ?? b}</b>`);
	out = out.replace(
		/(^|[^*\w])\*([^*\n]+)\*(?!\*)|(^|[^_\w])_([^_\n]+)_(?!\w)/g,
		(_m, p1: string, a: string, p2: string, b: string) => `${p1 ?? p2}<i>${a ?? b}</i>`,
	);
	out = out.replace(/~~([^~]+)~~/g, '<s>$1</s>');
	// eslint-disable-next-line no-control-regex
	return out.replace(/\u0000(\d+)\u0000/g, (_, i: string) => codes[Number(i)]);
}

/** Small Markdown -> HTML converter (headings, lists, quotes, code, links, emphasis, hr, tables). */
export function markdownToHtml(markdown: string): string {
	const lines = markdown.replace(/\r\n?/g, '\n').split('\n');
	const out: string[] = [];
	let i = 0;

	const isBlockStart = (l: string) =>
		/^(#{1,6})\s/.test(l) ||
		/^\s*([-*+]|\d+\.)\s+/.test(l) ||
		/^>\s?/.test(l) ||
		/^```/.test(l) ||
		/^\s*([-*_])(\s*\1){2,}\s*$/.test(l);

	while (i < lines.length) {
		const line = lines[i];

		if (line.trim() === '') {
			i++;
			continue;
		}

		const fence = /^```(\w*)/.exec(line);
		if (fence) {
			const buf: string[] = [];
			i++;
			while (i < lines.length && !/^```/.test(lines[i])) buf.push(lines[i++]);
			i++;
			out.push(`<pre>${escapeHtml(buf.join('\n'))}</pre>`);
			continue;
		}

		const heading = /^(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line);
		if (heading) {
			const level = Math.min(heading[1].length, 6);
			out.push(`<h${level}>${inlineMarkdown(heading[2])}</h${level}>`);
			i++;
			continue;
		}

		if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) {
			out.push('<hr />');
			i++;
			continue;
		}

		if (/^>\s?/.test(line)) {
			const buf: string[] = [];
			while (i < lines.length && /^>\s?/.test(lines[i])) buf.push(lines[i++].replace(/^>\s?/, ''));
			out.push(`<blockquote>${markdownToHtml(buf.join('\n'))}</blockquote>`);
			continue;
		}

		const list = /^(\s*)([-*+]|\d+\.)\s+(.*)$/.exec(line);
		if (list) {
			const ordered = /\d+\./.test(list[2]);
			const tag = ordered ? 'ol' : 'ul';
			const items: string[] = [];
			while (i < lines.length) {
				const m = /^(\s*)([-*+]|\d+\.)\s+(.*)$/.exec(lines[i]);
				if (!m) break;
				const task = /^\[( |x|X)\]\s+(.*)$/.exec(m[3]);
				if (task) {
					const checked = task[1] !== ' ';
					items.push(
						`<li data-tag="to-do${checked ? ':completed' : ''}">${inlineMarkdown(task[2])}</li>`,
					);
				} else {
					items.push(`<li>${inlineMarkdown(m[3])}</li>`);
				}
				i++;
			}
			out.push(`<${tag}>${items.join('')}</${tag}>`);
			continue;
		}

		// tables: header row + separator row
		if (
			line.includes('|') &&
			i + 1 < lines.length &&
			/^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(lines[i + 1])
		) {
			const cells = (row: string) =>
				row
					.trim()
					.replace(/^\|/, '')
					.replace(/\|$/, '')
					.split('|')
					.map((c) => inlineMarkdown(c.trim()));
			const head = cells(line);
			i += 2;
			const rows: string[][] = [];
			while (i < lines.length && lines[i].includes('|') && lines[i].trim() !== '') {
				rows.push(cells(lines[i++]));
			}
			out.push(
				'<table border="1"><tr>' +
					head.map((c) => `<td><b>${c}</b></td>`).join('') +
					'</tr>' +
					rows.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join('')}</tr>`).join('') +
					'</table>',
			);
			continue;
		}

		const buf: string[] = [];
		while (i < lines.length && lines[i].trim() !== '' && (buf.length === 0 || !isBlockStart(lines[i]))) {
			buf.push(lines[i++]);
		}
		out.push(`<p>${buf.map((l) => inlineMarkdown(l.trim())).join('<br />')}</p>`);
	}

	return out.join('\n');
}

/** OneNote page HTML -> readable plain text. */
export function htmlToText(html: string): string {
	let body = html;
	const bodyMatch = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
	if (bodyMatch) body = bodyMatch[1];
	body = body
		.replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
		.replace(/<img[^>]*?alt="([^"]*)"[^>]*>/gi, (_m, alt: string) => (alt ? `[Image: ${alt}]` : '[Image]'))
		.replace(/<img[^>]*>/gi, '[Image]')
		.replace(
			/<object[^>]*data-attachment="([^"]*)"[^>]*>(?:<\/object>)?/gi,
			(_m, name: string) => `[Attachment: ${decodeEntities(name)}]`,
		)
		.replace(/<a[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (_m, href: string, label: string) => {
			const text = label.replace(/<[^>]+>/g, '').trim();
			const url = decodeEntities(href);
			return text && text !== url ? `${text} (${url})` : url;
		})
		.replace(/<li[^>]*>/gi, '\n- ')
		.replace(/<br\s*\/?>/gi, '\n')
		.replace(/<\/(p|div|h[1-6]|tr|ul|ol|table|blockquote|pre)>/gi, '\n')
		.replace(/<\/t[dh]>/gi, '\t')
		.replace(/<[^>]+>/g, '');
	return decodeEntities(body)
		.replace(/[ \t]+\n/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
}

export interface PageResourceRef {
	kind: 'image' | 'attachment';
	/** Absolute Graph URL of the resource ($value endpoint) */
	url: string;
	/** OneNote resource id */
	resourceId: string;
	fileName?: string;
	mimeType?: string;
}

function attr(tag: string, name: string): string | undefined {
	const m = new RegExp(`\\s${name}\\s*=\\s*"([^"]*)"`, 'i').exec(tag);
	return m ? decodeEntities(m[1]) : undefined;
}

/** Find images and file attachments (audio, pdf, ...) hosted by OneNote inside page HTML. */
export function extractResources(html: string): PageResourceRef[] {
	const result: PageResourceRef[] = [];
	const seen = new Set<string>();
	const add = (ref: PageResourceRef) => {
		if (seen.has(ref.url)) return;
		seen.add(ref.url);
		result.push(ref);
	};
	const idOf = (url: string) => /\/resources\/([^/]+)\/\$value/i.exec(url)?.[1] ?? '';

	for (const m of html.matchAll(/<img\b[^>]*>/gi)) {
		const tag = m[0];
		// prefer the full resolution image when OneNote provides it
		const url = attr(tag, 'data-fullres-src') ?? attr(tag, 'src');
		if (!url || !/\/onenote\/resources\//i.test(url)) continue;
		add({
			kind: 'image',
			url,
			resourceId: idOf(url),
			mimeType: attr(tag, 'data-fullres-src-type') ?? attr(tag, 'data-src-type'),
			fileName: attr(tag, 'alt'),
		});
	}
	for (const m of html.matchAll(/<object\b[^>]*>/gi)) {
		const tag = m[0];
		const url = attr(tag, 'data');
		if (!url || !/\/onenote\/resources\//i.test(url)) continue;
		add({
			kind: 'attachment',
			url,
			resourceId: idOf(url),
			fileName: attr(tag, 'data-attachment'),
			mimeType: attr(tag, 'type'),
		});
	}
	return result;
}

export interface PageAttachment {
	/** Multipart part name (unique, alphanumeric) */
	partName: string;
	fileName: string;
	mimeType: string;
	data: Buffer;
}

/** HTML snippet that references an attachment uploaded in the same multipart request. */
export function attachmentHtml(att: PageAttachment): string {
	if (att.mimeType.startsWith('image/')) {
		return `<p><img src="name:${att.partName}" alt="${escapeHtml(att.fileName)}" data-src-type="${escapeHtml(att.mimeType)}" /></p>`;
	}
	return `<p><object data-attachment="${escapeHtml(att.fileName)}" data="name:${att.partName}" type="${escapeHtml(att.mimeType)}" /></p>`;
}

export function buildPageHtml(title: string, bodyHtml: string, created?: string): string {
	const createdMeta = created ? `<meta name="created" content="${escapeHtml(created)}" />` : '';
	return `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title>${createdMeta}</head><body>${bodyHtml}</body></html>`;
}

export interface MultipartPart {
	name: string;
	contentType: string;
	data: Buffer | string;
	fileName?: string;
}

/** Build a multipart/form-data body by hand so binary parts are sent byte-exact. */
export function buildMultipart(parts: MultipartPart[], boundary: string): Buffer {
	const chunks: Buffer[] = [];
	for (const part of parts) {
		let header = `--${boundary}\r\nContent-Disposition: form-data; name="${part.name}"`;
		if (part.fileName) header += `; filename="${part.fileName.replace(/"/g, '')}"`;
		header += `\r\nContent-Type: ${part.contentType}\r\n\r\n`;
		chunks.push(Buffer.from(header, 'utf8'));
		chunks.push(Buffer.isBuffer(part.data) ? part.data : Buffer.from(part.data, 'utf8'));
		chunks.push(Buffer.from('\r\n', 'utf8'));
	}
	chunks.push(Buffer.from(`--${boundary}--\r\n`, 'utf8'));
	return Buffer.concat(chunks);
}

export function makeBoundary(): string {
	return `n8nOneNote${Date.now().toString(16)}${Math.random().toString(16).slice(2, 10)}`;
}

export function isVideo(mimeType: string, fileName = ''): boolean {
	return mimeType.startsWith('video/') || /\.(mp4|mov|m4v|avi|mkv|webm|wmv)$/i.test(fileName);
}
