const test = require('node:test');
const assert = require('node:assert');
const h = require('../dist/nodes/OneNote/utils/html');

test('textToHtml escapes and splits paragraphs', () => {
	assert.strictEqual(h.textToHtml('a <b>\nc\n\nd'), '<p>a &lt;b&gt;<br />c</p>\n<p>d</p>');
});

test('markdownToHtml basics', () => {
	const html = h.markdownToHtml('# Title\n\nHello **bold** and *it* [l](https://x.y)\n\n- a\n- [x] done\n\n```\n<code>\n```');
	assert.match(html, /<h1>Title<\/h1>/);
	assert.match(html, /<b>bold<\/b>/);
	assert.match(html, /<i>it<\/i>/);
	assert.match(html, /<a href="https:\/\/x.y">l<\/a>/);
	assert.match(html, /<li>a<\/li>/);
	assert.match(html, /data-tag="to-do:completed"/);
	assert.match(html, /<pre>&lt;code&gt;<\/pre>/);
});

test('markdown blocks javascript: links and raw html', () => {
	const html = h.markdownToHtml('[x](javascript:alert(1)) <script>x</script>');
	assert.ok(!/javascript:/.test(html) || /href="#"/.test(html));
	assert.ok(!html.includes('<script>'));
});

test('htmlToText', () => {
	const t = h.htmlToText('<html><body><h1>T</h1><p>a<br/>b &amp; c</p><ul><li>x</li></ul><img alt="pic" src="u"/></body></html>');
	assert.strictEqual(t, 'T\na\nb & c\n\n- x\n[Image: pic]');
});

test('extractResources finds images and attachments', () => {
	const html =
		'<img src="https://graph.microsoft.com/v1.0/me/onenote/resources/0-abc/$value" data-fullres-src="https://graph.microsoft.com/v1.0/me/onenote/resources/0-full/$value" data-fullres-src-type="image/png" alt="x"/>' +
		'<object data-attachment="a.mp3" data="https://graph.microsoft.com/v1.0/me/onenote/resources/0-mp3/$value" type="audio/mpeg"></object>' +
		'<img src="https://example.com/x.png"/>';
	const r = h.extractResources(html);
	assert.strictEqual(r.length, 2);
	assert.strictEqual(r[0].resourceId, '0-full');
	assert.strictEqual(r[1].fileName, 'a.mp3');
});

test('multipart keeps binary bytes', () => {
	const data = Buffer.from([0, 255, 10, 13, 1]);
	const body = h.buildMultipart(
		[
			{ name: 'Presentation', contentType: 'text/html', data: '<p>x</p>' },
			{ name: 'f1', contentType: 'image/png', data, fileName: 'a.png' },
		],
		'B',
	);
	assert.ok(body.includes(data));
	assert.ok(body.toString('latin1').endsWith('--B--\r\n'));
});

test('attachmentHtml distinguishes images and files', () => {
	const img = h.attachmentHtml({ partName: 'p1', fileName: 'a.png', mimeType: 'image/png', data: Buffer.alloc(0) });
	const aud = h.attachmentHtml({ partName: 'p2', fileName: 'a.mp3', mimeType: 'audio/mpeg', data: Buffer.alloc(0) });
	assert.match(img, /<img src="name:p1"/);
	assert.match(aud, /<object data-attachment="a.mp3" data="name:p2"/);
});
