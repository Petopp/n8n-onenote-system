const test = require('node:test');
const assert = require('node:assert');
const { OneNote } = require('../dist/nodes/OneNote/OneNote.node');
const { OneNoteTrigger } = require('../dist/nodes/OneNote/OneNoteTrigger.node');

function makeCtx(params, items, responder, extra = {}) {
	const calls = [];
	const ctx = {
		calls,
		getInputData: () => items,
		getNode: () => ({ name: 'OneNote', type: 'x', typeVersion: 1, position: [0, 0], parameters: {} }),
		getNodeParameter: (name, _i, fallback) => {
			// trigger passes (name, fallback, opts); execute passes (name, i, fallback, opts)
			if (name === 'authentication') return 'oAuth2';
			return name in params ? params[name] : fallback;
		},
		continueOnFail: () => false,
		getMode: () => 'trigger',
		getWorkflowStaticData: () => extra.staticData ?? {},
		helpers: {
			httpRequestWithAuthentication: async (_cred, options) => {
				calls.push(options);
				return { body: await responder(options), headers: { 'content-type': 'image/png' }, statusCode: 200 };
			},
			getBinaryDataBuffer: async (i, name) => Buffer.from(items[i].binary[name].data, 'base64'),
			prepareBinaryData: async (buf, fileName, mimeType) => ({
				data: buf.toString('base64'),
				fileName,
				mimeType,
			}),
		},
	};
	return ctx;
}

test('create page with image + video uploads multipart and links video', async () => {
	const items = [
		{
			json: {},
			binary: {
				img: { data: Buffer.from([1, 2, 3]).toString('base64'), mimeType: 'image/png', fileName: 'p.png' },
				vid: { data: Buffer.from('video').toString('base64'), mimeType: 'video/mp4', fileName: 'v.mp4' },
			},
		},
	];
	const ctx = makeCtx(
		{
			resource: 'page',
			operation: 'create',
			sectionId: 'SEC1',
			title: 'Hi <there>',
			contentFormat: 'markdown',
			content: '# Head',
			binaryProperties: 'img,vid',
			options: {},
		},
		items,
		async (o) => {
			if (o.method === 'PUT') return { id: 'ITEM', webUrl: 'https://1drv.ms/v' };
			if (o.url.endsWith('/createLink')) return { link: { webUrl: 'https://1drv.ms/shared' } };
			return { id: 'PAGE1', title: 'Hi', links: { oneNoteWebUrl: { href: 'https://onenote/web' } } };
		},
	);
	const [[res]] = await new OneNote().execute.call(ctx);
	assert.strictEqual(res.json.id, 'PAGE1');
	assert.strictEqual(res.json.webUrl, 'https://onenote/web');

	const put = ctx.calls.find((c) => c.method === 'PUT');
	assert.match(put.url, /me\/drive\/root:\/n8n%20OneNote%20Videos\/v\.mp4:\/content$/);
	const post = ctx.calls.find((c) => c.url.endsWith('/me/onenote/sections/SEC1/pages'));
	assert.match(post.headers['Content-Type'], /^multipart\/form-data; boundary=/);
	const body = post.body.toString('latin1');
	assert.match(body, /<title>Hi &lt;there&gt;<\/title>/);
	assert.match(body, /<h1>Head<\/h1>/);
	assert.match(body, /<img src="name:file0x/);
	assert.match(body, /href="https:\/\/1drv\.ms\/shared"/);
});

test('create plain page sends text/html body', async () => {
	const ctx = makeCtx(
		{ resource: 'page', operation: 'create', sectionId: 'S', title: 'T', contentFormat: 'text', content: 'x', binaryProperties: '', options: {} },
		[{ json: {} }],
		async () => ({ id: 'P' }),
	);
	await new OneNote().execute.call(ctx);
	assert.strictEqual(ctx.calls[0].headers['Content-Type'], 'text/html');
	assert.ok(Buffer.isBuffer(ctx.calls[0].body));
});

test('update page appends via PATCH commands', async () => {
	const ctx = makeCtx(
		{ resource: 'page', operation: 'update', pageId: 'P', updateAction: 'append', contentFormat: 'text', content: 'more', binaryProperties: '', options: {} },
		[{ json: {} }],
		async () => ({}),
	);
	const [[res]] = await new OneNote().execute.call(ctx);
	assert.strictEqual(ctx.calls[0].method, 'PATCH');
	assert.deepStrictEqual(ctx.calls[0].body, [{ target: 'body', action: 'append', content: '<p>more</p>' }]);
	assert.strictEqual(res.json.success, true);
});

test('media download stores binaries', async () => {
	const html =
		'<html><body><img src="https://graph.microsoft.com/v1.0/me/onenote/resources/0-1/$value" alt="a.png"/>' +
		'<object data-attachment="s.mp3" data="https://graph.microsoft.com/v1.0/me/onenote/resources/0-2/$value" type="audio/mpeg"></object></body></html>';
	const ctx = makeCtx(
		{ resource: 'media', operation: 'download', pageId: 'P', binaryPrefix: 'data', mediaKind: 'all' },
		[{ json: {} }],
		async (o) => (o.url.endsWith('/content') ? html : Buffer.from('BIN')),
	);
	const [[res]] = await new OneNote().execute.call(ctx);
	assert.strictEqual(res.json.count, 2);
	assert.deepStrictEqual(Object.keys(res.binary), ['data0', 'data1']);
	assert.strictEqual(res.binary.data1.fileName, 's.mp3');
});

test('trigger returns new pages and advances checkpoint', async () => {
	const staticData = { lastTimeChecked: '2026-01-01T00:00:00.000Z' };
	const ctx = makeCtx(
		{ event: 'pageCreated', scope: 'all', content: 'none' },
		[],
		async () => ({
			value: [
				{ id: 'A', createdDateTime: '2026-01-02T00:00:00Z', parentNotebook: { id: 'N', displayName: 'NB' } },
				{ id: 'B', createdDateTime: '2026-01-03T00:00:00Z' },
			],
		}),
		{ staticData },
	);
	const out = await new OneNoteTrigger().poll.call(ctx);
	assert.strictEqual(out[0].length, 2);
	assert.strictEqual(out[0][0].json.notebookName, 'NB');
	assert.strictEqual(staticData.lastTimeChecked, '2026-01-03T00:00:00Z');
	assert.strictEqual(ctx.calls[0].qs.$filter, 'createdDateTime gt 2026-01-01T00:00:00.000Z');
});

test('device login credential refreshes token and keeps rotated refresh token', async () => {
	const { OneNoteDeviceApi } = require('../dist/credentials/OneNoteDeviceApi.credentials');
	let request;
	const helper = {
		helpers: {
			httpRequest: async (o) => {
				request = o;
				return { access_token: 'AT', refresh_token: 'RT2' };
			},
		},
	};
	const out = await new OneNoteDeviceApi().preAuthentication.call(helper, {
		clientId: 'cid',
		tenant: 'consumers',
		refreshToken: 'RT1',
		enableOneDrive: true,
	});
	assert.deepStrictEqual(out, { accessToken: 'AT', refreshToken: 'RT2' });
	assert.match(request.url, /login\.microsoftonline\.com\/consumers\/oauth2\/v2\.0\/token$/);
	assert.match(request.body, /grant_type=refresh_token/);
	assert.match(request.body, /Files\.ReadWrite/);
});

test('node picks device credential when selected', async () => {
	const used = [];
	const ctx = makeCtx({ resource: 'notebook', operation: 'getAll', returnAll: true, authentication: 'deviceLogin' }, [{ json: {} }], async () => ({ value: [] }));
	ctx.getNodeParameter = (name, _i, fb) => (name === 'authentication' ? 'deviceLogin' : name === 'returnAll' ? true : name === 'resource' ? 'notebook' : name === 'operation' ? 'getAll' : fb);
	const orig = ctx.helpers.httpRequestWithAuthentication;
	ctx.helpers.httpRequestWithAuthentication = async (cred, o) => (used.push(cred), orig(cred, o));
	await new OneNote().execute.call(ctx);
	assert.deepStrictEqual(used, ['oneNoteDeviceApi']);
});

test('login helper: start returns codes, finish returns refresh token', async () => {
	const { OneNoteLogin } = require('../dist/nodes/OneNote/OneNoteLogin.node');
	const mk = (params, responder) => ({
		getNode: () => ({ name: 'L' }),
		getNodeParameter: (n) => params[n],
		helpers: { httpRequest: async (o) => responder(o) },
	});
	const [[start]] = await new OneNoteLogin().execute.call(
		mk({ operation: 'start', clientId: 'c', tenant: 'consumers', enableOneDrive: false }, () => ({
			verification_uri: 'https://microsoft.com/devicelogin',
			user_code: 'ABCD',
			device_code: 'DEV',
			expires_in: 900,
		})),
	);
	assert.strictEqual(start.json.userCode, 'ABCD');
	const [[fin]] = await new OneNoteLogin().execute.call(
		mk({ operation: 'finish', clientId: 'c', tenant: 'consumers', enableOneDrive: false, deviceCode: 'DEV', waitSeconds: 0 }, () => ({ refresh_token: 'RT' })),
	);
	assert.strictEqual(fin.json.refreshToken, 'RT');
	await assert.rejects(
		new OneNoteLogin().execute.call(
			mk({ operation: 'finish', clientId: 'c', tenant: 'consumers', enableOneDrive: false, deviceCode: 'DEV', waitSeconds: 0 }, () => ({ error: 'authorization_pending' })),
		),
		/Not confirmed yet/,
	);
});
