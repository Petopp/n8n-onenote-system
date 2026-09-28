import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';
import { graphRequest, graphRequestAll, odataString } from './transport';
import {
	bodyToHtml,
	createPage,
	downloadPageMedia,
	getPageHtml,
	htmlToText,
	param,
	prepareBinaries,
	simplify,
	updatePage,
	waitForOperation,
	type AttachmentOptions,
	type UpdateCommand,
	type VideoHandling,
} from './operations';
import { searchNotebooks, searchPages, searchSections } from './utils/listSearch';

const PAGE_EXPAND = 'parentNotebook($select=id,displayName),parentSection($select=id,displayName)';

function locator(
	name: string,
	displayName: string,
	searchListMethod: string,
	displayOptions: INodeProperties['displayOptions'],
	required = true,
	description = '',
): INodeProperties {
	return {
		displayName,
		name,
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required,
		description,
		displayOptions,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod, searchable: true },
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: '1-0123456789abcdef!123',
			},
		],
	};
}

const show = (resource: string, operation?: string[]): INodeProperties['displayOptions'] => ({
	show: { resource: [resource], ...(operation ? { operation } : {}) },
});

const attachmentOptions: INodeProperties[] = [
	{
		displayName: 'Video Handling',
		name: 'videoHandling',
		type: 'options',
		options: [
			{
				name: 'Upload to OneDrive and Link',
				value: 'onedrive',
				description: 'OneNote cannot embed videos. The video is stored in OneDrive and a link is added to the page.',
			},
			{
				name: 'Attach as File',
				value: 'attachment',
				description: 'Try to attach the video as a file (may fail for larger files)',
			},
		],
		default: 'onedrive',
	},
	{
		displayName: 'OneDrive Folder',
		name: 'oneDriveFolder',
		type: 'string',
		default: 'n8n OneNote Videos',
		description: 'Folder in OneDrive for uploaded videos',
	},
];

export class OneNote implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'OneNote System',
		name: 'oneNote',
		icon: { light: 'file:onenote.svg', dark: 'file:onenote.dark.svg' },
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description:
			'Read and write Microsoft OneNote notebooks, sections and pages including images, audio, files and videos',
		defaults: { name: 'OneNote' },
		usableAsTool: true,
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		credentials: [{ name: 'oneNoteOAuth2Api', required: true }],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{ name: 'Media', value: 'media' },
					{ name: 'Notebook', value: 'notebook' },
					{ name: 'Page', value: 'page' },
					{ name: 'Section', value: 'section' },
				],
				default: 'page',
			},

			// ---------------- Notebook ----------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: show('notebook'),
				options: [
					{ name: 'Create', value: 'create', action: 'Create a notebook' },
					{ name: 'Get', value: 'get', action: 'Get a notebook' },
					{ name: 'Get Many', value: 'getAll', action: 'Get many notebooks' },
				],
				default: 'getAll',
			},
			locator('notebookId', 'Notebook', 'searchNotebooks', show('notebook', ['get'])),
			{
				displayName: 'Name',
				name: 'displayName',
				type: 'string',
				required: true,
				default: '',
				displayOptions: show('notebook', ['create']),
			},

			// ---------------- Section ----------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: show('section'),
				options: [
					{ name: 'Create', value: 'create', action: 'Create a section' },
					{ name: 'Get', value: 'get', action: 'Get a section' },
					{ name: 'Get Many', value: 'getAll', action: 'Get many sections' },
				],
				default: 'getAll',
			},
			locator('sectionId', 'Section', 'searchSections', show('section', ['get'])),
			locator(
				'notebookId',
				'Notebook',
				'searchNotebooks',
				show('section', ['create']),
				true,
				'Notebook in which the section is created',
			),
			locator(
				'notebookFilterId',
				'Notebook',
				'searchNotebooks',
				show('section', ['getAll']),
				false,
				'Only list sections of this notebook. Leave empty for all notebooks.',
			),
			{
				displayName: 'Name',
				name: 'displayName',
				type: 'string',
				required: true,
				default: '',
				displayOptions: show('section', ['create']),
			},

			// ---------------- Page ----------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: show('page'),
				options: [
					{ name: 'Copy', value: 'copy', action: 'Copy a page to another section' },
					{ name: 'Create', value: 'create', action: 'Create a page' },
					{ name: 'Delete', value: 'delete', action: 'Delete a page' },
					{ name: 'Get', value: 'get', action: 'Get a page' },
					{ name: 'Get Many', value: 'getAll', action: 'Get many pages' },
					{ name: 'Move', value: 'move', action: 'Move a page to another section' },
					{ name: 'Update', value: 'update', action: 'Add content to a page' },
				],
				default: 'create',
			},
			locator('sectionId', 'Section', 'searchSections', show('page', ['create'])),
			locator(
				'sectionFilterId',
				'Section',
				'searchSections',
				show('page', ['getAll']),
				false,
				'Only list pages of this section. Leave empty for all pages.',
			),
			locator('pageId', 'Page', 'searchPages', show('page', ['get', 'delete', 'copy', 'move', 'update'])),
			locator(
				'targetSectionId',
				'Target Section',
				'searchSections',
				show('page', ['copy', 'move']),
				true,
				'Section the page is copied or moved to',
			),
			{
				displayName: 'Title',
				name: 'title',
				type: 'string',
				default: '',
				required: true,
				displayOptions: show('page', ['create']),
			},
			{
				displayName: 'Content Format',
				name: 'contentFormat',
				type: 'options',
				displayOptions: show('page', ['create', 'update']),
				options: [
					{ name: 'HTML', value: 'html' },
					{ name: 'Markdown', value: 'markdown' },
					{ name: 'Plain Text', value: 'text' },
				],
				default: 'text',
			},
			{
				displayName: 'Content',
				name: 'content',
				type: 'string',
				typeOptions: { rows: 6 },
				default: '',
				displayOptions: show('page', ['create', 'update']),
			},
			{
				displayName: 'Binary Properties',
				name: 'binaryProperties',
				type: 'string',
				default: '',
				placeholder: 'data',
				description:
					'Comma-separated names of binary properties (images, audio, PDFs, videos ...) to add to the page. Use * for all. Leave empty for none.',
				displayOptions: show('page', ['create', 'update']),
			},
			{
				displayName: 'Update Action',
				name: 'updateAction',
				type: 'options',
				displayOptions: show('page', ['update']),
				options: [
					{ name: 'Append to Page', value: 'append' },
					{ name: 'Insert Next to Element', value: 'insert' },
					{ name: 'Prepend to Page', value: 'prepend' },
					{ name: 'Replace Element', value: 'replace' },
				],
				default: 'append',
			},
			{
				displayName: 'Target Element ID',
				name: 'targetId',
				type: 'string',
				default: '',
				placeholder: 'p:{guid}{12}',
				description:
					'Element ID as returned by Get with "Include Element IDs". Required for Replace and Insert.',
				displayOptions: {
					show: { resource: ['page'], operation: ['update'], updateAction: ['replace', 'insert'] },
				},
			},
			{
				displayName: 'Insert Position',
				name: 'position',
				type: 'options',
				options: [
					{ name: 'After', value: 'after' },
					{ name: 'Before', value: 'before' },
				],
				default: 'after',
				displayOptions: { show: { resource: ['page'], operation: ['update'], updateAction: ['insert'] } },
			},

			// Page: Get
			{
				displayName: 'Content Output',
				name: 'contentOutput',
				type: 'options',
				displayOptions: show('page', ['get']),
				options: [
					{ name: 'HTML and Text', value: 'both' },
					{ name: 'HTML', value: 'html' },
					{ name: 'None (Metadata Only)', value: 'none' },
					{ name: 'Text', value: 'text' },
				],
				default: 'both',
			},
			{
				displayName: 'Include Element IDs',
				name: 'includeIds',
				type: 'boolean',
				default: false,
				description: 'Whether the returned HTML contains element IDs (needed for targeted updates)',
				displayOptions: show('page', ['get']),
			},
			{
				displayName: 'Download Media',
				name: 'downloadMedia',
				type: 'boolean',
				default: false,
				description: 'Whether to also download all images and attachments of the page as binary data',
				displayOptions: show('page', ['get']),
			},

			// Page: Get Many
			{
				displayName: 'Return All',
				name: 'returnAll',
				type: 'boolean',
				default: false,
				description: 'Whether to return all results or only up to a given limit',
				displayOptions: show('notebook', ['getAll']),
			},
			{
				displayName: 'Return All',
				name: 'returnAll',
				type: 'boolean',
				default: false,
				description: 'Whether to return all results or only up to a given limit',
				displayOptions: show('section', ['getAll']),
			},
			{
				displayName: 'Return All',
				name: 'returnAll',
				type: 'boolean',
				default: false,
				description: 'Whether to return all results or only up to a given limit',
				displayOptions: show('page', ['getAll']),
			},
			{
				displayName: 'Limit',
				name: 'limit',
				type: 'number',
				typeOptions: { minValue: 1 },
				default: 50,
				description: 'Max number of results to return',
				displayOptions: {
					show: { resource: ['notebook', 'section', 'page'], operation: ['getAll'], returnAll: [false] },
				},
			},
			{
				displayName: 'Filters',
				name: 'filters',
				type: 'collection',
				placeholder: 'Add Filter',
				default: {},
				displayOptions: show('page', ['getAll']),
				options: [
					{
						displayName: 'Title Contains',
						name: 'titleContains',
						type: 'string',
						default: '',
					},
					{
						displayName: 'Created After',
						name: 'createdAfter',
						type: 'dateTime',
						default: '',
					},
					{
						displayName: 'Modified After',
						name: 'modifiedAfter',
						type: 'dateTime',
						default: '',
					},
					{
						displayName: 'Order By',
						name: 'orderBy',
						type: 'options',
						options: [
							{ name: 'Created (Newest First)', value: 'createdDateTime desc' },
							{ name: 'Modified (Newest First)', value: 'lastModifiedDateTime desc' },
							{ name: 'Title (A-Z)', value: 'title' },
						],
						default: 'lastModifiedDateTime desc',
					},
				],
			},

			// Page: Create options
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				displayOptions: show('page', ['create', 'update']),
				options: [
					{
						displayName: 'Created Date',
						name: 'created',
						type: 'dateTime',
						default: '',
						description: 'Creation date stored on the page (only used by Create)',
					},
					...attachmentOptions,
				],
			},

			// ---------------- Media ----------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: show('media'),
				options: [
					{
						name: 'Add to Page',
						value: 'attach',
						action: 'Add images audio files or videos to a page',
					},
					{
						name: 'Download From Page',
						value: 'download',
						action: 'Download images and files of a page',
					},
				],
				default: 'download',
			},
			locator('pageId', 'Page', 'searchPages', show('media')),
			{
				displayName: 'Binary Properties',
				name: 'binaryProperties',
				type: 'string',
				default: 'data',
				description: 'Comma-separated names of binary properties to add. Use * for all.',
				displayOptions: show('media', ['attach']),
			},
			{
				displayName: 'Caption',
				name: 'caption',
				type: 'string',
				default: '',
				description: 'Optional text added above the media',
				displayOptions: show('media', ['attach']),
			},
			{
				displayName: 'Binary Property Prefix',
				name: 'binaryPrefix',
				type: 'string',
				default: 'data',
				description: 'Downloaded files are stored as data0, data1, ... (prefix + index).',
				displayOptions: show('media', ['download']),
			},
			{
				displayName: 'Media Type',
				name: 'mediaKind',
				type: 'options',
				options: [
					{ name: 'All', value: 'all' },
					{ name: 'Attachments (Audio, PDF, Files)', value: 'attachment' },
					{ name: 'Images', value: 'image' },
				],
				default: 'all',
				displayOptions: show('media', ['download']),
			},
			{
				displayName: 'Options',
				name: 'mediaOptions',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				displayOptions: show('media', ['attach']),
				options: attachmentOptions,
			},
		],
	};

	methods = {
		listSearch: { searchNotebooks, searchSections, searchPages },
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const out: INodeExecutionData[] = [];
		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		for (let i = 0; i < items.length; i++) {
			try {
				const enc = encodeURIComponent;
				const pair = (json: IDataObject | IDataObject[]) =>
					(Array.isArray(json) ? json : [json]).map((j) => ({
						json: j,
						pairedItem: { item: i },
					}));
				const list = async (resourcePath: string, qs: IDataObject) => {
					const returnAll = this.getNodeParameter('returnAll', i, false) as boolean;
					const limit = returnAll ? undefined : (this.getNodeParameter('limit', i, 50) as number);
					return (await graphRequestAll.call(this, resourcePath, qs, limit)).map(simplify);
				};

				if (resource === 'notebook') {
					if (operation === 'getAll') {
						out.push(...pair(await list('/me/onenote/notebooks', {})));
					} else if (operation === 'get') {
						const id = param(this, 'notebookId', i);
						out.push(...pair(simplify(await graphRequest.call(this, { method: 'GET', resource: `/me/onenote/notebooks/${enc(id)}` }))));
					} else if (operation === 'create') {
						const created = await graphRequest.call(this, {
							method: 'POST',
							resource: '/me/onenote/notebooks',
							body: { displayName: this.getNodeParameter('displayName', i) as string },
						});
						out.push(...pair(simplify(created)));
					}
				} else if (resource === 'section') {
					if (operation === 'getAll') {
						const nb = param(this, 'notebookFilterId', i, '');
						const path = nb ? `/me/onenote/notebooks/${enc(nb)}/sections` : '/me/onenote/sections';
						out.push(...pair(await list(path, { $expand: 'parentNotebook($select=id,displayName)' })));
					} else if (operation === 'get') {
						const id = param(this, 'sectionId', i);
						const s = await graphRequest.call(this, {
							method: 'GET',
							resource: `/me/onenote/sections/${enc(id)}`,
							qs: { $expand: 'parentNotebook($select=id,displayName)' },
						});
						out.push(...pair(simplify(s)));
					} else if (operation === 'create') {
						const nb = param(this, 'notebookId', i);
						const s = await graphRequest.call(this, {
							method: 'POST',
							resource: `/me/onenote/notebooks/${enc(nb)}/sections`,
							body: { displayName: this.getNodeParameter('displayName', i) as string },
						});
						out.push(...pair(simplify(s)));
					}
				} else if (resource === 'page') {
					const getAttachOptions = (): AttachmentOptions => {
						const o = this.getNodeParameter('options', i, {}) as IDataObject;
						return {
							videoHandling: (o.videoHandling as VideoHandling) ?? 'onedrive',
							oneDriveFolder: (o.oneDriveFolder as string) ?? 'n8n OneNote Videos',
						};
					};

					if (operation === 'create') {
						const sectionId = param(this, 'sectionId', i);
						const title = this.getNodeParameter('title', i) as string;
						const format = this.getNodeParameter('contentFormat', i) as string;
						const content = this.getNodeParameter('content', i, '') as string;
						const binaries = this.getNodeParameter('binaryProperties', i, '') as string;
						const options = this.getNodeParameter('options', i, {}) as IDataObject;
						let bodyHtml = bodyToHtml(format, content);
						const media = binaries
							? await prepareBinaries(this, i, binaries, getAttachOptions(), items)
							: { attachments: [], html: '' };
						bodyHtml += `\n${media.html}`;
						const created = options.created
							? new Date(options.created as string).toISOString()
							: undefined;
						const page = await createPage(this, sectionId, title, bodyHtml, media.attachments, created);
						out.push(...pair(simplify(page)));
					} else if (operation === 'get') {
						const id = param(this, 'pageId', i);
						const meta = await graphRequest.call(this, {
							method: 'GET',
							resource: `/me/onenote/pages/${enc(id)}`,
							qs: { $expand: PAGE_EXPAND },
						});
						const result = simplify(meta);
						const output = this.getNodeParameter('contentOutput', i, 'both') as string;
						if (output !== 'none') {
							const html = await getPageHtml(
								this,
								id,
								this.getNodeParameter('includeIds', i, false) as boolean,
							);
							if (output === 'html' || output === 'both') result.html = html;
							if (output === 'text' || output === 'both') result.text = htmlToText(html);
						}
						const entry: INodeExecutionData = {
							json: result,
							pairedItem: { item: i },
							binary: items[i].binary ? { ...items[i].binary } : undefined,
						};
						if (this.getNodeParameter('downloadMedia', i, false) as boolean) {
							const dl = await downloadPageMedia(this, id, entry, 'data', 'all');
							result.media = dl.files;
						}
						out.push(entry);
					} else if (operation === 'getAll') {
						const section = param(this, 'sectionFilterId', i, '');
						const filters = this.getNodeParameter('filters', i, {}) as IDataObject;
						const clauses: string[] = [];
						if (filters.titleContains) {
							clauses.push(
								`contains(tolower(title),'${odataString(String(filters.titleContains).toLowerCase())}')`,
							);
						}
						if (filters.createdAfter) {
							clauses.push(`createdDateTime gt ${new Date(filters.createdAfter as string).toISOString()}`);
						}
						if (filters.modifiedAfter) {
							clauses.push(
								`lastModifiedDateTime gt ${new Date(filters.modifiedAfter as string).toISOString()}`,
							);
						}
						const qs: IDataObject = {
							$expand: PAGE_EXPAND,
							$orderby: (filters.orderBy as string) ?? 'lastModifiedDateTime desc',
						};
						if (clauses.length) qs.$filter = clauses.join(' and ');
						const path = section ? `/me/onenote/sections/${enc(section)}/pages` : '/me/onenote/pages';
						out.push(...pair(await list(path, qs)));
					} else if (operation === 'update') {
						const id = param(this, 'pageId', i);
						const action = this.getNodeParameter('updateAction', i) as UpdateCommand['action'];
						const format = this.getNodeParameter('contentFormat', i) as string;
						const content = this.getNodeParameter('content', i, '') as string;
						const binaries = this.getNodeParameter('binaryProperties', i, '') as string;
						const media = binaries
							? await prepareBinaries(this, i, binaries, getAttachOptions(), items)
							: { attachments: [], html: '' };
						const html = `${bodyToHtml(format, content)}\n${media.html}`.trim();
						if (!html) {
							throw new NodeOperationError(this.getNode(), 'Nothing to add: content and binary properties are empty', { itemIndex: i });
						}
						const command: UpdateCommand = {
							target: 'body',
							action,
							content: html,
						};
						if (action === 'replace' || action === 'insert') {
							const target = this.getNodeParameter('targetId', i, '') as string;
							if (!target) {
								throw new NodeOperationError(this.getNode(), 'Target Element ID is required for this action', { itemIndex: i });
							}
							command.target = target.startsWith('#') ? target : `#${target}`;
							if (action === 'insert') {
								command.position = this.getNodeParameter('position', i) as 'before' | 'after';
							}
						}
						await updatePage(this, id, [command], media.attachments);
						out.push(...pair({ success: true, pageId: id, action }));
					} else if (operation === 'delete') {
						const id = param(this, 'pageId', i);
						await graphRequest.call(this, { method: 'DELETE', resource: `/me/onenote/pages/${enc(id)}` });
						out.push(...pair({ success: true, pageId: id }));
					} else if (operation === 'copy' || operation === 'move') {
						const id = param(this, 'pageId', i);
						const target = param(this, 'targetSectionId', i);
						const op = await graphRequest.call(this, {
							method: 'POST',
							resource: `/me/onenote/pages/${enc(id)}/copyToSection`,
							body: { id: target },
						});
						const page = await waitForOperation(this, op ?? {});
						if (operation === 'move') {
							await graphRequest.call(this, { method: 'DELETE', resource: `/me/onenote/pages/${enc(id)}` });
						}
						out.push(...pair(simplify(page)));
					}
				} else if (resource === 'media') {
					const id = param(this, 'pageId', i);
					if (operation === 'download') {
						const entry: INodeExecutionData = {
							json: {},
							pairedItem: { item: i },
							binary: items[i].binary ? { ...items[i].binary } : undefined,
						};
						const dl = await downloadPageMedia(
							this,
							id,
							entry,
							this.getNodeParameter('binaryPrefix', i, 'data') as string,
							this.getNodeParameter('mediaKind', i, 'all') as 'all' | 'image' | 'attachment',
						);
						entry.json = { pageId: id, count: dl.count, files: dl.files };
						out.push(entry);
					} else if (operation === 'attach') {
						const o = this.getNodeParameter('mediaOptions', i, {}) as IDataObject;
						const media = await prepareBinaries(
							this,
							i,
							this.getNodeParameter('binaryProperties', i, 'data') as string,
							{
								videoHandling: (o.videoHandling as VideoHandling) ?? 'onedrive',
								oneDriveFolder: (o.oneDriveFolder as string) ?? 'n8n OneNote Videos',
							},
							items,
						);
						const caption = this.getNodeParameter('caption', i, '') as string;
						const html = `${caption ? bodyToHtml('text', caption) : ''}\n${media.html}`.trim();
						if (!html) {
							throw new NodeOperationError(this.getNode(), 'No binary data to add', { itemIndex: i });
						}
						await updatePage(this, id, [{ target: 'body', action: 'append', content: html }], media.attachments);
						out.push(...pair({ success: true, pageId: id, added: media.attachments.length }));
					}
				}
			} catch (error) {
				if (this.continueOnFail()) {
					out.push({ json: { error: (error as Error).message }, pairedItem: { item: i } });
					continue;
				}
				throw new NodeOperationError(this.getNode(), error as Error, { itemIndex: i });
			}
		}
		return [out];
	}
}
