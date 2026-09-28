import type {
	IDataObject,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IPollFunctions,
} from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';
import { graphRequest, graphRequestAll } from './transport';
import { htmlToText } from './utils/html';
import { searchNotebooks, searchSections } from './utils/listSearch';
import { simplify } from './operations';

const PAGE_EXPAND = 'parentNotebook($select=id,displayName),parentSection($select=id,displayName)';

export class OneNoteTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'OneNote System Trigger',
		name: 'oneNoteTrigger',
		icon: { light: 'file:onenote.svg', dark: 'file:onenote.dark.svg' },
		group: ['trigger'],
		version: 1,
		subtitle: '={{$parameter["event"]}}',
		description: 'Starts a workflow when OneNote pages are created or modified (polling)',
		defaults: { name: 'OneNote Trigger' },
		polling: true,
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		credentials: [{ name: 'oneNoteOAuth2Api', required: true }],
		properties: [
			{
				displayName: 'Trigger On',
				name: 'event',
				type: 'options',
				options: [
					{ name: 'Page Created', value: 'pageCreated' },
					{
						name: 'Page Created or Modified',
						value: 'pageModified',
						description: 'Also fires for newly created pages',
					},
				],
				default: 'pageCreated',
			},
			{
				displayName: 'Scope',
				name: 'scope',
				type: 'options',
				options: [
					{ name: 'All Notebooks', value: 'all' },
					{ name: 'One Notebook', value: 'notebook' },
					{ name: 'One Section', value: 'section' },
				],
				default: 'all',
			},
			{
				displayName: 'Notebook',
				name: 'notebookId',
				type: 'resourceLocator',
				default: { mode: 'list', value: '' },
				required: true,
				displayOptions: { show: { scope: ['notebook'] } },
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: { searchListMethod: 'searchNotebooks', searchable: true },
					},
					{ displayName: 'By ID', name: 'id', type: 'string' },
				],
			},
			{
				displayName: 'Section',
				name: 'sectionId',
				type: 'resourceLocator',
				default: { mode: 'list', value: '' },
				required: true,
				displayOptions: { show: { scope: ['section'] } },
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: { searchListMethod: 'searchSections', searchable: true },
					},
					{ displayName: 'By ID', name: 'id', type: 'string' },
				],
			},
			{
				displayName: 'Content',
				name: 'content',
				type: 'options',
				options: [
					{ name: 'HTML and Text', value: 'both' },
					{ name: 'HTML', value: 'html' },
					{ name: 'Metadata Only', value: 'none' },
					{ name: 'Text', value: 'text' },
				],
				default: 'none',
				description: 'Whether to also fetch the page content (one extra request per page)',
			},
		],
	};

	methods = { listSearch: { searchNotebooks, searchSections } };

	async poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
		const event = this.getNodeParameter('event') as string;
		const scope = this.getNodeParameter('scope') as string;
		const contentMode = this.getNodeParameter('content', 'none') as string;
		const dateField = event === 'pageCreated' ? 'createdDateTime' : 'lastModifiedDateTime';
		const staticData = this.getWorkflowStaticData('node') as { lastTimeChecked?: string };
		const manual = this.getMode() === 'manual';

		let path = '/me/onenote/pages';
		if (scope === 'section') {
			const section = this.getNodeParameter('sectionId', '', { extractValue: true }) as string;
			path = `/me/onenote/sections/${encodeURIComponent(section)}/pages`;
		}

		const now = new Date().toISOString();
		const since = staticData.lastTimeChecked ?? now;
		const qs: IDataObject = { $expand: PAGE_EXPAND };
		if (manual) {
			qs.$orderby = `${dateField} desc`;
		} else {
			qs.$orderby = `${dateField} asc`;
			qs.$filter = `${dateField} gt ${since}`;
		}

		let pages = await graphRequestAll.call(this, path, qs, manual ? 1 : undefined);

		if (scope === 'notebook') {
			const nb = this.getNodeParameter('notebookId', '', { extractValue: true }) as string;
			pages = pages.filter((p) => (p.parentNotebook as IDataObject | undefined)?.id === nb);
		}

		if (!manual) {
			if (pages.length > 0) {
				staticData.lastTimeChecked = String(pages[pages.length - 1][dateField]);
			} else if (!staticData.lastTimeChecked) {
				staticData.lastTimeChecked = now;
			}
		}
		if (pages.length === 0) return null;

		const results: INodeExecutionData[] = [];
		for (const page of pages) {
			const json = simplify(page);
			if (contentMode !== 'none') {
				const html = String(
					await graphRequest.call(this, {
						method: 'GET',
						resource: `/me/onenote/pages/${encodeURIComponent(String(page.id))}/content`,
						text: true,
					}),
				);
				if (contentMode === 'html' || contentMode === 'both') json.html = html;
				if (contentMode === 'text' || contentMode === 'both') json.text = htmlToText(html);
			}
			results.push({ json });
		}
		return [results];
	}
}
