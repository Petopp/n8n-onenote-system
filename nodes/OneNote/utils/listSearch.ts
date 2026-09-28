import type {
	IDataObject,
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';
import { graphRequest, odataString } from '../transport';

async function fetchPage(
	this: ILoadOptionsFunctions,
	resource: string,
	paginationToken: string | undefined,
	qs: IDataObject,
): Promise<{ values: IDataObject[]; next?: string }> {
	const response: IDataObject = await graphRequest.call(this, {
		method: 'GET',
		resource: paginationToken ?? resource,
		qs: paginationToken ? undefined : qs,
	});
	return {
		values: (response.value as IDataObject[]) ?? [],
		next: response['@odata.nextLink'] as string | undefined,
	};
}

function matches(name: string, filter?: string): boolean {
	return !filter || name.toLowerCase().includes(filter.toLowerCase());
}

function webUrl(item: IDataObject): string | undefined {
	return ((item.links as IDataObject | undefined)?.oneNoteWebUrl as IDataObject | undefined)?.href as
		| string
		| undefined;
}

export async function searchNotebooks(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const { values, next } = await fetchPage.call(this, '/me/onenote/notebooks', paginationToken, {
		$top: 100,
		$orderby: 'displayName',
	});
	const results: INodeListSearchItems[] = values
		.filter((n) => matches(String(n.displayName), filter))
		.map((n) => ({ name: String(n.displayName), value: String(n.id), url: webUrl(n) }));
	return { results, paginationToken: next };
}

export async function searchSections(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const { values, next } = await fetchPage.call(this, '/me/onenote/sections', paginationToken, {
		$top: 100,
		$expand: 'parentNotebook($select=displayName)',
	});
	const results: INodeListSearchItems[] = values
		.map((s) => ({
			name: `${(s.parentNotebook as IDataObject | undefined)?.displayName ?? '?'} / ${String(s.displayName)}`,
			value: String(s.id),
			url: webUrl(s),
		}))
		.filter((s) => matches(s.name, filter));
	return { results, paginationToken: next };
}

export async function searchPages(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const qs: IDataObject = {
		$top: 100,
		$orderby: 'lastModifiedDateTime desc',
		$select: 'id,title,links',
	};
	if (filter) qs.$filter = `contains(tolower(title),'${odataString(filter.toLowerCase())}')`;
	const { values, next } = await fetchPage.call(this, '/me/onenote/pages', paginationToken, qs);
	const results: INodeListSearchItems[] = values.map((p) => ({
		name: String(p.title || '(untitled)'),
		value: String(p.id),
		url: webUrl(p),
	}));
	return { results, paginationToken: next };
}
