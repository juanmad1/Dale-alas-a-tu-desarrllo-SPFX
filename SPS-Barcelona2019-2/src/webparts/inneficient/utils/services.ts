import { Web, sp, Items } from "@pnp/sp";
import { IListItem } from '../utils/interfaces';

export class Services {
	private _web: Web;

	constructor(siteUrl?: string) {
		this._web = sp.web;

		if (siteUrl) {
			this._web = new Web(siteUrl);
		}
	}

	public getList = (relativeUrl: string, cache: boolean, filter?: string, select?: string, expand?: string): Promise<IListItem[]> => {
		let items: Items = this._web.getList(relativeUrl).items;
		if (select) {
			items = items.select(select);
		}
		if (filter) {
			items = items.filter(filter);
		}
		if (expand) {
			items = items.expand(expand);
		}
		if (cache) {
			items = items.usingCaching();
		}
		return items.get();
	}

	public getSiteUrlFromUrl = (url: string): Promise<string> => {
		return sp.site.getWebUrlFromPageUrl(url);
	}

	public getImage = async (listUrl: string, itemId: number, columnInternalName: string): Promise<IListItem> => {
		return this._web.getList(listUrl).items.getById(itemId).fieldValuesAsHTML.select(columnInternalName).get();
	}
}