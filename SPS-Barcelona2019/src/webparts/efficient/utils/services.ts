import { Web, sp, Items } from "@pnp/sp/";
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

	public getMultipleImages = async (listUrl: string, itemsId: number[], columnInternalName: string): Promise<any[]> => {
		let batch = this._web.createBatch();
		let responses = new Array(itemsId.length);
		let promises = [];
		itemsId.map((id: number, i: number) => {
			// Realizamos las request de 15 en 15
			if (i % 15 == 0 && i != 0) {
				promises.push(batch.execute());
				batch = this._web.createBatch();
			}
			let response = this._web.getList(listUrl).items.getById(id).fieldValuesAsHTML.select(columnInternalName).inBatch(batch).usingCaching().get().catch(er => console.log(er));
			responses[i] = response;
		});
		promises.push(batch.execute());
		return responses;
	}
}