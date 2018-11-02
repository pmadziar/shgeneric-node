import * as _ from "lodash"; 
import { Promise } from "es6-promise";
import { INvPromiseSvc } from "./INvPromiseSvc";
import { Helpers } from "./globals";
import { NvSiteSvc } from "./NvSiteSvc";

export class NvWebSvc implements INvPromiseSvc<SP.Web> {
	private webUrlOrId: string = null;
	private _web: SP.Web = null;
	private _sitePromise: Promise<INvPromiseSvc<SP.Site>> = null;

	//private basicProperties: Array<string> = [ "currentUser", "description", "id", "lists", "masterUrl", "title", "url"];

	constructor(webUrlOrId?: string, site?: Promise<INvPromiseSvc<SP.Site>>) {
		if (_.isEmpty(webUrlOrId)) {
			this.webUrlOrId = webUrlOrId;
		}
		if(typeof site !== "undefined" && site !== null){
			this._sitePromise = site;
		}
	}

	GetAsync: () => Promise<INvPromiseSvc<SP.Web>> = (): Promise<INvPromiseSvc<SP.Web>> => {
		return new Promise<INvPromiseSvc<SP.Web>>((resolve: (webProm: Promise<NvWebSvc>) => void, reject: (error: any) => void): void => {
			try{
				if (this._sitePromise == null) {
					this._sitePromise = (new NvSiteSvc(null)).GetAsync();
				}

				Promise.resolve(this._sitePromise).then((site: INvPromiseSvc<SP.Site>): void => {
					this.Site = site;
					this.ClientContext = this.Site.ClientContext;

					if (_.isEmpty(this.webUrlOrId)) {
						this._web = this.ClientContext.get_web();
					} else {
						if (Helpers.guidRx.test(this.webUrlOrId)) {
							let webGuid: SP.Guid = new SP.Guid(this.webUrlOrId);
							this._web = this.Site.Target.openWebById(webGuid);
						} else {
							this._web = this.Site.Target.openWeb(this.webUrlOrId);
						}
					}
					this.ClientContext.load(this._web);
					this.ClientContext.executeQueryAsync(
						(): void => {
							this.Web = this;
							this.Target = this._web;
							resolve(Promise.resolve(this));
						},
						(sender: any, args: SP.ClientRequestFailedEventArgs): void => {
							let error = new Error(args.get_message());
							reject(error);
						}
					);
				});
			} catch(ex){
				let error = new Error(ex);
				reject(error);
			}

		});
	};

	public ClientContext: SP.ClientContext = null;
	public Site: INvPromiseSvc<SP.Site> = null;
	public Web: INvPromiseSvc<SP.Web> = null;
	public List: INvPromiseSvc<SP.List> = null;
	public Target: SP.Web = null;
}
