/// <reference path="../typings/index.d.ts" />

import { Promise } from "es6-promise";
import { INvPromiseSvc } from "./INvPromiseSvc";
import { Helpers } from "./globals";
import { NvWebSvc } from "./NvWebSvc";

export class NvListSvc implements INvPromiseSvc<SP.List> {
    private listNameOrId: string;
    private _list: SP.List;
    private _webPromise: Promise<INvPromiseSvc<SP.Web>> = null;


    //private basicProperties: Array<string> = [ "currentUser", "description", "id", "lists", "masterUrl", "title", "url"];

    constructor(listNameOrId: string, web?: Promise<INvPromiseSvc<SP.Web>>) {
        this.listNameOrId = listNameOrId;
        if(typeof web !== "undefined" && web !== null){
            this._webPromise = web;
        }
    }

    GetAsync: () => Promise<INvPromiseSvc<SP.List>> = (): Promise<INvPromiseSvc<SP.List>> => {
        return new Promise<INvPromiseSvc<SP.List>>((resolve: (listProm: Promise<INvPromiseSvc<SP.List>>) => void, reject: (error: any) => void): void => {
            try{
                if (this._webPromise == null) {
                    this._webPromise = (new NvWebSvc()).GetAsync();
                }

                Promise.resolve(this._webPromise).then((web: INvPromiseSvc<SP.Web>): void => {
                    this.Web = web;
                    this.Site = this.Web.Site;
                    this.ClientContext = this.Web.ClientContext;

                    let lists: SP.ListCollection = this.Web.Target.get_lists();
                    if (Helpers.guidRx.test(this.listNameOrId)) {
                        let listGuid: SP.Guid = new SP.Guid(this.listNameOrId);
                        this._list = lists.getById(listGuid);
                    } else {
                        this._list = lists.getByTitle(this.listNameOrId);
                    }

                    this.ClientContext.load(this._list);
                    this.ClientContext.executeQueryAsync(
                        (): void => {
                            this.List = this;
                            this.Target = this._list;
                            resolve(Promise.resolve(this));
                        },
                        (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
                            let error = new Error(args.get_message());
                            reject(error);
                        }
                    );

                });
            } catch (ex) {
                let error = new Error(ex);
                reject(error);
            }

        });
    };

    public ClientContext: SP.ClientContext = null;
    public Site: INvPromiseSvc<SP.Site> = null;
    public Web: INvPromiseSvc<SP.Web> = null;
    public List: INvPromiseSvc<SP.List> = null;
    public Target: SP.List = null;
}
