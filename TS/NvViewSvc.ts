/// <reference path="../typings/index.d.ts" />

import * as _ from "lodash"; 
import { Promise } from "es6-promise";
import { INvPromiseSvc } from "./INvPromiseSvc";
import { Helpers } from "./globals";

export class NvViewSvc implements INvPromiseSvc<SP.View> {
    private viewNameOrId: string;
    private _view: SP.View;
    private _listPromise: Promise<INvPromiseSvc<SP.List>> = null;


    //private basicProperties: Array<string> = [ "currentUser", "description", "id", "lists", "masterUrl", "title", "url"];

    constructor(viewNameOrId: string, list: Promise<INvPromiseSvc<SP.List>>) {
        this.viewNameOrId = viewNameOrId;
        if(typeof list !== "undefined" && list !== null){
            this._listPromise = list;
        }
    }

    GetAsync: () => Promise<INvPromiseSvc<SP.View>> = (): Promise<INvPromiseSvc<SP.View>> => {
        return new Promise<INvPromiseSvc<SP.View>>((resolve: (listProm: Promise<INvPromiseSvc<SP.View>>) => void, reject: (error: any) => void): void => {
            try{
                if (this._listPromise == null) {
                    throw new Error('The list promise is null');
                }

                Promise.resolve(this._listPromise).then((list: INvPromiseSvc<SP.List>): void => {
                    this.List = list;
                    this.Web = this.List.Web;
                    this.Site = this.List.Site;
                    this.ClientContext = this.List.ClientContext;

                    let views: SP.ViewCollection = this.List.Target.get_views();
                    if (Helpers.guidRx.test(this.viewNameOrId)) {
                        let viewGuid: SP.Guid = new SP.Guid(this.viewNameOrId);
                        this._view = views.getById(viewGuid);
                    } else {
                        this._view = views.getByTitle(this.viewNameOrId);
                    }

                    this.ClientContext.load(this._view);
                    this.ClientContext.executeQueryAsync(
                        (): void => {
                            this.Target = this._view;
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
    public Target: SP.View = null;
}
