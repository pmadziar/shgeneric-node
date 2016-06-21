import { Promise } from "es6-promise";
export namespace shgeneric {

    export interface IDictionary<T> {
        [key: string]: T;
    }
    export class Helpers {
        static guidRx: RegExp;
    }
    export function iterateSpCollection<T>(collection: SP.ClientObjectCollection<T>): Array<T>;
    export function newIgnoreErrorsPromise<T>(promise: Promise<T>): Promise<T>;
    export function cloneSpCamlQuery(query: SP.CamlQuery): SP.CamlQuery;
    export interface INvPromiseSvc<T> {
        GetAsync: () => Promise<INvPromiseSvc<T>>;
        ClientContext: SP.ClientContext;
        Site: INvPromiseSvc<SP.Site>;
        Web: INvPromiseSvc<SP.Web>;
        List: INvPromiseSvc<SP.List>;
        Target: T;
    }
    export class NvSiteSvc implements INvPromiseSvc<SP.Site> {
        private siteServerRelativeUrl;
        private _site;
        constructor(siteServerRelativeUrl?: string);
        GetAsync: () => Promise<INvPromiseSvc<SP.Site>>;
        ClientContext: SP.ClientContext;
        Site: INvPromiseSvc<SP.Site>;
        Web: INvPromiseSvc<SP.Web>;
        List: INvPromiseSvc<SP.List>;
        Target: SP.Site;
    }
    export class NvWebSvc implements INvPromiseSvc<SP.Web> {
        private webUrlOrId;
        private _web;
        private _sitePromise;
        constructor(webUrlOrId?: string, site?: Promise<INvPromiseSvc<SP.Site>>);
        GetAsync: () => Promise<INvPromiseSvc<SP.Web>>;
        ClientContext: SP.ClientContext;
        Site: INvPromiseSvc<SP.Site>;
        Web: INvPromiseSvc<SP.Web>;
        List: INvPromiseSvc<SP.List>;
        Target: SP.Web;
    }
    export class NvListSvc implements INvPromiseSvc<SP.List> {
        private listNameOrId;
        private _list;
        private _webPromise;
        constructor(listNameOrId: string, web?: Promise<INvPromiseSvc<SP.Web>>);
        GetAsync: () => Promise<INvPromiseSvc<SP.List>>;
        ClientContext: SP.ClientContext;
        Site: INvPromiseSvc<SP.Site>;
        Web: INvPromiseSvc<SP.Web>;
        List: INvPromiseSvc<SP.List>;
        Target: SP.List;
    }
    export class NvViewSvc implements INvPromiseSvc<SP.View> {
        private viewNameOrId;
        private _view;
        private _listPromise;
        constructor(viewNameOrId: string, list: Promise<INvPromiseSvc<SP.List>>);
        GetAsync: () => Promise<INvPromiseSvc<SP.View>>;
        ClientContext: SP.ClientContext;
        Site: INvPromiseSvc<SP.Site>;
        Web: INvPromiseSvc<SP.Web>;
        List: INvPromiseSvc<SP.List>;
        Target: SP.View;
    }
    export class NvListUtils {
        static getListFieldsInternalNames: (listPromise: Promise<INvPromiseSvc<SP.List>>) => Promise<string[]>;
        static rowLimitRx: RegExp;
        static getListItems: (listPromise: Promise<INvPromiseSvc<SP.List>>, query: SP.CamlQuery) => Promise<SP.ListItem[]>;
    }
}