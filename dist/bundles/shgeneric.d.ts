/// <reference path="../../typings/index.d.ts" />
import { Promise } from "es6-promise";
export interface INvPromiseSvc<T> {
    GetAsync: () => Promise<INvPromiseSvc<T>>;
    ClientContext: SP.ClientContext;
    Site: INvPromiseSvc<SP.Site>;
    Web: INvPromiseSvc<SP.Web>;
    List: INvPromiseSvc<SP.List>;
    Target: T;
}
export interface IDictionary<T> {
    [key: string]: T;
}
export declare class Helpers {
    static guidRx: RegExp;
}
export declare function iterateSpCollection<T>(collection: SP.ClientObjectCollection<T>): Array<T>;
export declare function newIgnoreErrorsPromise<T>(promise: Promise<T>): Promise<T>;
export declare function cloneSpCamlQuery(query: SP.CamlQuery): SP.CamlQuery;
export declare class NvListUtils {
    static getListFieldsInternalNames: (listPromise: Promise<INvPromiseSvc<SP.List>>) => Promise<string[]>;
    static rowLimitRx: RegExp;
    static getListItems: (listPromise: Promise<INvPromiseSvc<SP.List>>, query: SP.CamlQuery) => Promise<SP.ListItem[]>;
}
export declare class NvViewSvc implements INvPromiseSvc<SP.View> {
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
