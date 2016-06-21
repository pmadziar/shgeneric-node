import main = require('~lodash/index');
import { Promise } from "es6-promise";
import * as _ from "lodash";
                                                                                        

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
};

export class Helpers {
	public static guidRx: RegExp = new RegExp(`^[{(]{0,1}[a-fA-F0-9]{8}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{12}-{0,1}[})]{0,1}$`);
}


export function iterateSpCollection<T>(collection: SP.ClientObjectCollection<T>): Array<T> {
	let len = collection.get_count();
	let ret: Array<T> = null;
	if (len > 0) {
		ret = new Array<T>();
		let collectionEnumerator = collection.getEnumerator();

		while (collectionEnumerator.moveNext()) {
			let value: T = collectionEnumerator.get_current();
			ret.push(value);
		}
	}
	return ret;
} 

export function newIgnoreErrorsPromise<T>(promise: Promise<T>) {
	return new Promise<T>((resolve: (value: T) => void, reject: (error: any) => void): void => {
		try {
			Promise.resolve(promise)
				.then(
					(value:T):void => { //resolve
						resolve(value);
					},(error: any): void => { //reject
						resolve(null);
					}
				);
		} catch (ex) {
			resolve(null);
		}
	});
}

export function cloneSpCamlQuery(query: SP.CamlQuery): SP.CamlQuery {
	let ret = new SP.CamlQuery;
	let datesInUtc = query.get_datesInUtc();
	if (typeof datesInUtc === 'boolean') {
		ret.set_datesInUtc(datesInUtc);
	}

	let folderServerRelativeUrl = query.get_folderServerRelativeUrl();
	if(!_.isEmpty(folderServerRelativeUrl)){
		ret.set_folderServerRelativeUrl(folderServerRelativeUrl);
	}

	let listItemCollectionPosition = query.get_listItemCollectionPosition();
	if (typeof listItemCollectionPosition !== 'undefined' && listItemCollectionPosition !== null){
		ret.set_listItemCollectionPosition(listItemCollectionPosition);
	}

	let viewXml = query.get_viewXml();
	if (!_.isEmpty(viewXml)) {
		ret.set_viewXml(viewXml);
	}

	return ret;
}

                                                                                                                                                                                                                                                      

export class NvListUtils {
	public static getListFieldsInternalNames = (listPromise: Promise<INvPromiseSvc<SP.List>>): Promise<Array<string>> => {
		return new Promise<Array<string>>((resolve:(value:Array<string>) => Promise<Array<string>>, reject:(error:any) => void):void=>{
			Promise.resolve(listPromise).then((value: INvPromiseSvc<SP.List>):void =>{
				let ctx: SP.ClientContext = value.ClientContext;
				let lst: SP.List = value.Target;

				let fieldsColl: SP.FieldCollection = lst.get_fields();
				ctx.load(fieldsColl);
				ctx.executeQueryAsync((sender:any, args: SP.ClientRequestEventArgs):void => {
						let fields: Array<SP.Field> = iterateSpCollection<SP.Field>(fieldsColl);
						let ret: Array<string> = _.map(fields, (field: SP.Field):string => {
							return field.get_internalName();
						});
						resolve(ret);
					},
					(sender:any, args: SP.ClientRequestFailedEventArgs):void => {
						let exc = new Error(args.get_message());
						reject(exc);
					});
			}).then(undefined, (erorr:any):void => {
				reject(erorr);
			});
		});
	};

	public static rowLimitRx = /<RowLimit>\s*\d+\s*<\/RowLimit>/gm;

	public static getListItems = (listPromise: Promise<INvPromiseSvc<SP.List>>, query: SP.CamlQuery): Promise<Array<SP.ListItem>> => {
		return new Promise<Array<SP.ListItem>>((resolve: (value: Array<SP.ListItem>) => void, reject: (error: any) => void): void=> {
			let allItems: Array<SP.ListItem> = new Array<SP.ListItem>();

			Promise.resolve(listPromise).then((value: INvPromiseSvc<SP.List>): void => {
				let ctx: SP.ClientContext = value.ClientContext;
				let lst: SP.List = value.Target;
				let q: SP.CamlQuery = cloneSpCamlQuery(query);
				//q.set_viewXml(query.get_viewXml());

				let batchSize: number = 100;
				
				let viewXmlStr: string = q.get_viewXml();

				if (!_.isEmpty(viewXmlStr)) {
					if (!NvListUtils.rowLimitRx.test(viewXmlStr)) {
						let pos = viewXmlStr.indexOf("</View>");
						viewXmlStr = `${viewXmlStr.substr(0, pos)}<RowLimit>${batchSize.toString()}</RowLimit>${viewXmlStr.substr(pos)}`;
					}
				} else {
					viewXmlStr = `<View><RowLimit>${batchSize.toString()}</RowLimit></View>`;
				}

				q.set_viewXml(viewXmlStr);
				let listItems: SP.ListItemCollection;
				let position: SP.ListItemCollectionPosition = null;

				let getMoreItems = (): void => {
					listItems = lst.getItems(q);
					ctx.load(listItems);
					ctx.executeQueryAsync((sender: any, args: SP.ClientRequestEventArgs): void => {
						let itemsCount = listItems.get_count();
						if(itemsCount){
							let listItemsArray: Array<SP.ListItem> = iterateSpCollection<SP.ListItem>(listItems);
							allItems = allItems.concat(listItemsArray);
						}
						
						try {
							position = listItems.get_listItemCollectionPosition();
						} catch(exc){
							position = null;
						}

						if(position===null){
							resolve(allItems);
						} else {
							q.set_listItemCollectionPosition(position);
							getMoreItems();
						}

					},
					(sender: any, args: SP.ClientRequestFailedEventArgs): void => {
							let exc = new Error(args.get_message());
							reject(exc);
					});
					
				};

				getMoreItems();

			}).then(undefined, (erorr: any): void => {
				reject(erorr);
			});
		});
	};
}

                                                                                                                                                                                                               

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

