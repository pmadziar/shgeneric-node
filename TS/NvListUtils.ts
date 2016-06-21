/// <reference path="../typings/index.d.ts" />

import * as _ from "lodash"; 
import { Promise } from "es6-promise";
import { INvPromiseSvc } from "./INvPromiseSvc";
import { Helpers, iterateSpCollection, cloneSpCamlQuery } from "./globals";

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
