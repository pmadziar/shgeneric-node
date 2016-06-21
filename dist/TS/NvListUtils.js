"use strict";
var _ = require("lodash");
var es6_promise_1 = require("es6-promise");
var globals_1 = require("./globals");
var NvListUtils = (function () {
    function NvListUtils() {
    }
    NvListUtils.getListFieldsInternalNames = function (listPromise) {
        return new es6_promise_1.Promise(function (resolve, reject) {
            es6_promise_1.Promise.resolve(listPromise).then(function (value) {
                var ctx = value.ClientContext;
                var lst = value.Target;
                var fieldsColl = lst.get_fields();
                ctx.load(fieldsColl);
                ctx.executeQueryAsync(function (sender, args) {
                    var fields = globals_1.iterateSpCollection(fieldsColl);
                    var ret = _.map(fields, function (field) {
                        return field.get_internalName();
                    });
                    resolve(ret);
                }, function (sender, args) {
                    var exc = new Error(args.get_message());
                    reject(exc);
                });
            }).then(undefined, function (erorr) {
                reject(erorr);
            });
        });
    };
    NvListUtils.rowLimitRx = /<RowLimit>\s*\d+\s*<\/RowLimit>/gm;
    NvListUtils.getListItems = function (listPromise, query) {
        return new es6_promise_1.Promise(function (resolve, reject) {
            var allItems = new Array();
            es6_promise_1.Promise.resolve(listPromise).then(function (value) {
                var ctx = value.ClientContext;
                var lst = value.Target;
                var q = globals_1.cloneSpCamlQuery(query);
                var batchSize = 100;
                var viewXmlStr = q.get_viewXml();
                if (!_.isEmpty(viewXmlStr)) {
                    if (!NvListUtils.rowLimitRx.test(viewXmlStr)) {
                        var pos = viewXmlStr.indexOf("</View>");
                        viewXmlStr = viewXmlStr.substr(0, pos) + "<RowLimit>" + batchSize.toString() + "</RowLimit>" + viewXmlStr.substr(pos);
                    }
                }
                else {
                    viewXmlStr = "<View><RowLimit>" + batchSize.toString() + "</RowLimit></View>";
                }
                q.set_viewXml(viewXmlStr);
                var listItems;
                var position = null;
                var getMoreItems = function () {
                    listItems = lst.getItems(q);
                    ctx.load(listItems);
                    ctx.executeQueryAsync(function (sender, args) {
                        var itemsCount = listItems.get_count();
                        if (itemsCount) {
                            var listItemsArray = globals_1.iterateSpCollection(listItems);
                            allItems = allItems.concat(listItemsArray);
                        }
                        try {
                            position = listItems.get_listItemCollectionPosition();
                        }
                        catch (exc) {
                            position = null;
                        }
                        if (position === null) {
                            resolve(allItems);
                        }
                        else {
                            q.set_listItemCollectionPosition(position);
                            getMoreItems();
                        }
                    }, function (sender, args) {
                        var exc = new Error(args.get_message());
                        reject(exc);
                    });
                };
                getMoreItems();
            }).then(undefined, function (erorr) {
                reject(erorr);
            });
        });
    };
    return NvListUtils;
}());
exports.NvListUtils = NvListUtils;
