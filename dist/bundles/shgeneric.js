"use strict";
var es6_promise_1 = require("es6-promise");
var _ = require("lodash");
;
var Helpers = (function () {
    function Helpers() {
    }
    Helpers.guidRx = new RegExp("^[{(]{0,1}[a-fA-F0-9]{8}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{12}-{0,1}[})]{0,1}$");
    return Helpers;
}());
exports.Helpers = Helpers;
function iterateSpCollection(collection) {
    var len = collection.get_count();
    var ret = null;
    if (len > 0) {
        ret = new Array();
        var collectionEnumerator = collection.getEnumerator();
        while (collectionEnumerator.moveNext()) {
            var value = collectionEnumerator.get_current();
            ret.push(value);
        }
    }
    return ret;
}
exports.iterateSpCollection = iterateSpCollection;
function newIgnoreErrorsPromise(promise) {
    return new es6_promise_1.Promise(function (resolve, reject) {
        try {
            es6_promise_1.Promise.resolve(promise)
                .then(function (value) {
                resolve(value);
            }, function (error) {
                resolve(null);
            });
        }
        catch (ex) {
            resolve(null);
        }
    });
}
exports.newIgnoreErrorsPromise = newIgnoreErrorsPromise;
function cloneSpCamlQuery(query) {
    var ret = new SP.CamlQuery;
    var datesInUtc = query.get_datesInUtc();
    if (typeof datesInUtc === 'boolean') {
        ret.set_datesInUtc(datesInUtc);
    }
    var folderServerRelativeUrl = query.get_folderServerRelativeUrl();
    if (!_.isEmpty(folderServerRelativeUrl)) {
        ret.set_folderServerRelativeUrl(folderServerRelativeUrl);
    }
    var listItemCollectionPosition = query.get_listItemCollectionPosition();
    if (typeof listItemCollectionPosition !== 'undefined' && listItemCollectionPosition !== null) {
        ret.set_listItemCollectionPosition(listItemCollectionPosition);
    }
    var viewXml = query.get_viewXml();
    if (!_.isEmpty(viewXml)) {
        ret.set_viewXml(viewXml);
    }
    return ret;
}
exports.cloneSpCamlQuery = cloneSpCamlQuery;
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
                    var fields = iterateSpCollection(fieldsColl);
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
                var q = cloneSpCamlQuery(query);
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
                            var listItemsArray = iterateSpCollection(listItems);
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
var NvViewSvc = (function () {
    function NvViewSvc(viewNameOrId, list) {
        var _this = this;
        this._listPromise = null;
        this.GetAsync = function () {
            return new es6_promise_1.Promise(function (resolve, reject) {
                try {
                    if (_this._listPromise == null) {
                        throw new Error('The list promise is null');
                    }
                    es6_promise_1.Promise.resolve(_this._listPromise).then(function (list) {
                        _this.List = list;
                        _this.Web = _this.List.Web;
                        _this.Site = _this.List.Site;
                        _this.ClientContext = _this.List.ClientContext;
                        var views = _this.List.Target.get_views();
                        if (Helpers.guidRx.test(_this.viewNameOrId)) {
                            var viewGuid = new SP.Guid(_this.viewNameOrId);
                            _this._view = views.getById(viewGuid);
                        }
                        else {
                            _this._view = views.getByTitle(_this.viewNameOrId);
                        }
                        _this.ClientContext.load(_this._view);
                        _this.ClientContext.executeQueryAsync(function () {
                            _this.Target = _this._view;
                            resolve(es6_promise_1.Promise.resolve(_this));
                        }, function (sender, args) {
                            var error = new Error(args.get_message());
                            reject(error);
                        });
                    });
                }
                catch (ex) {
                    var error = new Error(ex);
                    reject(error);
                }
            });
        };
        this.ClientContext = null;
        this.Site = null;
        this.Web = null;
        this.List = null;
        this.Target = null;
        this.viewNameOrId = viewNameOrId;
        if (typeof list !== "undefined" && list !== null) {
            this._listPromise = list;
        }
    }
    return NvViewSvc;
}());
exports.NvViewSvc = NvViewSvc;
