var shgeneric;
(function (shgeneric) {
    ;
    var Helpers = (function () {
        function Helpers() {
        }
        Helpers.guidRx = new RegExp("^[{(]{0,1}[a-fA-F0-9]{8}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{12}-{0,1}[})]{0,1}$");
        return Helpers;
    }());
    shgeneric.Helpers = Helpers;
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
    shgeneric.iterateSpCollection = iterateSpCollection;
    function newIgnoreErrorsPromise(promise) {
        return new Promise(function (resolve, reject) {
            try {
                Promise.resolve(promise)
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
    shgeneric.newIgnoreErrorsPromise = newIgnoreErrorsPromise;
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
    shgeneric.cloneSpCamlQuery = cloneSpCamlQuery;


    var NvSiteSvc = (function () {
        function NvSiteSvc(siteServerRelativeUrl) {
            var _this = this;
            this.GetAsync = function () {
                return new Promise(function (resolve, reject) {
                    if (_.isEmpty(_this.siteServerRelativeUrl)) {
                        _this.ClientContext = SP.ClientContext.get_current();
                    }
                    else {
                        _this.ClientContext = new SP.ClientContext(_this.siteServerRelativeUrl);
                    }
                    _this._site = _this.ClientContext.get_site();
                    _this.ClientContext.load(_this._site);
                    _this.ClientContext.executeQueryAsync(function () {
                        _this.Target = _this._site;
                        _this.Site = _this;
                        resolve(Promise.resolve(_this));
                    }, function (sender, args) {
                        var error = new Error(args.get_message());
                        reject(error);
                    });
                });
            };
            this.ClientContext = null;
            this.Site = null;
            this.Web = null;
            this.List = null;
            this.Target = null;
            if (!_.isEmpty(siteServerRelativeUrl)) {
                this.siteServerRelativeUrl = siteServerRelativeUrl;
            }
        }
        return NvSiteSvc;
    }());
    shgeneric.NvSiteSvc = NvSiteSvc;

    var NvWebSvc = (function () {
        function NvWebSvc(webUrlOrId, site) {
            var _this = this;
            this.webUrlOrId = null;
            this._web = null;
            this._sitePromise = null;
            this.GetAsync = function () {
                return new Promise(function (resolve, reject) {
                    try {
                        if (_this._sitePromise == null) {
                            _this._sitePromise = (new shgeneric.NvSiteSvc(null)).GetAsync();
                        }
                        Promise.resolve(_this._sitePromise).then(function (site) {
                            _this.Site = site;
                            _this.ClientContext = _this.Site.ClientContext;
                            if (_.isEmpty(_this.webUrlOrId)) {
                                _this._web = _this.ClientContext.get_web();
                            }
                            else {
                                if (shgeneric.Helpers.guidRx.test(_this.webUrlOrId)) {
                                    var webGuid = new SP.Guid(_this.webUrlOrId);
                                    _this._web = _this.Site.Target.openWebById(webGuid);
                                }
                                else {
                                    _this._web = _this.Site.Target.openWeb(_this.webUrlOrId);
                                }
                            }
                            _this.ClientContext.load(_this._web);
                            _this.ClientContext.executeQueryAsync(function () {
                                _this.Web = _this;
                                _this.Target = _this._web;
                                resolve(Promise.resolve(_this));
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
            if (_.isEmpty(webUrlOrId)) {
                this.webUrlOrId = webUrlOrId;
            }
            if (typeof site !== "undefined" && site !== null) {
                this._sitePromise = site;
            }
        }
        return NvWebSvc;
    }());
    shgeneric.NvWebSvc = NvWebSvc;

    var NvListSvc = (function () {
        function NvListSvc(listNameOrId, web) {
            var _this = this;
            this._webPromise = null;
            this.GetAsync = function () {
                return new Promise(function (resolve, reject) {
                    try {
                        if (_this._webPromise == null) {
                            _this._webPromise = (new shgeneric.NvWebSvc()).GetAsync();
                        }
                        Promise.resolve(_this._webPromise).then(function (web) {
                            _this.Web = web;
                            _this.Site = _this.Web.Site;
                            _this.ClientContext = _this.Web.ClientContext;
                            var lists = _this.Web.Target.get_lists();
                            if (shgeneric.Helpers.guidRx.test(_this.listNameOrId)) {
                                var listGuid = new SP.Guid(_this.listNameOrId);
                                _this._list = lists.getById(listGuid);
                            }
                            else {
                                _this._list = lists.getByTitle(_this.listNameOrId);
                            }
                            _this.ClientContext.load(_this._list);
                            _this.ClientContext.executeQueryAsync(function () {
                                _this.List = _this;
                                _this.Target = _this._list;
                                resolve(Promise.resolve(_this));
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
            this.listNameOrId = listNameOrId;
            if (typeof web !== "undefined" && web !== null) {
                this._webPromise = web;
            }
        }
        return NvListSvc;
    }());
    shgeneric.NvListSvc = NvListSvc;

    var NvListUtils = (function () {
        function NvListUtils() {
        }
        NvListUtils.getListFieldsInternalNames = function (listPromise) {
            return new Promise(function (resolve, reject) {
                Promise.resolve(listPromise).then(function (value) {
                    var ctx = value.ClientContext;
                    var lst = value.Target;
                    var fieldsColl = lst.get_fields();
                    ctx.load(fieldsColl);
                    ctx.executeQueryAsync(function (sender, args) {
                        var fields = shgeneric.iterateSpCollection(fieldsColl);
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
            return new Promise(function (resolve, reject) {
                var allItems = new Array();
                Promise.resolve(listPromise).then(function (value) {
                    var ctx = value.ClientContext;
                    var lst = value.Target;
                    var q = shgeneric.cloneSpCamlQuery(query);
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
                                var listItemsArray = shgeneric.iterateSpCollection(listItems);
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
    shgeneric.NvListUtils = NvListUtils;

    var NvViewSvc = (function () {
        function NvViewSvc(viewNameOrId, list) {
            var _this = this;
            this._listPromise = null;
            this.GetAsync = function () {
                return new Promise(function (resolve, reject) {
                    try {
                        if (_this._listPromise == null) {
                            throw new Error('The list promise is null');
                        }
                        Promise.resolve(_this._listPromise).then(function (list) {
                            _this.List = list;
                            _this.Web = _this.List.Web;
                            _this.Site = _this.List.Site;
                            _this.ClientContext = _this.List.ClientContext;
                            var views = _this.List.Target.get_views();
                            if (shgeneric.Helpers.guidRx.test(_this.viewNameOrId)) {
                                var viewGuid = new SP.Guid(_this.viewNameOrId);
                                _this._view = views.getById(viewGuid);
                            }
                            else {
                                _this._view = views.getByTitle(_this.viewNameOrId);
                            }
                            _this.ClientContext.load(_this._view);
                            _this.ClientContext.executeQueryAsync(function () {
                                _this.Target = _this._view;
                                resolve(Promise.resolve(_this));
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
    shgeneric.NvViewSvc = NvViewSvc;
})(shgeneric || (shgeneric = {}));
