"use strict";
var es6_promise_1 = require("es6-promise");
var globals_1 = require("./globals");
var NvWebSvc_1 = require("./NvWebSvc");
var NvListSvc = (function () {
    function NvListSvc(listNameOrId, web) {
        var _this = this;
        this._webPromise = null;
        this.GetAsync = function () {
            return new es6_promise_1.Promise(function (resolve, reject) {
                try {
                    if (_this._webPromise == null) {
                        _this._webPromise = (new NvWebSvc_1.NvWebSvc()).GetAsync();
                    }
                    es6_promise_1.Promise.resolve(_this._webPromise).then(function (web) {
                        _this.Web = web;
                        _this.Site = _this.Web.Site;
                        _this.ClientContext = _this.Web.ClientContext;
                        var lists = _this.Web.Target.get_lists();
                        if (globals_1.Helpers.guidRx.test(_this.listNameOrId)) {
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
        this.listNameOrId = listNameOrId;
        if (typeof web !== "undefined" && web !== null) {
            this._webPromise = web;
        }
    }
    return NvListSvc;
}());
exports.NvListSvc = NvListSvc;
