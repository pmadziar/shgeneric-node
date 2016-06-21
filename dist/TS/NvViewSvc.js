"use strict";
var es6_promise_1 = require("es6-promise");
var globals_1 = require("./globals");
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
                        if (globals_1.Helpers.guidRx.test(_this.viewNameOrId)) {
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
