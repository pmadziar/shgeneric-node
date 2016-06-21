"use strict";
var _ = require("lodash");
var es6_promise_1 = require("es6-promise");
var globals_1 = require("./globals");
var NvSiteSvc_1 = require("./NvSiteSvc");
var NvWebSvc = (function () {
    function NvWebSvc(webUrlOrId, site) {
        var _this = this;
        this.webUrlOrId = null;
        this._web = null;
        this._sitePromise = null;
        this.GetAsync = function () {
            return new es6_promise_1.Promise(function (resolve, reject) {
                try {
                    if (_this._sitePromise == null) {
                        _this._sitePromise = (new NvSiteSvc_1.NvSiteSvc(null)).GetAsync();
                    }
                    es6_promise_1.Promise.resolve(_this._sitePromise).then(function (site) {
                        _this.Site = site;
                        _this.ClientContext = _this.Site.ClientContext;
                        if (_.isEmpty(_this.webUrlOrId)) {
                            _this._web = _this.ClientContext.get_web();
                        }
                        else {
                            if (globals_1.Helpers.guidRx.test(_this.webUrlOrId)) {
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
        if (_.isEmpty(webUrlOrId)) {
            this.webUrlOrId = webUrlOrId;
        }
        if (typeof site !== "undefined" && site !== null) {
            this._sitePromise = site;
        }
    }
    return NvWebSvc;
}());
exports.NvWebSvc = NvWebSvc;
