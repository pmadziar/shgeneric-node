"use strict";
var _ = require("lodash");
var es6_promise_1 = require("es6-promise");
var NvSiteSvc = (function () {
    function NvSiteSvc(siteServerRelativeUrl) {
        var _this = this;
        this.GetAsync = function () {
            return new es6_promise_1.Promise(function (resolve, reject) {
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
                    resolve(es6_promise_1.Promise.resolve(_this));
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
exports.NvSiteSvc = NvSiteSvc;
