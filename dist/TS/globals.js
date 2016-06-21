"use strict";
var _ = require("lodash");
var es6_promise_1 = require("es6-promise");
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
