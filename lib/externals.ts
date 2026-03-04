"use strict";

import JSZip = require("jszip");

/**
 * External modules.
 * @private
 */
const externals = {
    /**
     * The Promise library.
     * @type {Promise}
     */
    get Promise() {
        return JSZip.external.Promise;
    },

    set Promise(Promise) {
        JSZip.external.Promise = Promise;
    }
};
export = externals;
