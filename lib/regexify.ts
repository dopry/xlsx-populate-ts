"use strict";

import _ = require("lodash");

/**
 * Convert a pattern to a RegExp.
 * @param {RegExp|string} pattern - The pattern to convert.
 * @returns {RegExp} The regex.
 * @private
 */
const regexify = (pattern: RegExp | string): RegExp => {
    if (typeof pattern === "string") {
        pattern = new RegExp(_.escapeRegExp(pattern), "igm");
    }

    (pattern as RegExp).lastIndex = 0;

    return pattern as RegExp;
};

export = regexify;
