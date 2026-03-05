"use strict";

const escapeRegExp = (value: string) => value.replace(/[\\^$.*+?()[\]{}|]/g, "\\$&");

/**
 * Convert a pattern to a RegExp.
 * @param {RegExp|string} pattern - The pattern to convert.
 * @returns {RegExp} The regex.
 * @private
 */
const regexify = (pattern: RegExp | string): RegExp => {
  if (typeof pattern === "string") {
    pattern = new RegExp(escapeRegExp(pattern), "igm");
  }

  (pattern as RegExp).lastIndex = 0;

  return pattern as RegExp;
};

export default regexify;
