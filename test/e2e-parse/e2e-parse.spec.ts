import { describe, it, expect } from "vitest";
import { pathToFileURL } from "url";
import fs from "fs";
import glob from "glob";
import path from "path";
import XlsxPopulate from "../../lib/XlsxPopulate";

// Run tests relative to this directory so relative paths in test cases work
process.chdir(import.meta.dirname!);

// const testCases = ["./encrypted/"]; // To focus
const testCases = glob.sync("./*/");

describe("e2e-parse", () => {
  testCases.map((testCase) => {
    it(testCase, async () => {
      const passwordFile = `${testCase}password.txt`;
      const password =
        fs.existsSync(passwordFile) && fs.readFileSync(passwordFile, "utf8");

      const workbook = await (XlsxPopulate as any).fromFileAsync(
        `${testCase}input.xlsx`,
        { password: password || undefined },
      );

      const { default: parse } = await import(
        pathToFileURL(path.resolve(`${testCase}parse.js`)).href
      );
      const results = await parse(workbook);
      const expected = JSON.parse(
        fs.readFileSync(`${testCase}expected.json`, "utf8"),
      );
      expect(results).toEqual(expected);
    });
  });
});
