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
const isWindows = process.platform === "win32";

if (!isWindows) {
  describe.skip("e2e-generate", () => {});
} else {
  describe("e2e-generate", () => {
    let interopPath = glob.sync(
      "C:\\Program Files\\Microsoft Office\\root\\Office*\\ADDINS\\**\\Microsoft.Office.Interop.Excel.dll",
    )[0];
    if (!interopPath) {
      interopPath = glob.sync(
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office*\\DCF\\Microsoft.Office.Interop.Excel.dll",
      )[0];
    }
    if (!interopPath)
      throw new Error("Unable to find the Microsoft.Office.Interop.Excel.dll!");

    testCases.map((testCase: string | Function) => {
      it(testCase, async () => {
        const { default: edge } = await import("edge-js");

        const passwordFile = `${testCase}password.txt`;
        const password =
          fs.existsSync(passwordFile) && fs.readFileSync(passwordFile, "utf8");

        let workbook: any;
        if (fs.existsSync(`${testCase}template.xlsx`)) {
          workbook = await (XlsxPopulate as any).fromFileAsync(
            `${testCase}template.xlsx`,
          );
        } else {
          workbook = await (XlsxPopulate as any).fromBlankAsync();
        }

        const { default: generate } = await import(
          pathToFileURL(path.resolve(`${testCase}generate.js`)).href
        );
        generate(workbook);

        await workbook.toFileAsync(`${testCase}out.xlsx`, {
          password: password || undefined,
        });

        const results = await new Promise<any>((resolve, reject) => {
          const wbPath = path.resolve(`${testCase}out.xlsx`);
          const parseSource = fs.readFileSync(`${testCase}parse.cs`);
          const parseTemplate = fs.readFileSync("./template.cs");
          const source = parseTemplate.toString() + parseSource.toString();

          const parse = edge.func({
            source,
            references: ["System.Drawing.dll", interopPath],
          });

          parse(
            { path: wbPath, password: password || null },
            (err: any, results: any) => {
              if (err) return reject(err);
              resolve(results);
            },
          );
        });

        const expected = JSON.parse(
          fs.readFileSync(`${testCase}expected.json`, "utf8"),
        );
        expect(results).toEqual(expected);
      });
    });
  });
}
