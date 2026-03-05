import { describe, it, beforeEach, expect, vi } from "vitest";
import _ from "lodash";
import { createSpyObj } from "../helpers/spyObj";

// ─── Mock factories (hoisted so vi.mock factories can reference them) ────────

const MockJSZip = vi.hoisted(() => vi.fn() as any);
const MockFs = vi.hoisted(() => ({
  readFile: vi.fn(),
  writeFile: vi.fn(),
}));
const MockExternals = vi.hoisted(() => ({ Promise }) as any);
const MockStyleSheet = vi.hoisted(() => vi.fn() as any);
const MockSheet = vi.hoisted(() => vi.fn() as any);
const MockSharedStrings = vi.hoisted(() => vi.fn() as any);
const MockRelationships = vi.hoisted(() => vi.fn() as any);
const MockContentTypes = vi.hoisted(() => vi.fn() as any);
const MockCoreProperties = vi.hoisted(() => vi.fn() as any);
const MockXmlParser = vi.hoisted(() => vi.fn() as any);
const MockXmlBuilder = vi.hoisted(() => vi.fn() as any);
const MockEncryptor = vi.hoisted(() => vi.fn() as any);
const MockBlank = vi.hoisted(() => vi.fn().mockReturnValue("BLANK") as any);

vi.mock("fs", () => ({ default: MockFs }));
vi.mock("jszip", () => ({ default: MockJSZip }));
vi.mock("../../lib/externals", () => ({ default: MockExternals }));
vi.mock("../../lib/StyleSheet", () => ({ default: MockStyleSheet }));
vi.mock("../../lib/Sheet", () => ({ default: MockSheet }));
vi.mock("../../lib/SharedStrings", () => ({ default: MockSharedStrings }));
vi.mock("../../lib/Relationships", () => ({ default: MockRelationships }));
vi.mock("../../lib/ContentTypes", () => ({ default: MockContentTypes }));
vi.mock("../../lib/CoreProperties", () => ({ default: MockCoreProperties }));
vi.mock("../../lib/XmlParser", () => ({ default: MockXmlParser }));
vi.mock("../../lib/XmlBuilder", () => ({ default: MockXmlBuilder }));
vi.mock("../../lib/Encryptor", () => ({ default: MockEncryptor }));
vi.mock("../../lib/blank", () => ({ default: MockBlank }));

import Workbook from "../../lib/Workbook";

describe("Workbook", () => {
  let workbookNode: any;

  // Delayed-resolve helper: resolves with a small random delay so promises
  // resolve out of order, proving that callers handle ordering correctly.
  const resolved = (val: any) =>
    new Promise<void>((resolve) =>
      setTimeout(resolve, Math.random() * 10),
    ).then(() => val);

  beforeEach(() => {
    // JSZip
    MockJSZip.mockClear();
    MockJSZip.mockImplementation(function (this: any) {});
    MockJSZip.loadAsync = vi.fn().mockResolvedValue(new MockJSZip());
    MockJSZip.prototype.file = vi
      .fn()
      .mockImplementation((fileName: string) => ({
        async: () => Promise.resolve(`TEXT(${fileName})`),
      }));
    MockJSZip.prototype.remove = vi.fn();
    MockJSZip.prototype.generateAsync = vi.fn().mockResolvedValue("ZIP");
    MockJSZip.external = { Promise };

    // fs
    MockFs.readFile.mockReset();
    MockFs.writeFile.mockReset();
    MockFs.readFile.mockImplementation((_path: string, cb: Function) =>
      cb(null, "DATA"),
    );
    MockFs.writeFile.mockImplementation(
      (_path: string, _data: any, cb: Function) => cb(null),
    );

    // StyleSheet
    MockStyleSheet.mockClear();
    MockStyleSheet.prototype.toString = () => "STYLE SHEET";

    // Sheet
    MockSheet.mockClear();
    MockSheet.mockImplementation(function (
      this: any,
      workbook: any,
      sheetIdNode: any,
      sheetNode: any,
      sheetRelationshipsNode: any,
    ) {
      this.workbook = workbook;
      this.sheetIdNode = sheetIdNode;
      this.sheetNode = sheetNode;
      this.sheetRelationshipsNode = sheetRelationshipsNode;
    });
    MockSheet.prototype.find = vi.fn();
    let sheetOutput = false;
    MockSheet.prototype.toXmls = vi.fn().mockImplementation(() => {
      const relationships = sheetOutput ? "RELATIONSHIPS" : undefined;
      sheetOutput = !sheetOutput;
      return {
        sheet: "SHEET",
        id: { attributes: { "r:id": "RID" } },
        relationships,
      };
    });
    MockSheet.prototype.hidden = vi.fn().mockReturnValue(false);
    MockSheet.prototype.tabSelected = vi.fn();

    // SharedStrings
    MockSharedStrings.mockClear();
    MockSharedStrings.prototype.toString = () => "SHARED STRINGS";

    // Relationships
    MockRelationships.mockClear();
    MockRelationships.prototype.toString = () => "RELATIONSHIPS";
    MockRelationships.prototype.findByType = vi.fn();
    MockRelationships.prototype.add = vi.fn();

    // ContentTypes
    MockContentTypes.mockClear();
    MockContentTypes.prototype.toString = () => "CONTENT TYPES";
    MockContentTypes.prototype.findByPartName = vi.fn();
    MockContentTypes.prototype.add = vi.fn();

    // CoreProperties
    MockCoreProperties.mockClear();
    MockCoreProperties.prototype.toString = () => "CORE PROPERTIES";
    MockCoreProperties.prototype.get = vi.fn();
    MockCoreProperties.prototype.set = vi.fn();

    // XmlParser
    MockXmlParser.mockClear();
    MockXmlParser.prototype.parseAsync = vi
      .fn()
      .mockImplementation((text: string) => Promise.resolve(`JSON(${text})`));

    // XmlBuilder
    MockXmlBuilder.mockClear();
    MockXmlBuilder.prototype.build = vi
      .fn()
      .mockImplementation((obj: any) => `XML: ${obj && obj.toString()}`);

    // Encryptor
    MockEncryptor.mockClear();
    MockEncryptor.prototype.encrypt = vi
      .fn()
      .mockImplementation((input: any) => `ENCRYPTED(${input})`);
    MockEncryptor.prototype.decryptAsync = vi
      .fn()
      .mockImplementation((input: any) =>
        Promise.resolve(`DECRYPTED(${input})`),
      );

    // blank
    MockBlank.mockReturnValue("BLANK");

    // workbookNode
    workbookNode = {
      name: "workbook",
      attributes: {},
      children: [
        {
          name: "bookViews",
          attributes: {},
          children: [{ name: "workbookView", attributes: {}, children: [] }],
        },
        {
          name: "sheets",
          attributes: {},
          children: [
            { name: "sheet", attributes: { name: "A", sheetId: 5 } },
            { name: "sheet", attributes: { name: "B", sheetId: 9 } },
          ],
        },
      ],
    };
  });

  describe("static", () => {
    beforeEach(() => {
      vi.spyOn(Workbook.prototype as any, "_initAsync").mockResolvedValue(
        "WORKBOOK",
      );
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });

    describe("fromBlankAsync", () => {
      it("should init with blank data", async () => {
        const wb = await (Workbook as any).fromBlankAsync();
        expect((Workbook.prototype as any)._initAsync).toHaveBeenCalledWith(
          "BLANK",
          undefined,
        );
        expect(wb).toBe("WORKBOOK");
      });
    });

    describe("fromDataAsync", () => {
      it("should init with the data", async () => {
        const wb = await (Workbook as any).fromDataAsync("DATA", "OPTS");
        expect((Workbook.prototype as any)._initAsync).toHaveBeenCalledWith(
          "DATA",
          "OPTS",
        );
        expect(wb).toBe("WORKBOOK");
      });
    });

    describe("fromFileAsync", () => {
      it("should init with the file data", async () => {
        const wb = await (Workbook as any).fromFileAsync("PATH", "OPTS");
        expect((Workbook.prototype as any)._initAsync).toHaveBeenCalledWith(
          "DATA",
          "OPTS",
        );
        expect(MockFs.readFile).toHaveBeenCalledWith(
          "PATH",
          expect.any(Function),
        );
        expect(wb).toBe("WORKBOOK");
      });
    });
  });

  describe("instance", () => {
    let workbook: any;

    beforeEach(() => {
      workbook = new Workbook();
    });

    describe("activeSheet", () => {
      beforeEach(() => {
        workbook._node = workbookNode;
        workbook._sheets = [new MockSheet(), new MockSheet()];
        workbook._activeSheet = workbook._sheets[0];
      });

      it("should return the active sheet", () => {
        expect(workbook.activeSheet()).toBe(workbook._sheets[0]);
        workbook._activeSheet = workbook._sheets[1];
        expect(workbook.activeSheet()).toBe(workbook._sheets[1]);
      });

      it("should set the active sheet", () => {
        expect(workbook.activeSheet(workbook._sheets[1])).toBe(workbook);
        expect(workbook._sheets[0].tabSelected).toHaveBeenCalledWith(false);
        expect(workbook._sheets[1].tabSelected).toHaveBeenCalledWith(true);
        expect(workbook._activeSheet).toBe(workbook._sheets[1]);

        expect(workbook.activeSheet(0)).toBe(workbook);
        expect(workbook._activeSheet).toBe(workbook._sheets[0]);
      });
    });

    describe("addSheet", () => {
      beforeEach(() => {
        workbook._sheets = [new MockSheet()];
        vi.spyOn(workbook, "activeSheet").mockReturnValue(workbook._sheets[0]);
        vi.spyOn(workbook, "sheet").mockReturnValue(undefined);
        workbook._relationships = createSpyObj("relationships", ["add"]);
        workbook._relationships.add.mockReturnValue({
          attributes: { Id: "RID" },
        });
        workbook._maxSheetId = 7;
      });

      it("should throw an error if the sheet name is invalid", () => {
        expect(() => workbook.addSheet()).toThrow();
        expect(() => workbook.addSheet("foo?")).toThrow();
        expect(() =>
          workbook.addSheet("12345678901234567890123456789012"),
        ).toThrow();

        expect(() => workbook.addSheet("foo")).not.toThrow();

        workbook.sheet.mockReturnValue(workbook._sheets[0]);
        expect(() => workbook.addSheet("foo")).toThrow();
      });

      it("should add the sheet at the end", () => {
        const sheet = workbook.addSheet("foo");
        expect(sheet).toBeInstanceOf(MockSheet);
        expect(workbook._sheets.length).toBe(2);
        expect(workbook._sheets[1]).toBe(sheet);
        expect(sheet.workbook).toBe(workbook);
        expect(sheet.sheetIdNode).toEqual({
          name: "sheet",
          attributes: { name: "foo", sheetId: 8, "r:id": "RID" },
          children: [],
        });
        expect(sheet.sheetNode).toBeUndefined();
        expect(sheet.sheetRelationshipsNode).toBeUndefined();
      });

      it("should add the sheet at the given index", () => {
        const sheet1 = workbook.addSheet("foo", 0);
        expect(workbook._sheets.length).toBe(2);
        expect(workbook._sheets[0]).toBe(sheet1);

        const sheet2 = workbook.addSheet("bar", 2);
        expect(workbook._sheets.length).toBe(3);
        expect(workbook._sheets[2]).toBe(sheet2);
      });

      it("should add the sheet before the given sheet", () => {
        const sheet = workbook.addSheet("foo", workbook._sheets[0]);
        expect(workbook._sheets.length).toBe(2);
        expect(workbook._sheets[0]).toBe(sheet);
      });

      it("should add the sheet before the sheet with the given name", () => {
        workbook.sheet.mockImplementation((name: string) => {
          if (name === "existing") return workbook._sheets[0];
        });

        const sheet = workbook.addSheet("foo", "existing");
        expect(workbook._sheets.length).toBe(2);
        expect(workbook._sheets[0]).toBe(sheet);
      });
    });

    describe("definedName", () => {
      it("should return the scoped defined name", () => {
        vi.spyOn(workbook, "scopedDefinedName").mockReturnValue(
          "SCOPED DEFINED NAME",
        );
        expect(workbook.definedName("NAME")).toBe("SCOPED DEFINED NAME");
        expect(workbook.scopedDefinedName).toHaveBeenCalledWith(
          undefined,
          "NAME",
        );
      });
    });

    describe("deleteSheet", () => {
      let sheet1: any, sheet2: any, sheet3: any;
      beforeEach(() => {
        sheet1 = new MockSheet();
        sheet2 = new MockSheet();
        sheet3 = new MockSheet();
        sheet1.name = vi.fn().mockReturnValue("SHEET1");
        sheet2.name = vi.fn().mockReturnValue("SHEET2");
        sheet3.name = vi.fn().mockReturnValue("SHEET3");

        workbook._sheets = [sheet1, sheet2, sheet3];
        workbook._activeSheet = sheet2;
      });

      it("should throw an error if the sheet doesn't exist", () => {
        expect(() => workbook.deleteSheet("foo")).toThrow();
      });

      it("should throw an error if we are trying to hide the only visible sheet", () => {
        sheet1.hidden = vi.fn().mockReturnValue(true);
        sheet2.hidden = vi.fn().mockReturnValue(false);
        sheet3.hidden = vi.fn().mockReturnValue(true);
        expect(() => workbook.deleteSheet(1)).toThrow();
        expect(() => workbook.deleteSheet(0)).not.toThrow();
      });

      it("should delete the sheet and update the active sheet as needed", () => {
        workbook.deleteSheet(1);
        expect(workbook._sheets).toEqual([sheet1, sheet3]);
        expect(workbook._activeSheet).toBe(sheet3);

        workbook.deleteSheet(0);
        expect(workbook._sheets).toEqual([sheet3]);
        expect(workbook._activeSheet).toBe(sheet3);
      });
    });

    describe("find", () => {
      it("should return the matches", () => {
        workbook._sheets = [new MockSheet(), new MockSheet(), new MockSheet()];

        MockSheet.prototype.find.mockReturnValue(["A", "B"]);
        expect(workbook.find("foo")).toEqual(["A", "B", "A", "B", "A", "B"]);
        expect(MockSheet.prototype.find).toHaveBeenCalledWith(
          /foo/gim,
          undefined,
        );

        MockSheet.prototype.find.mockReturnValue("C");
        expect(workbook.find("bar", "baz")).toEqual(["C", "C", "C"]);
        expect(MockSheet.prototype.find).toHaveBeenCalledWith(/bar/gim, "baz");
      });
    });

    describe("moveSheet", () => {
      let sheet1: any, sheet2: any, sheet3: any;
      beforeEach(() => {
        sheet1 = new MockSheet();
        sheet2 = new MockSheet();
        sheet3 = new MockSheet();
        sheet1.name = vi.fn().mockReturnValue("SHEET1");
        sheet2.name = vi.fn().mockReturnValue("SHEET2");
        sheet3.name = vi.fn().mockReturnValue("SHEET3");

        workbook._sheets = [sheet1, sheet2, sheet3];
      });

      it("should throw an error if the sheet doesn't exist", () => {
        expect(() => workbook.moveSheet("foo")).toThrow();
        expect(() => workbook.moveSheet("SHEET1", "foo")).toThrow();
      });

      it("should move the sheet to the end", () => {
        workbook.moveSheet("SHEET2");
        expect(workbook._sheets).toEqual([sheet1, sheet3, sheet2]);
      });

      it("should move the sheet to the given index", () => {
        workbook.moveSheet("SHEET1", 1);
        expect(workbook._sheets).toEqual([sheet2, sheet1, sheet3]);
      });

      it("should move the sheet before the sheet with the given name", () => {
        workbook.moveSheet("SHEET3", "SHEET1");
        expect(workbook._sheets).toEqual([sheet3, sheet1, sheet2]);
      });

      it("should move the sheet before the given sheet", () => {
        workbook.moveSheet("SHEET2", sheet1);
        expect(workbook._sheets).toEqual([sheet2, sheet1, sheet3]);
      });
    });

    describe("outputAsync", () => {
      let relationships: any[];

      beforeEach(() => {
        relationships = [];
        workbook._contentTypes = new MockContentTypes();
        workbook._coreProperties = new MockCoreProperties();
        workbook._relationships = new MockRelationships();
        workbook._sharedStrings = new MockSharedStrings();
        workbook._styleSheet = new MockStyleSheet();
        workbook._node = "WORKBOOK";
        workbook._sheets = [new MockSheet(), new MockSheet()];
        workbook._sheetsNode = { name: "sheets", attributes: {}, children: [] };
        workbook._zip = new MockJSZip();

        workbook._relationships.findById = vi.fn().mockImplementation(() => {
          const relationship: any = { attributes: {} };
          relationships.push(relationship);
          return relationship;
        });
        vi.spyOn(workbook, "_setSheetRefs").mockImplementation(() => {});
        vi.spyOn(workbook, "_convertBufferToOutput").mockReturnValue("OUTPUT");
      });

      it("should output the data", async () => {
        const output = await workbook.outputAsync({ type: "TYPE" });

        expect(output).toBe("OUTPUT");
        expect(workbook._setSheetRefs).toHaveBeenCalledWith();

        expect(workbook._zip.file).toHaveBeenCalledWith(
          "[Content_Types].xml",
          "XML: CONTENT TYPES",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "docProps/core.xml",
          "XML: CORE PROPERTIES",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/_rels/workbook.xml.rels",
          "XML: RELATIONSHIPS",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/sharedStrings.xml",
          "XML: SHARED STRINGS",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/styles.xml",
          "XML: STYLE SHEET",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/workbook.xml",
          "XML: WORKBOOK",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/worksheets/sheet1.xml",
          "XML: SHEET",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).not.toHaveBeenCalledWith(
          "xl/worksheets/_rels/sheet1.xml.rels",
          "XML: RELATIONSHIPS",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/worksheets/sheet2.xml",
          "XML: SHEET",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.file).toHaveBeenCalledWith(
          "xl/worksheets/_rels/sheet2.xml.rels",
          "XML: RELATIONSHIPS",
          { date: new Date(0), createFolders: false },
        );
        expect(workbook._zip.generateAsync).toHaveBeenCalledWith({
          type: "nodebuffer",
          compression: "DEFLATE",
        });
        expect(relationships).toEqual([
          { attributes: { Target: "worksheets/sheet1.xml" } },
          { attributes: { Target: "worksheets/sheet2.xml" } },
        ]);
        expect(workbook._sheetsNode.children).toEqual([
          { attributes: { "r:id": "RID" } },
          { attributes: { "r:id": "RID" } },
        ]);
        expect(MockEncryptor.prototype.encrypt).not.toHaveBeenCalled();
        expect(workbook._convertBufferToOutput).toHaveBeenCalledWith(
          "ZIP",
          "TYPE",
        );
      });

      it("should encrypt the workbook if password is set", async () => {
        const output = await workbook.outputAsync({ password: "PASSWORD" });
        expect(MockEncryptor.prototype.encrypt).toHaveBeenCalledWith(
          "ZIP",
          "PASSWORD",
        );
        expect(workbook._convertBufferToOutput).toHaveBeenCalledWith(
          "ENCRYPTED(ZIP)",
          undefined,
        );
        expect(output).toBe("OUTPUT");
      });
    });

    describe("sheet", () => {
      it("should return the matching sheet", () => {
        workbook._sheets = [{ name: () => "A" }, { name: () => "B" }];

        expect(workbook.sheet(0)).toBe(workbook._sheets[0]);
        expect(workbook.sheet(1)).toBe(workbook._sheets[1]);
        expect(workbook.sheet("A")).toBe(workbook._sheets[0]);
        expect(workbook.sheet("B")).toBe(workbook._sheets[1]);
      });
    });

    describe("sheets", () => {
      it("should return the sheets", () => {
        workbook._sheets = ["SHEET1", "SHEET2"];
        expect(workbook.sheets()).toEqual(["SHEET1", "SHEET2"]);
        expect(workbook.sheets()).not.toBe(workbook._sheets);
      });
    });

    describe("toFileAsync", () => {
      it("should write the workbook to file", async () => {
        vi.spyOn(workbook, "outputAsync").mockResolvedValue("OUTPUT");
        await workbook.toFileAsync("PATH");
        expect(MockFs.writeFile).toHaveBeenCalledWith(
          "PATH",
          "OUTPUT",
          expect.any(Function),
        );
      });
    });

    describe("scopedDefinedName", () => {
      let sheet: any;

      beforeEach(() => {
        sheet = {
          cell: vi.fn().mockReturnValue("CELL"),
          range: vi.fn().mockReturnValue("RANGE"),
          row: vi.fn().mockReturnValue("ROW"),
          column: vi.fn().mockReturnValue("COLUMN"),
        };

        workbook._node = {
          children: [
            {
              name: "definedNames",
              children: [
                {
                  name: "definedName",
                  attributes: { name: "cell" },
                  children: ["Sheet1!$A$1"],
                },
                {
                  name: "definedName",
                  attributes: { name: "range" },
                  children: ["Sheet2!$A$1:B2"],
                },
                {
                  name: "definedName",
                  attributes: { name: "column" },
                  children: ["Sheet3!$A:$A"],
                },
                {
                  name: "definedName",
                  attributes: { name: "row" },
                  children: ["Sheet4!$1:$1"],
                },
                {
                  name: "definedName",
                  localSheet: sheet,
                  attributes: { name: "sheet scope" },
                  children: ["Sheet5!$A$1"],
                },
                {
                  name: "definedName",
                  attributes: { name: "row range" },
                  children: ["Sheet1!$1:$3"],
                },
                {
                  name: "definedName",
                  attributes: { name: "column range" },
                  children: ["Sheet1!$A:$C"],
                },
                {
                  name: "definedName",
                  attributes: { name: "group" },
                  children: ["A1,A2"],
                },
                {
                  name: "definedName",
                  attributes: { name: "formula" },
                  children: ["A1*A2"],
                },
              ],
            },
          ],
        };

        vi.spyOn(workbook, "sheet").mockReturnValue(sheet);
      });

      it("should return undefined if not found", () => {
        expect(
          workbook.scopedDefinedName(undefined, "not found"),
        ).toBeUndefined();
      });

      it("should return the string if not supported", () => {
        expect(workbook.scopedDefinedName(undefined, "row range")).toEqual(
          "Sheet1!$1:$3",
        );
        expect(workbook.scopedDefinedName(undefined, "column range")).toEqual(
          "Sheet1!$A:$C",
        );
        expect(workbook.scopedDefinedName(undefined, "group")).toEqual("A1,A2");
        expect(workbook.scopedDefinedName(undefined, "formula")).toEqual(
          "A1*A2",
        );
      });

      it("should return the selection", () => {
        expect(workbook.scopedDefinedName(undefined, "cell")).toBe("CELL");
        expect(workbook.sheet).toHaveBeenCalledWith("Sheet1");
        expect(sheet.cell).toHaveBeenCalledWith(1, 1);

        expect(workbook.scopedDefinedName(undefined, "range")).toBe("RANGE");
        expect(workbook.sheet).toHaveBeenCalledWith("Sheet2");
        expect(sheet.range).toHaveBeenCalledWith(1, 1, 2, 2);

        expect(workbook.scopedDefinedName(undefined, "column")).toBe("COLUMN");
        expect(workbook.sheet).toHaveBeenCalledWith("Sheet3");
        expect(sheet.column).toHaveBeenCalledWith(1);

        expect(workbook.scopedDefinedName(undefined, "row")).toBe("ROW");
        expect(workbook.sheet).toHaveBeenCalledWith("Sheet3"); // note: uses last called sheet
        expect(sheet.row).toHaveBeenCalledWith(1);
      });

      it("should return the scoped selection", () => {
        expect(
          workbook.scopedDefinedName(undefined, "sheet scope"),
        ).toBeUndefined();
        expect(workbook.scopedDefinedName({}, "sheet scope")).toBeUndefined();
        expect(workbook.scopedDefinedName(sheet, "sheet scope")).toBe("CELL");
      });

      it("should set the defined name with a string", () => {
        expect(workbook.scopedDefinedName(undefined, "NAME", "VALUE")).toBe(
          workbook,
        );
        expect(workbook._node.children[0].children[9]).toEqual({
          name: "definedName",
          attributes: { name: "NAME" },
          children: ["VALUE"],
        });
      });

      it("should define a sheet scoped name", () => {
        expect(workbook.scopedDefinedName(sheet, "NAME", "VALUE")).toBe(
          workbook,
        );
        const node = workbook._node.children[0].children[9];
        expect(node.name).toBe("definedName");
        expect(node.attributes).toEqual({ name: "NAME" });
        expect(node.children).toEqual(["VALUE"]);
        expect(node.localSheet).toBe(sheet);
      });

      it("should set the defined name with a cell", () => {
        const cell = createSpyObj("cell", ["address"]);
        cell.address.mockReturnValue("ADDRESS");

        expect(workbook.scopedDefinedName(undefined, "NAME", cell)).toBe(
          workbook,
        );
        expect(workbook._node.children[0].children[9]).toEqual({
          name: "definedName",
          attributes: { name: "NAME" },
          children: ["ADDRESS"],
        });
        expect(cell.address).toHaveBeenCalledWith({
          includeSheetName: true,
          anchored: true,
        });
      });

      it("should unset a name", () => {
        workbook._node.children[0].children.length = 2;

        workbook.scopedDefinedName(undefined, "cell", null);
        expect(workbook._node.children[0].children.length).toBe(1);

        workbook.scopedDefinedName(undefined, "range", null);
        expect(workbook._node.children.length).toBe(0);
      });
    });

    describe("sharedStrings", () => {
      it("should return the shared strings", () => {
        workbook._sharedStrings = "SHARED STRINGS";
        expect(workbook.sharedStrings()).toBe("SHARED STRINGS");
      });
    });

    describe("styleSheet", () => {
      it("should return the style sheet", () => {
        workbook._styleSheet = "STYLE SHEET";
        expect(workbook.styleSheet()).toBe("STYLE SHEET");
      });
    });

    describe("coreProperties", () => {
      it("should return the core properties", () => {
        workbook._coreProperties = "CORE PROPERTIES";
        expect(workbook.properties()).toBe("CORE PROPERTIES");
      });
    });

    describe("_initAsync", () => {
      beforeEach(() => {
        vi.spyOn(workbook, "_parseNodesAsync").mockImplementation(
          (files: string[]) => {
            return Promise.all(
              _.map(files, (file: string) => {
                if (file === "xl/workbook.xml") return resolved(workbookNode);
                return resolved(`PARSED(${file})`);
              }),
            );
          },
        );
        vi.spyOn(workbook, "_parseSheetRefs");
        vi.spyOn(workbook, "_convertInputToBufferAsync").mockResolvedValue(
          "BUFFER",
        );
      });

      it("should extract the files from the data zip and load the objects", async () => {
        const wb = await workbook._initAsync("DATA", { base64: "BASE64" });
        expect(wb).toBe(workbook);

        expect(workbook._convertInputToBufferAsync).toHaveBeenCalledWith(
          "DATA",
          "BASE64",
        );
        expect(MockEncryptor.prototype.decryptAsync).not.toHaveBeenCalled();
        expect(MockJSZip.loadAsync).toHaveBeenCalledWith("BUFFER");
        expect(workbook._zip).toBeInstanceOf(MockJSZip);

        expect(workbook._contentTypes).toBeInstanceOf(MockContentTypes);
        expect(workbook._relationships).toBeInstanceOf(MockRelationships);
        expect(workbook._sharedStrings).toBeInstanceOf(MockSharedStrings);
        expect(workbook._styleSheet).toBeInstanceOf(MockStyleSheet);
        expect(workbook._sheets[0]).toBeInstanceOf(MockSheet);
        expect(workbook._sheets[1]).toBeInstanceOf(MockSheet);
        expect(workbook._node).toBe(workbookNode);

        expect(workbook._sheets[0].workbook).toBe(workbook);
        expect(workbook._sheets[0].sheetIdNode).toEqual({
          name: "sheet",
          attributes: { name: "A", sheetId: 5 },
        });
        expect(workbook._sheets[0].sheetNode).toEqual(
          "PARSED(xl/worksheets/sheet1.xml)",
        );
        expect(workbook._sheets[0].sheetRelationshipsNode).toEqual(
          "PARSED(xl/worksheets/_rels/sheet1.xml.rels)",
        );
        expect(workbook._sheets[1].workbook).toBe(workbook);
        expect(workbook._sheets[1].sheetIdNode).toEqual({
          name: "sheet",
          attributes: { name: "B", sheetId: 9 },
        });
        expect(workbook._sheets[1].sheetNode).toEqual(
          "PARSED(xl/worksheets/sheet2.xml)",
        );
        expect(workbook._sheets[1].sheetRelationshipsNode).toEqual(
          "PARSED(xl/worksheets/_rels/sheet2.xml.rels)",
        );

        expect(MockContentTypes).toHaveBeenCalledWith(
          "PARSED([Content_Types].xml)",
        );
        expect(MockRelationships).toHaveBeenCalledWith(
          "PARSED(xl/_rels/workbook.xml.rels)",
        );
        expect(MockSharedStrings).toHaveBeenCalledWith(
          "PARSED(xl/sharedStrings.xml)",
        );
        expect(MockStyleSheet).toHaveBeenCalledWith("PARSED(xl/styles.xml)");

        expect(MockRelationships.prototype.findByType).toHaveBeenCalledWith(
          "sharedStrings",
        );
        expect(MockContentTypes.prototype.findByPartName).toHaveBeenCalledWith(
          "/xl/sharedStrings.xml",
        );

        expect(workbook._zip.remove).toHaveBeenCalledWith("xl/calcChain.xml");
        expect(workbook._maxSheetId).toBe(9);
        expect(workbook._parseSheetRefs).toHaveBeenCalledWith();
      });

      it("should decrypt the data if a password is set", async () => {
        const wb = await workbook._initAsync("DATA", { password: "PASSWORD" });
        expect(wb).toBe(workbook);
        expect(workbook._convertInputToBufferAsync).toHaveBeenCalledWith(
          "DATA",
          undefined,
        );
        expect(MockEncryptor.prototype.decryptAsync).toHaveBeenCalledWith(
          "BUFFER",
          "PASSWORD",
        );
        expect(MockJSZip.loadAsync).toHaveBeenCalledWith("DECRYPTED(BUFFER)");
      });

      it("should not add the shared strings if already present (findByType returns value)", async () => {
        MockRelationships.prototype.findByType.mockReturnValue({});
        MockContentTypes.prototype.findByPartName.mockReturnValue({});

        await workbook._initAsync("DATA");
        expect(MockRelationships.prototype.add).not.toHaveBeenCalled();
        expect(MockContentTypes.prototype.add).not.toHaveBeenCalled();
      });

      it("should add the shared strings if not present", async () => {
        MockRelationships.prototype.findByType.mockReturnValue(undefined);
        MockContentTypes.prototype.findByPartName.mockReturnValue(undefined);

        await workbook._initAsync("DATA");
        expect(MockRelationships.prototype.add).toHaveBeenCalledWith(
          "sharedStrings",
          "sharedStrings.xml",
        );
        expect(MockContentTypes.prototype.add).toHaveBeenCalledWith(
          "/xl/sharedStrings.xml",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        );
      });
    });

    describe("_parseNodesAsync", () => {
      it("should parse the nodes", async () => {
        workbook._zip = new MockJSZip();
        const nodes = await workbook._parseNodesAsync(["foo", "bar", "baz"]);

        expect(workbook._zip.file).toHaveBeenCalledWith("foo");
        expect(workbook._zip.file).toHaveBeenCalledWith("bar");
        expect(workbook._zip.file).toHaveBeenCalledWith("baz");

        expect(MockXmlParser.prototype.parseAsync).toHaveBeenCalledWith(
          "TEXT(foo)",
        );
        expect(MockXmlParser.prototype.parseAsync).toHaveBeenCalledWith(
          "TEXT(bar)",
        );
        expect(MockXmlParser.prototype.parseAsync).toHaveBeenCalledWith(
          "TEXT(baz)",
        );

        expect(nodes).toEqual([
          "JSON(TEXT(foo))",
          "JSON(TEXT(bar))",
          "JSON(TEXT(baz))",
        ]);
      });
    });

    describe("_parseSheetRefs", () => {
      beforeEach(() => {
        workbook._node = { name: "workbook", attributes: {}, children: [] };
        workbook._sheets = ["SHEET1", "SHEET2"];
      });

      it("should parse the active sheet", () => {
        workbook._node.children = [
          {
            name: "bookViews",
            attributes: {},
            children: [
              {
                name: "workbookView",
                attributes: { activeTab: 1 },
                children: [],
              },
            ],
          },
        ];

        workbook._parseSheetRefs();
        expect(workbook._activeSheet).toBe("SHEET2");
      });

      it("should parse the defined names sheets", () => {
        workbook._node.children = [
          {
            name: "definedNames",
            attributes: {},
            children: [
              {
                name: "definedName",
                attributes: { name: "WORKBOOK_SCOPE" },
                children: ["VALUE1"],
              },
              {
                name: "definedName",
                attributes: { name: "SHEET_SCOPE", localSheetId: 0 },
                children: ["VALUE2"],
              },
            ],
          },
        ];

        workbook._parseSheetRefs();

        expect(workbook._node.children).toEqual([
          {
            name: "definedNames",
            attributes: {},
            children: [
              {
                name: "definedName",
                attributes: { name: "WORKBOOK_SCOPE" },
                children: ["VALUE1"],
              },
              {
                name: "definedName",
                attributes: { name: "SHEET_SCOPE", localSheetId: 0 },
                children: ["VALUE2"],
                localSheet: "SHEET1",
              },
            ],
          },
        ]);
      });
    });

    describe("_setSheetRefs", () => {
      beforeEach(() => {
        workbook._node = { name: "workbook", attributes: {}, children: [] };
        workbook._sheets = ["SHEET1", "SHEET2"];
        workbook._activeSheet = "SHEET2";
      });

      it("should set the active sheet and create the book view", () => {
        workbook._setSheetRefs();
        expect(workbook._node.children).toEqual([
          {
            name: "bookViews",
            attributes: {},
            children: [
              {
                name: "workbookView",
                attributes: { activeTab: 1 },
                children: [],
              },
            ],
          },
        ]);
      });

      it("should set the active sheet and use the existing book view", () => {
        workbook._node.children = [
          { name: "bookViews", attributes: { foo: true }, children: [] },
        ];

        workbook._setSheetRefs();
        expect(workbook._node.children).toEqual([
          {
            name: "bookViews",
            attributes: { foo: true },
            children: [
              {
                name: "workbookView",
                attributes: { activeTab: 1 },
                children: [],
              },
            ],
          },
        ]);
      });

      it("should set the defined names sheets", () => {
        workbook._node.children = [
          {
            name: "definedNames",
            attributes: {},
            children: [
              {
                name: "definedName",
                attributes: { name: "WORKBOOK_SCOPE" },
                children: ["VALUE1"],
              },
              {
                name: "definedName",
                attributes: { name: "SHEET_SCOPE" },
                children: ["VALUE2"],
                localSheet: "SHEET1",
              },
            ],
          },
        ];

        workbook._setSheetRefs();

        expect(workbook._node.children[1]).toEqual({
          name: "definedNames",
          attributes: {},
          children: [
            {
              name: "definedName",
              attributes: { name: "WORKBOOK_SCOPE" },
              children: ["VALUE1"],
            },
            {
              name: "definedName",
              attributes: { name: "SHEET_SCOPE", localSheetId: 0 },
              children: ["VALUE2"],
              localSheet: "SHEET1",
            },
          ],
        });
      });
    });

    describe("_convertBufferToOutput", () => {
      it("should default to buffer in Node", () => {
        const input = Buffer.alloc(5);
        const output = workbook._convertBufferToOutput(input);
        expect(Buffer.isBuffer(output)).toBe(true);
      });

      it("should return buffers unchanged", () => {
        const input = Buffer.alloc(5);
        const output = workbook._convertBufferToOutput(input, "nodebuffer");
        expect(output).toBe(input);
      });

      it("should convert to a base64 string", () => {
        const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
        const output = workbook._convertBufferToOutput(input, "base64");
        expect(output).toEqual("Zm9vYmFy");
      });

      it("should convert to a binary string", () => {
        const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
        const output = workbook._convertBufferToOutput(input, "binarystring");
        expect(output).toEqual("foobar");
      });

      it("should convert to a Uint8Array", () => {
        const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
        const output = workbook._convertBufferToOutput(input, "uint8array");
        expect(output).toBeInstanceOf(Uint8Array);
        expect(output).toEqual(Uint8Array.from(input));
      });

      it("should convert to an ArrayBuffer", () => {
        const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
        const output = workbook._convertBufferToOutput(input, "arraybuffer");
        expect(output).toBeInstanceOf(ArrayBuffer);
        expect(new Uint8Array(output)).toEqual(Uint8Array.from(input));
      });
    });

    describe("_convertInputToBufferAsync", () => {
      it("should return buffers unchanged", async () => {
        const input = Buffer.alloc(5);
        const output = await workbook._convertInputToBufferAsync(input);
        expect(output).toBe(input);
      });

      it("should convert a base64 string", async () => {
        const output = await workbook._convertInputToBufferAsync(
          "Zm9vYmFy",
          true,
        );
        expect(Buffer.isBuffer(output)).toBe(true);
        expect(output).toEqual(
          Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]),
        );
      });

      it("should convert a binary string", async () => {
        const output = await workbook._convertInputToBufferAsync("foobar");
        expect(Buffer.isBuffer(output)).toBe(true);
        expect(output).toEqual(
          Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]),
        );
      });

      it("should convert a Uint8Array", async () => {
        const input = new Uint8Array([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
        const output = await workbook._convertInputToBufferAsync(input);
        expect(Buffer.isBuffer(output)).toBe(true);
        expect(output).toEqual(
          Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]),
        );
      });

      it("should convert an ArrayBuffer", async () => {
        const input = new Uint8Array([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72])
          .buffer;
        const output = await workbook._convertInputToBufferAsync(input);
        expect(Buffer.isBuffer(output)).toBe(true);
        expect(output).toEqual(
          Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]),
        );
      });
    });

    describe("cloneSheet", () => {
      beforeEach(() => {
        workbook._sheets = [new MockSheet()];
        vi.spyOn(workbook, "activeSheet").mockReturnValue(workbook._sheets[0]);
        vi.spyOn(workbook, "sheet").mockReturnValue(undefined);
        workbook._relationships = createSpyObj("relationships", ["add"]);
        workbook._relationships.add.mockReturnValue({
          attributes: { Id: "RID" },
        });
      });

      it("should throw an error if params are invalid", () => {
        expect(() => workbook.cloneSheet()).toThrow();
        const from = workbook.addSheet("foo");
        expect(() => workbook.cloneSheet(from)).toThrow();
      });

      it("should add the sheet at the end", () => {
        const from = workbook._sheets[0];
        const sheet = workbook.cloneSheet(from, "foo");
        expect(sheet).toBeInstanceOf(MockSheet);
        expect(workbook._sheets.length).toBe(2);
        expect(workbook._sheets[1]).toBe(sheet);
        expect(sheet.workbook).toBe(workbook);
      });

      it("should add the sheet before the given sheet", () => {
        const from = workbook._sheets[0];
        const sheet = workbook.cloneSheet(from, "foo", workbook._sheets[0]);
        expect(workbook._sheets.length).toBe(2);
        expect(workbook._sheets[0]).toBe(sheet);
      });
    });
  });
});
