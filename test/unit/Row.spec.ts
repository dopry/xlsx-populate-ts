import { describe, it, beforeEach, expect, vi } from 'vitest';
import _ from 'lodash';
import { createSpyObj } from '../helpers/spyObj';

const MockCell = vi.hoisted(() => vi.fn());
vi.mock('../../lib/Cell', () => ({ default: MockCell }));

import Row from '../../lib/Row';

describe("Row", () => {
    let row: any, rowNode: any, sheet: any, style: any, styleSheet: any, workbook: any, horizontalPageBreaks: any;

    beforeEach(() => {
        let i = 1;

        MockCell.mockClear();
        MockCell.mockImplementation(function (this: any) { this.id = i++; });
        MockCell.prototype.columnNumber = vi.fn().mockReturnValue(2);
        MockCell.prototype.toObject = vi.fn().mockImplementation(function (this: any) { return this.id; });
        MockCell.prototype.find = vi.fn();
        MockCell.prototype.style = vi.fn();

        const Style = class {} as any;
        Style.prototype.id = vi.fn().mockReturnValue("STYLE_ID");
        Style.prototype.style = vi.fn().mockImplementation((name: string) => `STYLE:${name}`);
        style = new Style();

        styleSheet = createSpyObj("styleSheet", ["createStyle"]);
        styleSheet.createStyle.mockReturnValue(style);

        workbook = createSpyObj("workbook", ["sharedStrings", "styleSheet"]);
        workbook.styleSheet.mockReturnValue(styleSheet);

        horizontalPageBreaks = createSpyObj("horizontalPageBreaks", ["add"]);

        sheet = createSpyObj('sheet', ['name', 'workbook', 'existingColumnStyleId', 'forEachExistingColumnNumber', 'horizontalPageBreaks']);
        sheet.name.mockReturnValue('NAME');
        sheet.workbook.mockReturnValue(workbook);
        sheet.horizontalPageBreaks.mockReturnValue(horizontalPageBreaks);
        sheet.existingColumnStyleId.mockImplementation((columnNumber: number) => columnNumber === 4 ? "STYLE_ID" : undefined);
        sheet.forEachExistingColumnNumber.mockImplementation((callback: (n: number) => void) => _.forEach([1, 2, 4], callback));

        rowNode = {
            name: 'row',
            attributes: { r: 7 },
            children: [{ name: 'c', attributes: { r: "B7" } }]
        };

        row = new Row(sheet, rowNode);
    });

    /* PUBLIC */

    describe("address", () => {
        it("should return the address", () => {
            expect(row.address()).toBe('7:7');
            expect(row.address({ anchored: true })).toBe('$7:$7');
            expect(row.address({ includeSheetName: true })).toBe("'NAME'!7:7");
            expect(row.address({ includeSheetName: true, anchored: true })).toBe("'NAME'!$7:$7");
        });
    });

    describe("cell", () => {
        beforeEach(() => {
            MockCell.mockClear();
        });

        it("should return an existing cell", () => {
            expect(row.cell(2)).toBeInstanceOf(MockCell);
            expect(MockCell).not.toHaveBeenCalled();
        });

        it("should return an existing cell by column name", () => {
            expect(row.cell('B')).toBeInstanceOf(MockCell);
            expect(MockCell).not.toHaveBeenCalled();
        });

        it("should create a new cell as needed by number", () => {
            const cell = row.cell(5);
            expect(cell).toBeInstanceOf(MockCell);
            expect(MockCell).toHaveBeenCalledWith(row, 5, undefined);
            expect(row._cells[5]).toBe(cell);
        });

        it("should create a new cell as needed by column name", () => {
            const cell = row.cell('C');
            expect(cell).toBeInstanceOf(MockCell);
            expect(MockCell).toHaveBeenCalledWith(row, 3, undefined);
            expect(row._cells[3]).toBe(cell);
        });

        it("should create a new cell with an existing column style id", () => {
            sheet.existingColumnStyleId.mockReturnValue(5);
            expect(row.cell('C')).toBeInstanceOf(MockCell);
            expect(MockCell).toHaveBeenCalledWith(row, 3, 5);
        });

        it("should create a new cell with an existing row style id", () => {
            rowNode.attributes.s = 3;
            expect(row.cell('C')).toBeInstanceOf(MockCell);
            expect(MockCell).toHaveBeenCalledWith(row, 3, 3);
        });

        it("should create a new cell with an existing row and column style id", () => {
            sheet.existingColumnStyleId.mockReturnValue(5);
            rowNode.attributes.s = 3;
            expect(row.cell('C')).toBeInstanceOf(MockCell);
            expect(MockCell).toHaveBeenCalledWith(row, 3, 3);
        });

        it("should throw an exception on an index of 0", () => {
            expect(() => row.cell(0)).toThrowError(RangeError);
        });

        it("should throw an exception on an index of -1", () => {
            expect(() => row.cell(-1)).toThrowError(RangeError);
        });
    });

    describe("height", () => {
        it("should get/set the height", () => {
            expect(row.height()).toBeUndefined();

            row.height(56.7);
            expect(row.height()).toBe(56.7);
            expect(rowNode.attributes).toEqual({ r: 7, customHeight: 1, ht: 56.7 });

            row.height(undefined);
            expect(row.height()).toBeUndefined();
            expect(rowNode.attributes).toEqual({ r: 7 });
        });
    });

    describe("hidden", () => {
        it("should get/set hidden", () => {
            expect(row.hidden()).toBe(false);

            row.hidden(true);
            expect(row.hidden()).toBe(true);
            expect(rowNode.attributes).toEqual({ r: 7, hidden: 1 });

            row.hidden(false);
            expect(row.hidden()).toBe(false);
            expect(rowNode.attributes).toEqual({ r: 7 });
        });
    });

    describe("rowNumber", () => {
        it("should return the row number", () => {
            expect(row.rowNumber()).toBe(7);
        });
    });

    describe("sheet", () => {
        it("should return the sheet", () => {
            expect(row.sheet()).toBe(sheet);
        });
    });

    describe("style", () => {
        beforeEach(() => {
            vi.spyOn(row, "_createStyleIfNeeded");
            row._style = style;
        });

        it("should get a single style", () => {
            expect(row.style("foo")).toBe("STYLE:foo");
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(row._createStyleIfNeeded).toHaveBeenCalledWith();
        });

        it("should get multiple styles", () => {
            expect(row.style(["foo", "bar", "baz"])).toEqual({
                foo: "STYLE:foo", bar: "STYLE:bar", baz: "STYLE:baz"
            });
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(style.style).toHaveBeenCalledWith("bar");
            expect(style.style).toHaveBeenCalledWith("baz");
            expect(row._createStyleIfNeeded).toHaveBeenCalledWith();
        });

        it("should set a single style", () => {
            expect(row._cells[2]).toBeDefined();
            expect(row._cells[4]).toBeUndefined();

            expect(row.style("foo", "value")).toBe(row);
            expect(style.style).toHaveBeenCalledWith("foo", "value");

            expect(row._cells[1]).toBeUndefined();
            expect(row._cells[2].style).toHaveBeenCalledWith("foo", "value");
            expect(row._cells[3]).toBeUndefined();
            expect(row._cells[4].style).toHaveBeenCalledWith("foo", "value");
        });

        it("should assign a style when asked", () => {
            row._style = undefined;
            expect(row._cells[2]).toBeDefined();
            expect(row._cells[4]).toBeUndefined();

            expect(row.style(style)).toBe(row);

            expect(row._cells[1]).toBeUndefined();
            expect(row._cells[2].style).toHaveBeenCalledWith(style);
            expect(row._cells[3]).toBeUndefined();
            expect(row._cells[4].style).toHaveBeenCalledWith(style);
        });

        it("should set multiple styles", () => {
            expect(row.style({ foo: "FOO", bar: "BAR", baz: "BAZ" })).toBe(row);
            expect(style.style).toHaveBeenCalledWith("foo", "FOO");
            expect(style.style).toHaveBeenCalledWith("bar", "BAR");
            expect(style.style).toHaveBeenCalledWith("baz", "BAZ");

            expect(row._cells[2].style).toHaveBeenCalledWith("foo", "FOO");
            expect(row._cells[2].style).toHaveBeenCalledWith("bar", "BAR");
            expect(row._cells[2].style).toHaveBeenCalledWith("baz", "BAZ");
            expect(row._cells[4].style).toHaveBeenCalledWith("foo", "FOO");
            expect(row._cells[4].style).toHaveBeenCalledWith("bar", "BAR");
            expect(row._cells[4].style).toHaveBeenCalledWith("baz", "BAZ");
        });
    });

    describe("workbook", () => {
        it("should return the workbook", () => {
            expect(row.workbook()).toBe(workbook);
        });
    });

    describe('addPageBreak', () => {
        it("should add a rowBreak and return the row", () => {
            expect(row.addPageBreak()).toBe(row);
        });
    });

    /* INTERNAL */

    describe("clearCellsUsingSharedFormula", () => {
        it("should clear cells with matching shared formula", () => {
            row._cells = [
                undefined,
                {
                    sharesFormula: vi.fn().mockReturnValue(true),
                    clear: vi.fn()
                },
                undefined,
                {
                    sharesFormula: vi.fn().mockReturnValue(false),
                    clear: vi.fn()
                }
            ];

            row.clearCellsUsingSharedFormula(7);
            expect(row._cells[1].sharesFormula).toHaveBeenCalledWith(7);
            expect(row._cells[1].clear).toHaveBeenCalledWith();
            expect(row._cells[3].sharesFormula).toHaveBeenCalledWith(7);
            expect(row._cells[3].clear).not.toHaveBeenCalled();
        });
    });

    describe("find", () => {
        it("should return the matches", () => {
            MockCell.prototype.find.mockReturnValue(true);
            expect(row.find('foo')).toEqual([row.cell(2)]);
            expect(MockCell.prototype.find).toHaveBeenCalledWith(/foo/gim, undefined);

            MockCell.prototype.find.mockReturnValue(false);
            expect(row.find('bar', 'baz')).toEqual([]);
            expect(MockCell.prototype.find).toHaveBeenCalledWith(/bar/gim, 'baz');
        });
    });

    describe("hasCell", () => {
        it("should return true/false if the cell exists or not", () => {
            expect(row.hasCell(1)).toBe(false);
            expect(row.hasCell(2)).toBe(true);
            expect(row.hasCell(3)).toBe(false);
        });

        it("should throw an exception on an index of 0", () => {
            expect(() => row.hasCell(0)).toThrowError(RangeError);
        });

        it("should throw an exception on an index of -1", () => {
            expect(() => row.hasCell(-1)).toThrowError(RangeError);
        });
    });

    describe("hasStyle", () => {
        it("should return true/false if the row has a style set", () => {
            expect(row.hasStyle()).toBe(false);
            rowNode.attributes.s = 0;
            expect(row.hasStyle()).toBe(true);
        });
    });

    describe("minUsedColumnNumber", () => {
        it("should return the min column number", () => {
            row._cells = [];
            row._cells[5] = row._cells[7] = {};
            expect(row.minUsedColumnNumber()).toBe(5);
        });
    });

    describe("maxUsedColumnNumber", () => {
        it("should return the max column number", () => {
            row._cells = [];
            row._cells[5] = row._cells[7] = {};
            expect(row.maxUsedColumnNumber()).toBe(7);
        });
    });

    describe("toXml", () => {
        it("should return the node", () => {
            expect(row.toXml()).toBe(rowNode);
        });
    });

    /* PRIVATE */

    describe("_createStyleIfNeeded", () => {
        it("should create a style", () => {
            rowNode.attributes.s = 3;
            row._createStyleIfNeeded();
            expect(row._style).toBe(style);
            expect(rowNode.attributes.s).toBe("STYLE_ID");
            expect(styleSheet.createStyle).toHaveBeenCalledWith(3);
        });

        it("should NOT create a style", () => {
            const existingStyle = {};
            row._style = existingStyle;
            row._createStyleIfNeeded();
            expect(row._style).toBe(existingStyle);
            expect(styleSheet.createStyle).not.toHaveBeenCalled();
        });
    });

    describe("_init", () => {
        it("should store existing rows", () => {
            expect(row._cells).toEqual([undefined, undefined, expect.any(MockCell)]);
            expect(rowNode.children).toBe(row._cells);
            expect(MockCell).toHaveBeenCalledWith(row, { name: 'c', attributes: { r: "B7" } });
        });
    });
});
