import Range from '../../lib/Range';
import { createSpyObj } from '../helpers/spyObj';
import { describe, expect, it, beforeEach, vi } from 'vitest';

describe("Range", () => {
    let range: any, startCell: any, endCell: any, sheet: any, style: any;

    beforeEach(() => {
        const Style = class {} as any;
        if (!Style.name) Style.name = "Style";
        Style.prototype.style = vi.fn().mockImplementation((name: string) => `STYLE:${name}`);
        style = new Style();

        sheet = createSpyObj('sheet', ['name', 'workbook', 'cell', 'merged', 'incrementMaxSharedFormulaId', 'dataValidation', 'autoFilter']);
        sheet.name.mockReturnValue('NAME');
        sheet.cell.mockImplementation((row: number, column: number) => `CELL[${row}, ${column}]`);
        sheet.workbook.mockReturnValue('WORKBOOK');
        sheet.dataValidation.mockReturnValue('DATAVALIDATION');

        startCell = createSpyObj("startCell", ["rowNumber", "columnNumber", "columnName", "sheet", "value"]);
        startCell.columnName.mockReturnValue("B");
        startCell.columnNumber.mockReturnValue(2);
        startCell.rowNumber.mockReturnValue(3);
        startCell.sheet.mockReturnValue(sheet);

        endCell = createSpyObj("endCell", ["rowNumber", "columnNumber", "columnName", "sheet", "value"]);
        endCell.columnName.mockReturnValue("C");
        endCell.columnNumber.mockReturnValue(3);
        endCell.rowNumber.mockReturnValue(5);
        endCell.sheet.mockReturnValue(sheet);

        range = new Range(startCell, endCell);
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(range.address()).toBe('B3:C5');
            expect(range.address({ startRowAnchored: true })).toBe('B$3:C5');
            expect(range.address({ startColumnAnchored: true })).toBe('$B3:C5');
            expect(range.address({ endRowAnchored: true })).toBe('B3:C$5');
            expect(range.address({ endColumnAnchored: true })).toBe('B3:$C5');
            expect(range.address({ includeSheetName: true })).toBe("'NAME'!B3:C5");
            expect(range.address({
                includeSheetName: true,
                startRowAnchored: true,
                startColumnAnchored: true,
                endRowAnchored: true,
                endColumnAnchored: true
            })).toBe("'NAME'!$B$3:$C$5");
            expect(range.address({ anchored: true })).toBe("$B$3:$C$5");
        });
    });

    describe("cell", () => {
        it("should get the cell relative to the top/left corner", () => {
            expect(range.cell(0, 0)).toBe("CELL[3, 2]");
            expect(sheet.cell).toHaveBeenCalledWith(3, 2);

            expect(range.cell(-2, 3)).toBe("CELL[1, 5]");
            expect(sheet.cell).toHaveBeenCalledWith(1, 5);

            expect(range.cell(4, -1)).toBe("CELL[7, 1]");
            expect(sheet.cell).toHaveBeenCalledWith(7, 1);
        });
    });

    describe("cells", () => {
        it("should get the cells", () => {
            expect(range.cells()).toEqual([
                ["CELL[3, 2]", "CELL[3, 3]"],
                ["CELL[4, 2]", "CELL[4, 3]"],
                ["CELL[5, 2]", "CELL[5, 3]"]
            ]);
        });
    });

    describe("autoFilter", () => {
        it("should mark the range as having an automatic filter", () => {
            expect(range.autoFilter()).toBe(range);
        });
    });

    describe("clear", () => {
        it("should clear the cell", () => {
            vi.spyOn(range, "value").mockReturnValue("RETURN");
            expect(range.clear()).toBe("RETURN");
            expect(range.value).toHaveBeenCalledWith(undefined);
        });
    });

    describe("endCell", () => {
        it("should return the end cell", () => {
            expect(range.endCell()).toBe(endCell);
        });
    });

    describe("forEach", () => {
        it("should call the callback for each cell", () => {
            const callback = vi.fn();
            expect(range.forEach(callback)).toBe(range);
            expect(callback.mock.calls[0]).toEqual(["CELL[3, 2]", 0, 0, range]);
            expect(callback.mock.calls[1]).toEqual(["CELL[3, 3]", 0, 1, range]);
            expect(callback.mock.calls[2]).toEqual(["CELL[4, 2]", 1, 0, range]);
            expect(callback.mock.calls[3]).toEqual(["CELL[4, 3]", 1, 1, range]);
            expect(callback.mock.calls[4]).toEqual(["CELL[5, 2]", 2, 0, range]);
            expect(callback.mock.calls[5]).toEqual(["CELL[5, 3]", 2, 1, range]);
        });
    });

    describe("formula", () => {
        it("should return the top-left cell shared ref formula", () => {
            vi.spyOn(range, "startCell").mockReturnValue({
                getSharedRefFormula: vi.fn().mockReturnValue("RETURN")
            });

            expect(range.formula()).toBe("RETURN");
        });

        it("should set the shared formula", () => {
            sheet.incrementMaxSharedFormulaId.mockReturnValue(8);
            const cells: any = [];
            sheet.cell.mockImplementation((rowNumber: number, columnNumber: number) => {
                return cells[`${rowNumber}, ${columnNumber}`] = {
                    setSharedFormula: vi.fn()
                };
            });

            expect(range.formula("FORMULA")).toBe(range);
            expect(cells["3, 2"].setSharedFormula).toHaveBeenCalledWith(8, "FORMULA", "B3:C5");
            expect(cells["3, 3"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["4, 2"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["4, 3"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["5, 2"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["5, 3"].setSharedFormula).toHaveBeenCalledWith(8);
        });
    });

    describe("map", () => {
        it("should call the callback for each cell and return the values", () => {
            const callback = vi.fn().mockImplementation((cell: any, ri: number, ci: number) => `RETURN[${ri}, ${ci}]`);
            expect(range.map(callback)).toEqual([
                ["RETURN[0, 0]", "RETURN[0, 1]"],
                ["RETURN[1, 0]", "RETURN[1, 1]"],
                ["RETURN[2, 0]", "RETURN[2, 1]"]
            ]);
            expect(callback.mock.calls[0]).toEqual(["CELL[3, 2]", 0, 0, range]);
            expect(callback.mock.calls[1]).toEqual(["CELL[3, 3]", 0, 1, range]);
            expect(callback.mock.calls[2]).toEqual(["CELL[4, 2]", 1, 0, range]);
            expect(callback.mock.calls[3]).toEqual(["CELL[4, 3]", 1, 1, range]);
            expect(callback.mock.calls[4]).toEqual(["CELL[5, 2]", 2, 0, range]);
            expect(callback.mock.calls[5]).toEqual(["CELL[5, 3]", 2, 1, range]);
        });
    });

    describe("merged", () => {
        it("should get merged", () => {
            sheet.merged.mockReturnValue("RETURN");
            expect(range.merged()).toBe("RETURN");
            expect(sheet.merged).toHaveBeenCalledWith("B3:C5");
        });

        it("should merge the cells", () => {
            expect(range.merged(true)).toBe(range);
            expect(sheet.merged).toHaveBeenCalledWith("B3:C5", true);
        });

        it("should unmerge the cells", () => {
            expect(range.merged(false)).toBe(range);
            expect(sheet.merged).toHaveBeenCalledWith("B3:C5", false);
        });
    });

    describe("reduce", () => {
        it("should call the callback for each cell and return the aggregate value", () => {
            const callback = vi.fn().mockImplementation((accumulator: string, cell: any, ri: number, ci: number) => `${accumulator} RETURN[${ri}, ${ci}]`);
            expect(range.reduce(callback, "INITIAL")).toBe("INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0] RETURN[1, 1] RETURN[2, 0] RETURN[2, 1]");
            expect(callback.mock.calls[0]).toEqual(["INITIAL", "CELL[3, 2]", 0, 0, range]);
            expect(callback.mock.calls[1]).toEqual(["INITIAL RETURN[0, 0]", "CELL[3, 3]", 0, 1, range]);
            expect(callback.mock.calls[2]).toEqual(["INITIAL RETURN[0, 0] RETURN[0, 1]", "CELL[4, 2]", 1, 0, range]);
            expect(callback.mock.calls[3]).toEqual(["INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0]", "CELL[4, 3]", 1, 1, range]);
            expect(callback.mock.calls[4]).toEqual(["INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0] RETURN[1, 1]", "CELL[5, 2]", 2, 0, range]);
            expect(callback.mock.calls[5]).toEqual(["INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0] RETURN[1, 1] RETURN[2, 0]", "CELL[5, 3]", 2, 1, range]);
        });
    });

    describe("sheet", () => {
        it("should return the sheet", () => {
            expect(range.sheet()).toBe(sheet);
        });
    });

    describe("startCell", () => {
        it("should return the end cell", () => {
            expect(range.endCell()).toBe(endCell);
        });
    });

    describe("style", () => {
        let cell: any;
        beforeEach(() => {
            cell = { style: style.style };
            sheet.cell.mockReturnValue(cell);
        });

        it("should get a single style value", () => {
            expect(range.style("foo")).toEqual([
                ["STYLE:foo", "STYLE:foo"],
                ["STYLE:foo", "STYLE:foo"],
                ["STYLE:foo", "STYLE:foo"]
            ]);
            expect(cell.style).toHaveBeenCalledWith("foo");
        });

        it("should get multiple style values", () => {
            expect(range.style(["foo", "bar"])).toEqual({
                foo: [
                    ["STYLE:foo", "STYLE:foo"],
                    ["STYLE:foo", "STYLE:foo"],
                    ["STYLE:foo", "STYLE:foo"]
                ],
                bar: [
                    ["STYLE:bar", "STYLE:bar"],
                    ["STYLE:bar", "STYLE:bar"],
                    ["STYLE:bar", "STYLE:bar"]
                ]
            });
            expect(cell.style).toHaveBeenCalledWith("foo");
        });

        it("should set a style from the callback", () => {
            let i = 0;
            const callback = vi.fn().mockImplementation(() => i++);
            expect(range.style("foo", callback)).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 0);
            expect(cell.style).toHaveBeenCalledWith("foo", 1);
            expect(cell.style).toHaveBeenCalledWith("foo", 2);
            expect(cell.style).toHaveBeenCalledWith("foo", 3);
            expect(cell.style).toHaveBeenCalledWith("foo", 4);
            expect(cell.style).toHaveBeenCalledWith("foo", 5);
            expect(callback).toHaveBeenCalledWith(cell, 0, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 0, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 1, range);
        });

        it("should set a style from an array", () => {
            expect(range.style("foo", [
                [0, 1],
                [2, 3],
                [4, 5]
            ])).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 0);
            expect(cell.style).toHaveBeenCalledWith("foo", 1);
            expect(cell.style).toHaveBeenCalledWith("foo", 2);
            expect(cell.style).toHaveBeenCalledWith("foo", 3);
            expect(cell.style).toHaveBeenCalledWith("foo", 4);
            expect(cell.style).toHaveBeenCalledWith("foo", 5);
        });

        it("should set a style from a single value", () => {
            expect(range.style("foo", "bar")).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
        });

        it("should assign a style when asked", () => {
            expect(range.style(style)).toBe(range);
            expect(range._style).toBe(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
        });

        it("should set styles from an object", () => {
            let i = 0;
            expect(range.style({
                foo: "FOO",
                bar: [["BAR0", "BAR1"], ["BAR2", "BAR3"], ["BAR4", "BAR5"]],
                baz: () => `BAZ${i++}`
            })).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR0');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR1');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR2');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR3');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR4');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR5');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ0');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ1');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ2');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ3');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ4');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ5');
        });
    });

    describe('dataValidation', () => {
        it('should return the range', () => {
            expect(range.dataValidation('testing, testing2')).toBe(range);
            expect(sheet.dataValidation).toHaveBeenCalledWith('B3:C5', 'testing, testing2');
        });

        it('should return the range', () => {
            expect(range.dataValidation({ type: 'list',
                allowBlank: false,
                showInputMessage: false,
                prompt: '',
                promptTitle: '',
                showErrorMessage: false,
                error: '',
                errorTitle: '',
                operator: '',
                formula1: 'test1, test2, test3',
                formula2: ''
            })).toBe(range);

            expect(sheet.dataValidation).toHaveBeenCalledWith('B3:C5', { type: 'list',
                allowBlank: false,
                showInputMessage: false,
                prompt: '',
                promptTitle: '',
                showErrorMessage: false,
                error: '',
                errorTitle: '',
                operator: '',
                formula1: 'test1, test2, test3',
                formula2: ''
            });
        });

        it("should get the dataValidation from the range", () => {
            expect(range.dataValidation()).toBe("DATAVALIDATION");
            expect(sheet.dataValidation).toHaveBeenCalledWith("B3:C5");
        });
    });

    describe("tap", () => {
        it("should call the callback and return the range", () => {
            const callback = vi.fn().mockReturnValue("RETURN");
            expect(range.tap(callback)).toBe(range);
            expect(callback).toHaveBeenCalledWith(range);
        });
    });

    describe("thru", () => {
        it("should call the callback and return the callback return value", () => {
            const callback = vi.fn().mockReturnValue("RETURN");
            expect(range.thru(callback)).toBe("RETURN");
            expect(callback).toHaveBeenCalledWith(range);
        });
    });

    describe("values", () => {
        let cell: any;
        beforeEach(() => {
            cell = { value: vi.fn().mockReturnValue("VALUE") };
            sheet.cell.mockReturnValue(cell);
        });

        it("should get the value", () => {
            expect(range.value()).toEqual([
                ["VALUE", "VALUE"],
                ["VALUE", "VALUE"],
                ["VALUE", "VALUE"]
            ]);
            expect(cell.value).toHaveBeenCalledWith();
        });

        it("should set the value from the callback", () => {
            let i = 0;
            const callback = vi.fn().mockImplementation(() => i++);
            expect(range.value(callback)).toBe(range);
            expect(cell.value).toHaveBeenCalledWith(0);
            expect(cell.value).toHaveBeenCalledWith(1);
            expect(cell.value).toHaveBeenCalledWith(2);
            expect(cell.value).toHaveBeenCalledWith(3);
            expect(cell.value).toHaveBeenCalledWith(4);
            expect(cell.value).toHaveBeenCalledWith(5);
            expect(callback).toHaveBeenCalledWith(cell, 0, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 0, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 1, range);
        });

        it("should set values from an array", () => {
            expect(range.value([
                [0, 1],
                [2, 3],
                [4, 5]
            ])).toBe(range);
            expect(cell.value).toHaveBeenCalledWith(0);
            expect(cell.value).toHaveBeenCalledWith(1);
            expect(cell.value).toHaveBeenCalledWith(2);
            expect(cell.value).toHaveBeenCalledWith(3);
            expect(cell.value).toHaveBeenCalledWith(4);
            expect(cell.value).toHaveBeenCalledWith(5);
        });

        it("should set a single value", () => {
            expect(range.value("foo")).toBe(range);
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
        });
    });

    describe("workbook", () => {
        it("should return the workbook", () => {
            expect(range.workbook()).toBe("WORKBOOK");
        });
    });

    describe("_findRangeExtent", () => {
        it("should set the min/max row/column", () => {
            range._startCell = startCell;
            range._endCell = endCell;
            expect(range._minRowNumber).toBe(3);
            expect(range._maxRowNumber).toBe(5);
            expect(range._minColumnNumber).toBe(2);
            expect(range._maxColumnNumber).toBe(3);
            expect(range._numRows).toBe(3);
            expect(range._numColumns).toBe(2);

            range._startCell = endCell;
            range._endCell = startCell;
            expect(range._minRowNumber).toBe(3);
            expect(range._maxRowNumber).toBe(5);
            expect(range._minColumnNumber).toBe(2);
            expect(range._maxColumnNumber).toBe(3);
            expect(range._numRows).toBe(3);
            expect(range._numColumns).toBe(2);
        });
    });
});
