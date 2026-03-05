import { describe, it, beforeEach, expect, vi } from "vitest";

const MockExternals = vi.hoisted(() => {
  const mock = vi.fn() as any;
  mock.Promise = Promise;
  return mock;
});
const MockDateConverter = vi.hoisted(() => ({
  dateToNumber: vi.fn().mockReturnValue("NUMBER"),
  numberToDate: vi.fn().mockReturnValue("DATE"),
}));
const MockWorkbook = vi.hoisted(() => {
  const mock = {
    fromBlankAsync: vi.fn().mockReturnValue("WORKBOOK"),
    fromDataAsync: vi.fn().mockReturnValue("WORKBOOK"),
    fromFileAsync: vi.fn().mockReturnValue("WORKBOOK"),
    MIME_TYPE: "MIME_TYPE",
  } as any;
  return mock;
});
const MockFormulaError = vi.hoisted(() => function FormulaError() {});

vi.mock("../../lib/externals", () => ({ default: MockExternals }));
vi.mock("../../lib/dateConverter", () => ({ default: MockDateConverter }));
vi.mock("../../lib/Workbook", () => ({ default: MockWorkbook }));
vi.mock("../../lib/FormulaError", () => ({ default: MockFormulaError }));

import XlsxPopulate from "../../lib/XlsxPopulate";

describe("XlsxPopulate", () => {
  beforeEach(() => {
    MockDateConverter.dateToNumber.mockReturnValue("NUMBER");
    MockDateConverter.numberToDate.mockReturnValue("DATE");
    MockWorkbook.fromBlankAsync.mockReturnValue("WORKBOOK");
    MockWorkbook.fromDataAsync.mockReturnValue("WORKBOOK");
    MockWorkbook.fromFileAsync.mockReturnValue("WORKBOOK");
    MockWorkbook.MIME_TYPE = "MIME_TYPE";
  });

  describe("dateToNumber", () => {
    it("should call dateConverter.dateToNumber", () => {
      expect(XlsxPopulate.dateToNumber("DATE")).toBe("NUMBER");
      expect(MockDateConverter.dateToNumber).toHaveBeenCalledWith("DATE");
    });
  });

  describe("fromBlankAsync", () => {
    it("should call Workbook.fromBlankAsync", () => {
      expect(XlsxPopulate.fromBlankAsync()).toBe("WORKBOOK");
      expect(MockWorkbook.fromBlankAsync).toHaveBeenCalledWith();
    });
  });

  describe("fromDataAsync", () => {
    it("should call Workbook.fromDataAsync", () => {
      expect(XlsxPopulate.fromDataAsync("DATA", "OPTS")).toBe("WORKBOOK");
      expect(MockWorkbook.fromDataAsync).toHaveBeenCalledWith("DATA", "OPTS");
    });
  });

  describe("fromFileAsync", () => {
    it("should call Workbook.fromFileAsync", () => {
      expect(XlsxPopulate.fromFileAsync("PATH", "OPTS")).toBe("WORKBOOK");
      expect(MockWorkbook.fromFileAsync).toHaveBeenCalledWith("PATH", "OPTS");
    });
  });

  describe("numberToDate", () => {
    it("should call dateConverter.numberToDate", () => {
      expect(XlsxPopulate.numberToDate("NUMBER")).toBe("DATE");
      expect(MockDateConverter.numberToDate).toHaveBeenCalledWith("NUMBER");
    });
  });

  describe("statics", () => {
    it("should set the statics", () => {
      expect(XlsxPopulate.MIME_TYPE).toBe("MIME_TYPE");
      expect(XlsxPopulate.FormulaError).toBe(MockFormulaError);
    });
  });
});
