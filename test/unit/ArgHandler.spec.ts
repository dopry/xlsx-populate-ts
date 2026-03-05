import ArgHandler from "../../lib/ArgHandler";
import { describe, expect, it, beforeEach, vi } from "vitest";

describe("ArgHandler", () => {
  let argHandler: any, handlers: any, Style: any;

  beforeEach(() => {
    Style = class {};
    if (!Style.name) Style.name = "Style";

    handlers = {
      empty: vi.fn().mockReturnValue("empty"),
      nil: vi.fn().mockReturnValue("nil"),
      string: vi.fn().mockReturnValue("string"),
      boolean: vi.fn().mockReturnValue("boolean"),
      number: vi.fn().mockReturnValue("number"),
      integer: vi.fn().mockReturnValue("integer"),
      function: vi.fn().mockReturnValue("function"),
      array: vi.fn().mockReturnValue("array"),
      date: vi.fn().mockReturnValue("date"),
      object: vi.fn().mockReturnValue("object"),
      Style: vi.fn().mockReturnValue("Style"),
      "*": vi.fn().mockReturnValue("*"),
    };

    argHandler = new ArgHandler("METHOD")
      .case(handlers.empty)
      .case("nil", handlers.nil)
      .case("string", handlers.string)
      .case("boolean", handlers.boolean)
      .case("number", handlers.number)
      .case(["nil", "integer"], handlers.integer)
      .case("function", handlers.function)
      .case("array", handlers.array)
      .case("date", handlers.date)
      .case("object", handlers.object)
      .case("Style", handlers.Style)
      .case(["nil", "nil", "*"], handlers["*"]);
  });

  describe("handle", () => {
    it("should handle empty", () => {
      expect(argHandler.handle([])).toBe("empty");
      expect(handlers.empty).toHaveBeenCalledWith();
    });

    it("should handle nil", () => {
      expect(argHandler.handle([undefined])).toBe("nil");
      expect(handlers.nil).toHaveBeenCalledWith(undefined);
    });

    it("should handle string", () => {
      expect(argHandler.handle(["foo"])).toBe("string");
      expect(handlers.string).toHaveBeenCalledWith("foo");

      expect(argHandler.handle([""])).toBe("string");
      expect(handlers.string).toHaveBeenCalledWith("");
    });

    it("should handle boolean", () => {
      expect(argHandler.handle([true])).toBe("boolean");
      expect(handlers.boolean).toHaveBeenCalledWith(true);

      expect(argHandler.handle([false])).toBe("boolean");
      expect(handlers.boolean).toHaveBeenCalledWith(false);
    });

    it("should handle number", () => {
      expect(argHandler.handle([0])).toBe("number");
      expect(handlers.number).toHaveBeenCalledWith(0);

      expect(argHandler.handle([-5])).toBe("number");
      expect(handlers.number).toHaveBeenCalledWith(-5);

      expect(argHandler.handle([1.23])).toBe("number");
      expect(handlers.number).toHaveBeenCalledWith(1.23);
    });

    it("should handle integer", () => {
      expect(() => argHandler.handle([undefined, 1.5])).toThrow();

      expect(argHandler.handle([undefined, 3])).toBe("integer");
      expect(handlers.integer).toHaveBeenCalledWith(undefined, 3);

      expect(argHandler.handle([undefined, 0])).toBe("integer");
      expect(handlers.integer).toHaveBeenCalledWith(undefined, 0);

      expect(argHandler.handle([undefined, -5])).toBe("integer");
      expect(handlers.integer).toHaveBeenCalledWith(undefined, -5);
    });

    it("should handle function", () => {
      const func = () => {};
      expect(argHandler.handle([func])).toBe("function");
      expect(handlers.function).toHaveBeenCalledWith(func);
    });

    it("should handle array", () => {
      expect(argHandler.handle([[1, 2, 3]])).toBe("array");
      expect(handlers.array).toHaveBeenCalledWith([1, 2, 3]);
    });

    it("should handle date", () => {
      const date = new Date();
      expect(argHandler.handle([date])).toBe("date");
      expect(handlers.date).toHaveBeenCalledWith(date);
    });

    it("should handle object", () => {
      expect(argHandler.handle([{}])).toBe("object");
      expect(handlers.object).toHaveBeenCalledWith({});
    });

    it("should handle Styles", () => {
      const style = new Style();
      expect(argHandler.handle([style])).toBe("Style");
      expect(handlers.Style).toHaveBeenCalledWith(style);
    });

    it("should handle *", () => {
      expect(argHandler.handle([undefined, undefined, 1])).toBe("*");
      expect(handlers["*"]).toHaveBeenCalledWith(undefined, undefined, 1);
    });
  });
});
