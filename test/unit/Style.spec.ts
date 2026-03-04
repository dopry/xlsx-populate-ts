import _ from 'lodash';
import Style from '../../lib/Style';
import { createSpyObj } from '../helpers/spyObj';
import { describe, expect, it, beforeEach, vi } from 'vitest';

describe("Style", () => {
    let style: any, styleSheet: any, id: any, xfNode: any, fontNode: any, fillNode: any, borderNode: any, emptyBorderNode: any;

    beforeEach(() => {
        styleSheet = createSpyObj("styleSheet", ['getNumberFormatCode', 'getNumberFormatId']);
        id = "ID";
        xfNode = { name: "xf", attributes: {}, children: [] };
        fontNode = { name: "font", attributes: {}, children: [] };
        fillNode = { name: "fill", attributes: {}, children: [] };
        borderNode = {
            name: "border",
            attributes: {},
            children: [
                { name: "left", attributes: {}, children: [] },
                { name: "right", attributes: {}, children: [] },
                { name: "top", attributes: {}, children: [] },
                { name: "bottom", attributes: {}, children: [] },
                { name: "diagonal", attributes: {}, children: [] }
            ]
        };
        emptyBorderNode = _.cloneDeep(borderNode);
        style = new Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode);
    });

    describe("id", () => {
        it("should return the ID", () => {
            expect(style.id()).toBe("ID");
        });
    });

    describe("style", () => {
        it("should get the style with the given name", () => {
            style._get_foo = vi.fn().mockReturnValue("FOO");
            expect(style.style("foo")).toBe("FOO");
            expect(style._get_foo).toHaveBeenCalledWith();
        });

        it("should set the style with the given name", () => {
            style._set_foo = vi.fn();
            expect(style.style("foo", "FOO")).toBe(style);
            expect(style._set_foo).toHaveBeenCalledWith("FOO");
        });
    });

    describe("bold", () => {
        it("should get/set bold", () => {
            expect(style.style("bold")).toBe(false);
            style.style("bold", true);
            expect(style.style("bold")).toBe(true);
            expect(fontNode.children).toEqual([{ name: "b", attributes: {}, children: [] }]);
            style.style("bold", false);
            expect(style.style("bold")).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("italic", () => {
        it("should get/set italic", () => {
            expect(style.style("italic")).toBe(false);
            style.style("italic", true);
            expect(style.style("italic")).toBe(true);
            expect(fontNode.children).toEqual([{ name: "i", attributes: {}, children: [] }]);
            style.style("italic", false);
            expect(style.style("italic")).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("underline", () => {
        it("should get/set underline", () => {
            expect(style.style("underline")).toBe(false);
            style.style("underline", true);
            expect(style.style("underline")).toBe(true);
            expect(fontNode.children).toEqual([{ name: "u", attributes: {}, children: [] }]);
            style.style("underline", "double");
            expect(style.style("underline")).toBe("double");
            expect(fontNode.children).toEqual([{ name: "u", attributes: { val: "double" }, children: [] }]);
            style.style("underline", true);
            expect(style.style("underline")).toBe(true);
            expect(fontNode.children).toEqual([{ name: "u", attributes: {}, children: [] }]);
            style.style("underline", false);
            expect(style.style("underline")).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("strikethrough", () => {
        it("should get/set strikethrough", () => {
            expect(style.style("strikethrough")).toBe(false);
            style.style("strikethrough", true);
            expect(style.style("strikethrough")).toBe(true);
            expect(fontNode.children).toEqual([{ name: 'strike', attributes: {}, children: [] }]);
            style.style("strikethrough", false);
            expect(style.style("strikethrough")).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("subscript", () => {
        it("should get/set subscript", () => {
            expect(style.style("subscript")).toBe(false);
            style.style("subscript", true);
            expect(style.style("subscript")).toBe(true);
            expect(fontNode.children).toEqual([{ name: "vertAlign", attributes: { val: "subscript" }, children: [] }]);
            style.style("subscript", false);
            expect(style.style("subscript")).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("superscript", () => {
        it("should get/set superscript", () => {
            expect(style.style("superscript")).toBe(false);
            style.style("superscript", true);
            expect(style.style("superscript")).toBe(true);
            expect(fontNode.children).toEqual([{ name: "vertAlign", attributes: { val: "superscript" }, children: [] }]);
            style.style("superscript", false);
            expect(style.style("superscript")).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("fontSize", () => {
        it("should get/set fontSize", () => {
            expect(style.style("fontSize")).toBe(undefined);
            style.style("fontSize", 17);
            expect(style.style("fontSize")).toBe(17);
            expect(fontNode.children).toEqual([{ name: 'sz', attributes: { val: 17 }, children: [] }]);
            style.style("fontSize", undefined);
            expect(style.style("fontSize")).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("fontFamily", () => {
        it("should get/set fontFamily", () => {
            expect(style.style("fontFamily")).toBe(undefined);
            style.style("fontFamily", "Comic Sans MS");
            expect(style.style("fontFamily")).toBe("Comic Sans MS");
            expect(fontNode.children).toEqual([{ name: 'name', attributes: { val: "Comic Sans MS" }, children: [] }]);
            style.style("fontFamily", undefined);
            expect(style.style("fontFamily")).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("fontGenericFamily", () => {
        it("should get/set fontGenericFamily", () => {
            expect(style.style("fontGenericFamily")).toBe(undefined);
            style.style("fontGenericFamily", 1);
            expect(style.style("fontGenericFamily")).toBe(1);
            expect(fontNode.children).toEqual([{ name: 'family', attributes: { val: 1 }, children: [] }]);
            style.style("fontGenericFamily", undefined);
            expect(style.style("fontGenericFamily")).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("fontScheme", () => {
        it("should get/set fontScheme", () => {
            expect(style.style("fontScheme")).toBe(undefined);
            style.style("fontScheme", 'minor');
            expect(style.style("fontScheme")).toBe('minor');
            expect(fontNode.children).toEqual([{ name: 'scheme', attributes: { val: 'minor' }, children: [] }]);
            style.style("fontScheme", undefined);
            expect(style.style("fontScheme")).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe("fontColor", () => {
        it("should get/set fontColor", () => {
            expect(style.style("fontColor")).toBe(undefined);

            style.style("fontColor", "ff0000");
            expect(style.style("fontColor")).toEqual({ rgb: "FF0000" });
            expect(fontNode.children).toEqual([{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }]);

            style.style("fontColor", 5);
            expect(style.style("fontColor")).toEqual({ theme: 5 });
            expect(fontNode.children).toEqual([{ name: 'color', attributes: { theme: 5 }, children: [] }]);

            style.style("fontColor", { theme: 3, tint: -0.2 });
            expect(style.style("fontColor")).toEqual({ theme: 3, tint: -0.2 });
            expect(fontNode.children).toEqual([{ name: 'color', attributes: { theme: 3, tint: -0.2 }, children: [] }]);

            style.style("fontColor", undefined);
            expect(style.style("fontColor")).toBe(undefined);
            expect(fontNode.children).toEqual([]);

            fontNode.children = [{ name: 'color', attributes: { indexed: 7 }, children: [] }];
            expect(style.style("fontColor")).toEqual({ rgb: "00FFFF" });
        });
    });

    describe("horizontalAlignment", () => {
        it("should get/set horizontalAlignment", () => {
            expect(style.style("horizontalAlignment")).toBe(undefined);
            style.style("horizontalAlignment", "center");
            expect(style.style("horizontalAlignment")).toBe("center");
            expect(xfNode.children).toEqual([{ name: "alignment", attributes: { horizontal: "center" }, children: [] }]);
            style.style("horizontalAlignment", undefined);
            expect(style.style("horizontalAlignment")).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("justifyLastLine", () => {
        it("should get/set justifyLastLine", () => {
            expect(style.style("justifyLastLine")).toBe(false);
            style.style("justifyLastLine", true);
            expect(style.style("justifyLastLine")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { justifyLastLine: 1 }, children: [] }]);
            style.style("justifyLastLine", false);
            expect(style.style("justifyLastLine")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("indent", () => {
        it("should get/set indent", () => {
            expect(style.style("indent")).toBe(undefined);
            style.style("indent", 3);
            expect(style.style("indent")).toBe(3);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { indent: 3 }, children: [] }]);
            style.style("indent", undefined);
            expect(style.style("indent")).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("verticalAlignment", () => {
        it("should get/set verticalAlignment", () => {
            expect(style.style("verticalAlignment")).toBe(undefined);
            style.style("verticalAlignment", "center");
            expect(style.style("verticalAlignment")).toBe("center");
            expect(xfNode.children).toEqual([{ name: "alignment", attributes: { vertical: "center" }, children: [] }]);
            style.style("verticalAlignment", undefined);
            expect(style.style("verticalAlignment")).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("wrapText", () => {
        it("should get/set wrapText", () => {
            expect(style.style("wrapText")).toBe(false);
            style.style("wrapText", true);
            expect(style.style("wrapText")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { wrapText: 1 }, children: [] }]);
            style.style("wrapText", false);
            expect(style.style("wrapText")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("shrinkToFit", () => {
        it("should get/set shrinkToFit", () => {
            expect(style.style("shrinkToFit")).toBe(false);
            style.style("shrinkToFit", true);
            expect(style.style("shrinkToFit")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { shrinkToFit: 1 }, children: [] }]);
            style.style("shrinkToFit", false);
            expect(style.style("shrinkToFit")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("textDirection", () => {
        it("should get/set textDirection", () => {
            expect(style.style("textDirection")).toBe(undefined);
            style.style("textDirection", "left-to-right");
            expect(style.style("textDirection")).toBe("left-to-right");
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { readingOrder: 1 }, children: [] }]);
            style.style("textDirection", "right-to-left");
            expect(style.style("textDirection")).toBe("right-to-left");
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { readingOrder: 2 }, children: [] }]);
            style.style("textDirection", undefined);
            expect(style.style("textDirection")).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("textRotation", () => {
        it("should get/set indent", () => {
            expect(style.style("textRotation")).toBe(undefined);
            style.style("textRotation", 15);
            expect(style.style("textRotation")).toBe(15);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 15 }, children: [] }]);
            style.style("textRotation", -25);
            expect(style.style("textRotation")).toBe(-25);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 115 }, children: [] }]);
            style.style("textRotation", undefined);
            expect(style.style("textRotation")).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("angleTextCounterclockwise", () => {
        it("should get/set angleTextCounterclockwise", () => {
            expect(style.style("angleTextCounterclockwise")).toBe(false);
            style.style("angleTextCounterclockwise", true);
            expect(style.style("angleTextCounterclockwise")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 45 }, children: [] }]);
            style.style("angleTextCounterclockwise", false);
            expect(style.style("angleTextCounterclockwise")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("angleTextClockwise", () => {
        it("should get/set angleTextClockwise", () => {
            expect(style.style("angleTextClockwise")).toBe(false);
            style.style("angleTextClockwise", true);
            expect(style.style("angleTextClockwise")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 90 + 45 }, children: [] }]);
            style.style("angleTextClockwise", false);
            expect(style.style("angleTextClockwise")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("rotateTextUp", () => {
        it("should get/set rotateTextUp", () => {
            expect(style.style("rotateTextUp")).toBe(false);
            style.style("rotateTextUp", true);
            expect(style.style("rotateTextUp")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 90 }, children: [] }]);
            style.style("rotateTextUp", false);
            expect(style.style("rotateTextUp")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("rotateTextDown", () => {
        it("should get/set rotateTextDown", () => {
            expect(style.style("rotateTextDown")).toBe(false);
            style.style("rotateTextDown", true);
            expect(style.style("rotateTextDown")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 90 + 90 }, children: [] }]);
            style.style("rotateTextDown", false);
            expect(style.style("rotateTextDown")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("verticalText", () => {
        it("should get/set verticalText", () => {
            expect(style.style("verticalText")).toBe(false);
            style.style("verticalText", true);
            expect(style.style("verticalText")).toBe(true);
            expect(xfNode.children).toEqual([{ name: 'alignment', attributes: { textRotation: 255 }, children: [] }]);
            style.style("verticalText", false);
            expect(style.style("verticalText")).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe("fill", () => {
        it("should get/set solid fill", () => {
            expect(style.style("fill")).toBe(undefined);

            style.style("fill", "ff0000");
            expect(style.style("fill")).toEqual({ type: "solid", color: { rgb: "FF0000" } });
            expect(fillNode.children).toEqual([{
                name: 'patternFill',
                attributes: { patternType: "solid" },
                children: [{ name: 'fgColor', attributes: { rgb: "FF0000" }, children: [] }]
            }]);

            style.style("fill", 5);
            expect(style.style("fill")).toEqual({ type: "solid", color: { theme: 5 } });
            expect(fillNode.children).toEqual([{
                name: 'patternFill',
                attributes: { patternType: "solid" },
                children: [{ name: 'fgColor', attributes: { theme: 5 }, children: [] }]
            }]);

            style.style("fill", { theme: 6, tint: -0.25 });
            expect(style.style("fill")).toEqual({ type: "solid", color: { theme: 6, tint: -0.25 } });

            style.style("fill", { type: "solid", color: { rgb: "ff00ff", tint: 0.7 } });
            expect(style.style("fill")).toEqual({ type: "solid", color: { rgb: "FF00FF", tint: 0.7 } });

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode.children).toEqual([]);
        });

        it("should get/set pattern fill", () => {
            style.style("fill", {
                type: "pattern",
                pattern: "darkVertical",
                foreground: "FF0000",
                background: 7
            });
            expect(style.style("fill")).toEqual({
                type: "pattern",
                pattern: "darkVertical",
                foreground: { rgb: "FF0000" },
                background: { theme: 7 }
            });

            style.style("fill", {
                type: "pattern",
                pattern: "gray0625",
                foreground: { rgb: "aa0000", tint: -1 },
                background: { theme: 3, tint: 1 }
            });
            expect(style.style("fill")).toEqual({
                type: "pattern",
                pattern: "gray0625",
                foreground: { rgb: "AA0000", tint: -1 },
                background: { theme: 3, tint: 1 }
            });

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode.children).toEqual([]);
        });

        it("should get/set gradient fill", () => {
            style.style("fill", {
                type: "gradient",
                angle: 27,
                stops: [
                    { position: 0, color: "ffffff" },
                    { position: 0.5, color: 7 },
                    { position: 1, color: { rgb: "000000", tint: 0.5 } }
                ]
            });
            expect(style.style("fill")).toEqual({
                type: "gradient",
                gradientType: "linear",
                angle: 27,
                stops: [
                    { position: 0, color: { rgb: "FFFFFF" } },
                    { position: 0.5, color: { theme: 7 } },
                    { position: 1, color: { rgb: "000000", tint: 0.5 } }
                ]
            });

            style.style("fill", {
                type: "gradient",
                gradientType: "path",
                top: 0.1,
                bottom: 0.2,
                left: 0.3,
                right: 0.4,
                stops: [
                    { position: 0, color: { theme: 0, tint: -0.3 } },
                    { position: 1, color: "acacac" }
                ]
            });
            expect(style.style("fill")).toEqual({
                type: "gradient",
                gradientType: "path",
                top: 0.1,
                bottom: 0.2,
                left: 0.3,
                right: 0.4,
                stops: [
                    { position: 0, color: { theme: 0, tint: -0.3 } },
                    { position: 1, color: { rgb: "ACACAC" } }
                ]
            });

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode.children).toEqual([]);
        });
    });

    describe("border", () => {
        describe("border", () => {
            it("should get/set border", () => {
                expect(style.style("borderColor")).toEqual({});
                expect(borderNode).toEqual(emptyBorderNode);

                style.style("border", "thin");
                expect(style.style("border")).toEqual({
                    left: { style: "thin" },
                    right: { style: "thin" },
                    top: { style: "thin" },
                    bottom: { style: "thin" }
                });

                style.style("border", undefined);
                expect(style.style("border")).toEqual({});
                expect(borderNode).toEqual(emptyBorderNode);

                style.style("border", { style: "medium", color: { rgb: "acacac" } });
                expect(style.style("border")).toEqual({
                    left: { style: "medium", color: { rgb: "ACACAC" } },
                    right: { style: "medium", color: { rgb: "ACACAC" } },
                    top: { style: "medium", color: { rgb: "ACACAC" } },
                    bottom: { style: "medium", color: { rgb: "ACACAC" } }
                });

                style.style("border", undefined);
                expect(style.style("border")).toEqual({});

                style.style("border", {
                    left: { color: 0 },
                    top: "dashed"
                });
                expect(style.style("border")).toEqual({
                    left: { color: { theme: 0 } },
                    top: { style: "dashed" }
                });
            });
        });

        describe("borderColor", () => {
            it("should get/set borderColor", () => {
                expect(style.style("borderColor")).toEqual({});

                style.style("borderColor", { left: 1, right: "ff0000" });
                expect(style.style("borderColor")).toEqual({
                    left: { theme: 1 },
                    right: { rgb: "FF0000" }
                });

                style.style("borderColor", "ff0000");
                expect(style.style("borderColor")).toEqual({
                    left: { rgb: "FF0000" },
                    right: { rgb: "FF0000" },
                    top: { rgb: "FF0000" },
                    bottom: { rgb: "FF0000" },
                    diagonal: { rgb: "FF0000" }
                });

                style.style("borderColor", 0);
                expect(style.style("borderColor")).toEqual({
                    left: { theme: 0 },
                    right: { theme: 0 },
                    top: { theme: 0 },
                    bottom: { theme: 0 },
                    diagonal: { theme: 0 }
                });

                style.style("borderColor", undefined);
                expect(style.style("borderColor")).toEqual({});
                expect(borderNode).toEqual(emptyBorderNode);
            });
        });

        describe("borderStyle", () => {
            it("should get/set borderStyle", () => {
                expect(style.style("borderStyle")).toEqual({});

                style.style("borderStyle", { left: "thin", right: "thick" });
                expect(style.style("borderStyle")).toEqual({ left: "thin", right: "thick" });

                style.style("borderStyle", "dashed");
                expect(style.style("borderStyle")).toEqual({
                    left: "dashed",
                    right: "dashed",
                    top: "dashed",
                    bottom: "dashed"
                });

                style.style("borderStyle", undefined);
                expect(style.style("borderStyle")).toEqual({});
                expect(borderNode).toEqual(emptyBorderNode);
            });
        });

        describe("diagonalBorderDirection", () => {
            it("should get/set diagonalBorderDirection", () => {
                expect(style.style("diagonalBorderDirection")).toBe(undefined);

                style.style("diagonalBorderDirection", "up");
                expect(style.style("diagonalBorderDirection")).toBe("up");

                style.style("diagonalBorderDirection", "down");
                expect(style.style("diagonalBorderDirection")).toBe("down");

                style.style("diagonalBorderDirection", "both");
                expect(style.style("diagonalBorderDirection")).toBe("both");

                style.style("diagonalBorderDirection", undefined);
                expect(style.style("diagonalBorderDirection")).toBe(undefined);
                expect(borderNode).toEqual(emptyBorderNode);
            });
        });

        describe("sideBorder", () => {
            it("should get/set sideBorder", () => {
                expect(style.style("topBorder")).toBe(undefined);

                style.style("topBorder", "thin");
                expect(style.style("topBorder")).toEqual({ style: "thin" });

                style.style("bottomBorder", { style: "double", color: 6 });
                expect(style.style("bottomBorder")).toEqual({ style: "double", color: { theme: 6 } });

                style.style("topBorder", undefined).style("bottomBorder", undefined);
                expect(style.style("topBorder")).toBe(undefined);
                expect(borderNode).toEqual(emptyBorderNode);
            });
        });

        describe("sideBorderColor", () => {
            it("should get/set sideBorderColor", () => {
                expect(style.style("rightBorderColor")).toBe(undefined);

                style.style("rightBorderColor", "ff0000");
                expect(style.style("rightBorderColor")).toEqual({ rgb: "FF0000" });

                style.style("rightBorderColor", undefined);
                expect(style.style("rightBorderColor")).toBe(undefined);
                expect(borderNode).toEqual(emptyBorderNode);
            });
        });

        describe("sideBorderStyle", () => {
            it("should get/set sideBorderStyle", () => {
                expect(style.style("leftBorderStyle")).toBe(undefined);

                style.style("leftBorderStyle", "thick");
                expect(style.style("leftBorderStyle")).toBe("thick");

                style.style("leftBorderStyle", undefined);
                expect(style.style("leftBorderStyle")).toBe(undefined);
                expect(borderNode).toEqual(emptyBorderNode);
            });
        });
    });

    describe("numberFormat", () => {
        it("should get/set numberFormat", () => {
            styleSheet.getNumberFormatCode.mockReturnValue("foo");
            styleSheet.getNumberFormatId.mockReturnValue(7);

            expect(style.style("numberFormat")).toBe("foo");
            expect(styleSheet.getNumberFormatCode).toHaveBeenCalledWith(0);

            style.style("numberFormat", "bar");
            expect(styleSheet.getNumberFormatId).toHaveBeenCalledWith('bar');
            expect(xfNode).toEqual({ name: "xf", attributes: { numFmtId: 7 }, children: [] });
            expect(style.style("numberFormat")).toBe("foo");
            expect(styleSheet.getNumberFormatCode).toHaveBeenCalledWith(7);
        });
    });
});
