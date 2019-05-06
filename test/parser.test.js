import { expect } from "chai";
import Parser from "../src/parser";
import { getWorkSheet, getTable, defaultOpts as _opts } from "./utils/utils";

describe("Parser", function() {
  describe("Basic conversion", function() {
    it("should convert simple html table to worksheet", function() {
      let table = getTable("simpleTable");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      expect(ws).to.not.be.null;
      expect(ws.getCell("A1").value).to.equals("#");
      expect(ws.getCell("B1").value).to.equals("City");
      expect(ws.rowCount).to.equals(4);
    });

    it("should successfully handle colspan", function() {
      let table = getTable("colSpan");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      expect(ws.getCell("B2").value).to.equals(ws.getCell("C2").value);
      expect(ws.getCell("C2").master).to.equals(ws.getCell("B2"));
      expect(ws.getCell("A4").value).to.equals(ws.getCell("C4").value);
      expect(ws.getCell("C4").master).to.equals(ws.getCell("A4"));
    });

    it("should successfully handle both colspan and rowspan", function() {
      let table = getTable("colAndRowSpan");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      expect(ws.getCell("B2").value).to.equals(ws.getCell("C2").value);
      expect(ws.getCell("B2").value).to.equals(ws.getCell("B3").value);
      expect(ws.getCell("B2").value).to.equals(ws.getCell("C3").value);
      expect(ws.getCell("C5").value).to.equals(ws.getCell("C6").value);
    });
  });

  describe("Styles", function() {
    var ws;
    beforeEach(function() {
      let table = getTable("styles");
      ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
    });
    describe("Font attributes", function() {
      it("should handle font name properly", function() {
        expect(ws.getCell("A1").font.name).to.be.undefined;
        expect(ws.getCell("A2").font.name).to.equals("Arial");
      });

      it("should handle font size properly", function() {
        expect(parseInt(ws.getCell("A1").font.size)).to.equals(25);
        expect(ws.getCell("A2").font.size).to.be.undefined;
      });

      it("should handle font color properly", function() {
        expect(ws.getCell("A1").font.color).to.deep.equals({
          argb: "FFFFAA00"
        });
        expect(ws.getCell("A2").font.color).to.be.undefined;
      });

      it("should handle bold cells properly", function() {
        expect(ws.getCell("A7").font.bold).to.equals(true);
        expect(ws.getCell("A2").font.bold).to.be.undefined;
      });

      it("should handle italics cells properly", function() {
        expect(ws.getCell("A2").font.italic).to.equals(true);
        expect(ws.getCell("A1").font.italic).to.be.undefined;
      });
    });

    describe("Alignment Attributes", function() {
      it("should handle horizontal alignment", function() {
        expect(ws.getCell("A2").alignment.horizontal).to.equals("center");
      });
      it("should handle vertical alignment", function() {
        expect(ws.getCell("A1").alignment.vertical).to.equals("middle");
        expect(ws.getCell("A2").alignment.vertical).to.equals("top");
      });
      it("should handle wrap text properly", function() {
        expect(ws.getCell("C3").alignment.wrapText).to.equals(true);
        expect(ws.getCell("C4").alignment.wrapText).to.be.undefined;
      });
      it("should handle text rotation properly", function() {
        expect(ws.getCell("A3").alignment.textRotation).to.equals("90");
        expect(ws.getCell("B3").alignment.textRotation).to.equals("vertical");
        expect(ws.getCell("D3").alignment.textRotation).to.equals("-45");
        expect(ws.getCell("E3").alignment.textRotation).to.equals("-90");
        expect(ws.getCell("C4").alignment.wrapText).to.be.undefined;
      });
      it("should handle indent  properly", function() {
        expect(ws.getCell("C5").alignment.indent).to.equals("3");
        expect(ws.getCell("C4").alignment.wrapText).to.be.undefined;
      });
      it("should handle text direction properly", function() {
        expect(ws.getCell("A8").alignment.readingOrder).to.equals("rtl");
      });
    });

    describe("Border Attributes", function() {
      it("handle all border set by b-a-s", function() {
        expect(ws.getCell("A9").border.top.style).to.equals("dashed");
        expect(ws.getCell("A9").border.left.style).to.equals("dashed");
        expect(ws.getCell("A9").border.bottom.style).to.equals("dashed");
        expect(ws.getCell("A9").border.right.style).to.equals("dashed");
      });
      it("handle border set independently", function() {
        expect(ws.getCell("A11").border.top.style).to.equals("thick");
        expect(ws.getCell("A11").border.left.style).to.equals("thick");
        expect(ws.getCell("A11").border.bottom.style).to.equals("thick");
        expect(ws.getCell("A11").border.right.style).to.equals("thick");
      });
      it("handle border color set by b-a-c", function() {
        let color = { argb: "FFFF0000" };
        expect(ws.getCell("A9").border.top.color).to.deep.equals(color);
        expect(ws.getCell("A9").border.left.color).to.deep.equals(color);
        expect(ws.getCell("A9").border.bottom.color).to.deep.equals(color);
        expect(ws.getCell("A9").border.right.color).to.deep.equals(color);
      });
      it("handle border color set independently", function() {
        let color = { argb: "FF00FF00" };
        expect(ws.getCell("A11").border.top.color).to.deep.equals(color);
        expect(ws.getCell("A11").border.left.color).to.deep.equals(color);
        expect(ws.getCell("A11").border.bottom.color).to.deep.equals(color);
        expect(ws.getCell("A11").border.right.color).to.deep.equals(color);
      });
    });

    describe("Fill Attributes", function() {
      it("should handle fill color properly", function() {
        expect(ws.getCell("B5").fill.pattern).to.equals("solid");
        expect(ws.getCell("B5").fill.fgColor).to.deep.equals({
          argb: "FFFF0000"
        });
      });
    });
  });

  describe("Data Type", function() {
    var ws;
    beforeEach(function() {
      let table = getTable("styles");
      ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
    });
    it("should handle number", function() {
      expect(ws.getCell("A4").value).to.be.a("number");
      expect(ws.getCell("B4").value).to.be.a("string");
    });
    it("should handle date", function() {
      let expectedDate = new Date("05-20-2018");
      let actualDate = ws.getCell("D4").value;
      expect(actualDate).to.be.a("date");
      expect(actualDate.getDate()).to.equals(expectedDate.getDate());
      expect(actualDate.getMonth()).to.equals(expectedDate.getMonth());
      expect(actualDate.getFullYear()).to.equals(expectedDate.getFullYear());
    });
    it("should handle boolean", function() {
      expect(ws.getCell("A10").value).to.be.a("boolean");
      expect(ws.getCell("A10").value).to.equals(true);
      expect(ws.getCell("B10").value).to.be.a("boolean");
      expect(ws.getCell("B10").value).to.equals(false);
      expect(ws.getCell("C10").value).to.be.a("boolean");
      expect(ws.getCell("C10").value).to.equals(true);
      expect(ws.getCell("D10").value).to.be.a("boolean");
      expect(ws.getCell("D10").value).to.equals(false);
    });
    it("should handle hyperlink", function() {
      expect(ws.getCell("A7").value.text).to.exist;
      expect(ws.getCell("A7").value.hyperlink).to.exist;
    });
    it("should handle error", function() {
      expect(ws.getCell("E10").value.error).to.exist;
    });
  });

  describe("Exclude row and cell", function() {
    it("should handle exclude row", function() {
      let table = getTable("styles");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      let actualsRows = [...table.getElementsByTagName("tr")];
      expect(ws.rowCount).to.equals(actualsRows.length - 1);
    });
    it("should handle exclude cell", function() {
      let table = getTable("styles");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      let actualsRows = [...table.getElementsByTagName("tr")];
      let actualCells = [...actualsRows[12].children];
      expect(ws.getRow(12).cellCount).to.equals(actualCells.length - 1);
    });
  });
  describe("Column widths", function() {
    it("Should handle the col widths", function() {
      let table = getTable("styles");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      expect(ws.columns[0].width).to.equals(70);
    });
  });

  describe("Row height", function() {
    it("Should handle the row height", function() {
      let table = getTable("styles");
      let ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      expect(ws.getRow(13).height).to.equals(45);
    });
  });

  describe("Multiple merges", function() {
    var ws;
    it("should handle multiple merges properly", function() {
      let table = getTable("multipleMerges");
      ws = getWorkSheet();
      ws = Parser.parseDomToTable(ws, table, _opts);
      expect(ws.getCell("A1").value).to.equals(ws.getCell("F1").value);
      expect(ws.getCell("A2").value).to.equals(ws.getCell("B2").value);
      expect(ws.getCell("C2").value).to.equals(ws.getCell("D2").value);
      expect(ws.getCell("E2").value).to.equals(ws.getCell("F2").value);
      expect(ws.getCell("A3").value).to.equals(ws.getCell("A4").value);
      expect(ws.getCell("B3").value).to.equals("B");
      expect(ws.getCell("B4").value).to.equals("Y");
      expect(ws.getCell("A5").value).to.equals(ws.getCell("A6").value);
      expect(ws.getCell("B5").value).to.equals("B");
      expect(ws.getCell("B6").value).to.equals("Y");
    });
  });
});
