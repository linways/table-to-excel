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
  });
});
