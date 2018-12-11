import { expect } from "chai";
import Parser from "../src/parser";
import { getWorkSheet, getTable } from "./utils/utils";

describe("Parser", function() {
  it("should convert simple html table to worksheet", function() {
    let table = getTable("simpleTable");
    let ws = getWorkSheet();
    ws = Parser.parseDomToTable(ws, table);
    expect(ws).to.not.be.null;
    expect(ws.getCell("A1").value).to.equals("#");
    expect(ws.getCell("B1").value).to.equals("City");
    expect(ws.rowCount).to.equals(4);
  });

  it("should successfully handle colspan", function() {
    let table = getTable("colSpan");
    let ws = getWorkSheet();
    ws = Parser.parseDomToTable(ws, table);
    expect(ws.getCell("B2").value).to.equals(ws.getCell("C2").value);
    expect(ws.getCell("C2").master).to.equals(ws.getCell("B2"));
    expect(ws.getCell("A4").value).to.equals(ws.getCell("C4").value);
    expect(ws.getCell("C4").master).to.equals(ws.getCell("A4"));
  });

  it("should successfully handle both colspan and rowspan", function() {
    let table = getTable("colAndRowSpan");
    let ws = getWorkSheet();
    ws = Parser.parseDomToTable(ws, table);
    expect(ws.getCell("B2").value).to.equals(ws.getCell("C2").value);
    expect(ws.getCell("B2").value).to.equals(ws.getCell("B3").value);
    expect(ws.getCell("B2").value).to.equals(ws.getCell("C3").value);
    expect(ws.getCell("C5").value).to.equals(ws.getCell("C6").value);
  });
});
