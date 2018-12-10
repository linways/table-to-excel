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
});
