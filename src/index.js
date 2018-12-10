import Parser from "./parser";
import saveAs from "file-saver";
import ExcelJS from "../node_modules/exceljs/dist/es5/exceljs.browser";
// const ExcelJS = require("../node_modules/exceljs/dist/es5/exceljs.browser");

const TableToExcel = (function(Parser) {
  let methods = {};

  let initWorkBook = function() {
    let wb = new ExcelJS.Workbook();
    return wb;
  };

  let initSheet = function(wb, sheetName) {
    let ws = wb.addWorksheet(sheetName);
    return ws;
  };

  let save = function(wb, fileName) {
    wb.xlsx.writeBuffer().then(function(buffer) {
      saveAs(
        new Blob([buffer], { type: "application/octet-stream" }),
        fileName
      );
    });
  };

  let tableToSheet = function(wb, table, opts) {
    let ws = initSheet(wb, opts.sheet.name);
    ws = Parser.parseDomToTable(ws, table);
    return wb;
  };

  let tableToBook = function(table, opts) {
    let wb = initWorkBook();
    wb = tableToSheet(wb, table, opts);
    return wb;
  };

  methods.convert = function(table, opts) {
    let defaultOpts = {
      name: "export.xlsx",
      sheet: {
        name: "Sheet 1"
      }
    };
    opts = { ...defaultOpts, ...opts };
    let wb = tableToBook(table, opts);
    save(wb, opts.name);
  };

  return methods;
})(Parser);

window.TableToExcel = TableToExcel;
// let ExcelJS = require("../node_modules/exceljs/dist/es5/exceljs.browser");
// import saveAs from "file-saver";

// var wb = new ExcelJS.Workbook();
// var ws = wb.addWorksheet("blort");

// ws.getCell("A1").value = "Hello, World!";
// ws.getCell("A2").value = 7;

// wb.xlsx.writeBuffer().then(function(buffer) {
//   saveAs(new Blob([buffer], { type: "application/octet-stream" }), "test.xlsx");
// });

// saveAs(
//     new Blob(wb.xlsx.writeBuffer())
// )

// wb.xlsx
//   .writeBuffer()
//   .then(function(buffer) {
//     var wb2 = new ExcelJS.Workbook();
//     return wb2.xlsx.load(buffer).then(function() {
//       var ws2 = wb2.getWorksheet("blort");
//       expect(ws2).toBeTruthy();

//       expect(ws2.getCell("A1").value).toEqual("Hello, World!");
//       expect(ws2.getCell("A2").value).toEqual(7);
//       done();
//     });
//   })
//   .catch(function(error) {
//     throw error;
//   })
//   .catch(unexpectedError(done));
// sheet.columns = [
//   { header: "Id", key: "id", width: 10 },
//   { header: "Name", key: "name", width: 32 },
//   { header: "D.O.B.", key: "dob", width: 10, outlineLevel: 1 }
// ];
// sheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });
// sheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });
