import TTE from "../../src/tableToExcel";
import jsdom from "jsdom";
import fs from "fs";

export const defaultOpts = {
  name: "export.xlsx",
  autoStyle: false,
  sheet: {
    name: "Sheet 1"
  }
};

/**
 * Returns an empty worksheet for testing.
 * @param {object} opts
 */
export function getWorkSheet(opts = {}) {
  let _opts = defaultOpts;
  opts = { ..._opts, ...opts };
  let wb = TTE.initWorkBook();
  return TTE.initSheet(wb, opts.sheet.name);
}
/**
 * Read the conents of the a given file and returns the parsed DOM element
 * corresponding to the first table in the document.
 * @param {string} filename name of the html file where the table to test is written
 * @returns {HTMLTableElement}
 */
export function getTable(filename) {
  if (filename) {
    let path = __dirname + "/../samples/" + filename + ".html";
    let htmlString = fs.readFileSync(path, "utf8");
    const { JSDOM } = jsdom;
    const dom = new JSDOM(htmlString);
    return dom.window.document.querySelector("table");
  }
}
