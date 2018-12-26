import TTE from "../../src/tableToExcel";
import fs from "fs";
export const defaultOpts = {
  name: "export.xlsx",
  autoStyle: false,
  sheet: {
    name: "Sheet 1"
  }
};
export function getWorkSheet(opts = {}) {
  let _opts = defaultOpts;
  opts = { ..._opts, ...opts };
  let wb = TTE.initWorkBook();
  return TTE.initSheet(wb, opts.sheet.name);
}

export function getTable(filename) {
  if (filename) {
    let path = __dirname + "/../samples/" + filename + ".html";
    let htmlString = fs.readFileSync(path, "utf8");
    let document = new window.DOMParser().parseFromString(
      htmlString,
      "text/xml"
    );
    return document.firstChild;
  }
}
