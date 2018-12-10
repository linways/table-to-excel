import TTE from "../../src/index";
import fs from "fs";

export function getWorkSheet(opts = {}) {
  let defaultOpts = {
    name: "export.xlsx",
    sheet: {
      name: "Sheet 1"
    }
  };
  opts = { ...defaultOpts, ...opts };
  let wb = TTE.initWorkBook();
  return TTE.initSheet(wb, opts.sheet.name);
}

export function getTable(filename) {
  if (filename) {
    let path = __dirname + "/../samples/" + filename + ".html";
    let htmlString = fs.readFileSync(path, "utf8");
    return new window.DOMParser().parseFromString(htmlString, "text/xml");
  }
}
