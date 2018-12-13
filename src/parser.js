const TTEParser = (function() {
  let methods = {};

  /**
   * Parse HTML table to excel worksheet
   * @param {object} ws The worksheet object
   * @param {HTML entity} table The table to be converted to excel sheet
   */
  methods.parseDomToTable = function(ws, table, opts) {
    let _r, _c, cs, rs;
    let rows = table.getElementsByTagName("tr");
    let merges = [];
    for (_r = 0; _r < rows.length; ++_r) {
      let row = rows[_r];
      let tds = row.children;
      for (_c = 0; _c < tds.length; ++_c) {
        let td = tds[_c];
        // calculate merges
        cs = parseInt(td.getAttribute("colspan")) || 1;
        rs = parseInt(td.getAttribute("rowspan")) || 1;
        if (cs > 1 || rs > 1) {
          merges.push([
            getExcelColumnName(_c + 1) + (_r + 1),
            getExcelColumnName(_c + cs) + (_r + rs)
          ]);
        }
        let exCell = ws.getCell(getColumnAddress(_c + 1, _r + 1));
        exCell.value = htmldecode(td.innerHTML);
        let styles = getStylesFromCss(td);
        // If first row, set width of the columns.
        if (_r == 0)
          ws.columns[_c].width = Math.round(tds[_c].offsetWidth / 7.2); // convert pixel to character width
      }
    }
    // applying merges to the sheet
    merges.forEach(element => {
      ws.mergeCells(element[0] + ":" + element[1]);
    });
    return ws;
  };

  /**
   * Convert HTML special characters back to normal chars
   */
  let htmldecode = (function() {
    let entities = [
      ["nbsp", " "],
      ["middot", "·"],
      ["quot", '"'],
      ["apos", "'"],
      ["gt", ">"],
      ["lt", "<"],
      ["amp", "&"]
    ].map(function(x) {
      return [new RegExp("&" + x[0] + ";", "g"), x[1]];
    });
    return function htmldecode(str) {
      let o = str
        .trim()
        .replace(/\s+/g, " ")
        .replace(/<\s*[bB][rR]\s*\/?>/g, "\n")
        .replace(/<[^>]*>/g, "");
      for (let i = 0; i < entities.length; ++i)
        o = o.replace(entities[i][0], entities[i][1]);
      return o;
    };
  })();

  /**
   * Takes a positive integer and returns the corresponding column name.
   * @param {number} num  The positive integer to convert to a column name.
   * @return {string}  The column name.
   */
  let getExcelColumnName = function(num) {
    for (var ret = "", a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
      ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
    }
    return ret;
  };

  let getColumnAddress = function(col, row) {
    return getExcelColumnName(col) + row;
  };

  let getStylesFromCss = function(td) {};
  return methods;
})();

export default TTEParser;
