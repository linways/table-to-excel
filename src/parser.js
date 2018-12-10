const TTEParser = (function() {
  let methods = {};

  methods.parseDomToTable = function(ws, table) {
    let _r, _c;
    let rows = table.getElementsByTagName("tr");
    for (_r = 0; _r < rows.length; ++_r) {
      let row = rows[_r];
      let cells = row.children;
      let exRow = [];
      for (_c = 0; _c < cells.length; ++_c) {
        let cell = cells[_c];
        exRow[_c] = htmldecode(cell.innerHTML);
      }
      ws.addRow(exRow);
    }
    return ws;
  };

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

  return methods;
})();

export default TTEParser;
