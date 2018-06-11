var html_utils = {};
function make_html_utils_lib(XLSX) {
    function parse_dom_table(table, _opts) {
        var DENSE = null;
        var opts = _opts || {};
        if (DENSE != null) opts.dense = DENSE;
        var ws = opts.dense ? ([]) : ({});
        var rows = table.getElementsByTagName('tr');
        var sheetRows = opts.sheetRows || 10000000;
        var range = {s: {r: 0, c: 0}, e: {r: 0, c: 0}};
        var merges = [], midx = 0;
        var rowInfo = [];
        var _colWidths = table.getAttribute("data-cols-width");
        if (_colWidths)
            var colInfo = _colWidths.split(',').map(function (i) {
                return {'wch': parseInt(i)}
            });
        var _R = 0, R = 0, _C, C, RS, CS;
        for (; _R < rows.length && R < sheetRows; ++_R) {
            var row = rows[_R];
            if (is_dom_element_hidden(row)) {
                if (opts.display) continue;
                rowInfo[R] = {hidden: true};
            }
            var elts = (row.children);
            for (_C = C = 0; _C < elts.length; ++_C) {
                var elt = elts[_C];
                if (opts.display && is_dom_element_hidden(elt)) continue;
                var v = htmldecode(elt.innerHTML);
                for (midx = 0; midx < merges.length; ++midx) {
                    var m = merges[midx];
                    if (m.s.c == C && m.s.r <= R && R <= m.e.r) {
                        C = m.e.c + 1;
                        midx = -1;
                    }
                }
                CS = +elt.getAttribute("colspan") || 1;
                if ((RS = +elt.getAttribute("rowspan")) > 0 || CS > 1) merges.push({
                    s: {r: R, c: C},
                    e: {r: R + (RS || 1) - 1, c: C + CS - 1}
                });
                var s = {
                    font: {},
                    border: {},
                    alignment: {}
                };
                var o = {t: 's', v: v, s: s};
                var _t = elt.getAttribute("data-t"); //cell type
                var _numFmt = elt.getAttribute("data-numFmt"); // number format
                var _fontSz = elt.getAttribute("data-f-sz"); //font size
                var _fontBold = elt.getAttribute("data-f-bold"); // font-bold
                var _fontColor = elt.getAttribute("data-f-color"); // font-color
                var _underline = elt.getAttribute("data-f-underline"); // underline
                var _italic = elt.getAttribute("data-f-italic"); // italic
                var _aHorizontal = elt.getAttribute("data-a-h"); // align horizontal
                var _aVertical = elt.getAttribute("data-a-v"); // align vertical
                var _aWrap = elt.getAttribute("data-a-wrap"); // text wrap
                if (v != null) {
                    if (v.length == 0) o.t = _t || 'z';
                    else if (opts.raw || v.trim().length == 0 || _t == "s") {
                    }
                    else if (_t && _t.length) {
                        o.t = _t;
                    }
                    if (_fontSz) o.s.font.sz = _fontSz;
                    if (_fontBold) o.s.font.bold = _fontBold;
                    if (_fontColor) o.s.font.color = {rgb: _fontColor};
                    if (_underline) o.s.font.underline = _underline;
                    if (_italic) o.s.font.italic = _italic;
                    if (_numFmt) o.s.numFmt = _numFmt;
                    if (_aHorizontal) o.s.alignment.horizontal = _aHorizontal;
                    if (_aVertical) o.s.alignment.vertical = _aVertical;
                    if (_aWrap) o.s.alignment.wrapText = _aWrap;
                }
                if (opts.dense) {
                    if (!ws[R]) ws[R] = [];
                    ws[R][C] = o;
                }
                else ws[encode_cell({c: C, r: R})] = o;
                if (range.e.c < C) range.e.c = C;
                C += CS;
            }
            ++R;
        }
        if (merges.length) ws['!merges'] = merges;
        if (rowInfo.length) ws['!rows'] = rowInfo;
        if (colInfo && colInfo.length) ws['!cols'] = colInfo;
        range.e.r = R - 1;
        ws['!ref'] = encode_range(range);
        if (R >= sheetRows) ws['!fullref'] = encode_range((range.e.r = rows.length - _R + R - 1, range)); // We can count the real number of rows to parse but we don't to improve the performance
        return ws;
    }


    function is_dom_element_hidden(element) {
        var display = '';
        var get_computed_style = get_get_computed_style_function(element);
        if (get_computed_style) display = get_computed_style(element).getPropertyValue('display');
        if (!display) display = element.style.display; // Fallback for cases when getComputedStyle is not available (e.g. an old browser or some Node.js environments) or doesn't work (e.g. if the element is not inserted to a document)
        return display === 'none';
    }

    /* global getComputedStyle */
    function get_get_computed_style_function(element) {
        // The proper getComputedStyle implementation is the one defined in the element window
        if (element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
        // If it is not available, try to get one from the global namespace
        if (typeof getComputedStyle === 'function') return getComputedStyle;
        return null;
    }

    var htmldecode = (function () {
        var entities = [
            ['nbsp', ' '], ['middot', '·'],
            ['quot', '"'], ['apos', "'"], ['gt', '>'], ['lt', '<'], ['amp', '&']
        ].map(function (x) {
            return [new RegExp('&' + x[0] + ';', "g"), x[1]];
        });
        return function htmldecode(str) {
            var o = str.trim().replace(/\s+/g, " ").replace(/<\s*[bB][rR]\s*\/?>/g, "\n").replace(/<[^>]*>/g, "");
            for (var i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
            return o;
        };
    })();

    function encode_cell(cell) {
        return encode_col(cell.c) + encode_row(cell.r);
    }

    function encode_col(col) {
        var s = "";
        for (++col; col; col = Math.floor((col - 1) / 26)) s = String.fromCharCode(((col - 1) % 26) + 65) + s;
        return s;
    }

    function encode_row(row) {
        return "" + (row + 1);
    }

    function encode_range(cs, ce) {
        if (typeof ce === 'undefined' || typeof ce === 'number') {
            return encode_range(cs.s, cs.e);
        }
        if (typeof cs !== 'string') cs = encode_cell((cs));
        if (typeof ce !== 'string') ce = encode_cell((ce));
        return cs == ce ? cs : cs + ":" + ce;
    }


    function sheet_to_workbook(sheet, opts) {
        var n = opts && opts.sheet ? opts.sheet : "Sheet1";
        var sheets = {};
        sheets[n] = sheet;
        return {SheetNames: [n], Sheets: sheets};
    }

    function table_to_book(table, opts) {
        return sheet_to_workbook(parse_dom_table(table, opts), opts);
    }

    function save_table_as_excel(table, opts) {
        if(typeof saveAs == 'undefined')
            throw new Error('FileSave.js is required. https://github.com/eligrey/FileSaver.js/blob/master/src/FileSaver.js');
        var fileName = 'excel-export.xlsx';
        if(opts.name){
            fileName = opts.name;
            delete opts.name;
        }
        if(Object.keys(opts).length === 0 && opts.constructor === Object) opts = {};
        if(!opts.bookSST) opts.bookSST = false;
        if(!opts.bookType) opts.bookType = 'xlsx';
        if(!opts.type) opts.type = 'binary';
        var wb = table_to_book(table);
        var wbout = XLSX.write(wb, opts);

        function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }
        saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), fileName);
    }

    html_utils.table_to_book = table_to_book;
    html_utils.save_table_as_excel = save_table_as_excel;
}
if (typeof XLSX == 'undefined') {
    throw new Error('You need to include xlsx.core.min.js (https://github.com/protobi/js-xlsx/blob/master/dist/xlsx.core.min.js)  before xlsx_html_utils.js');
}
make_html_utils_lib(XLSX);
XLSX.utils.html = html_utils;