# table-to-excel
Wrapper for [js-xlsx](https://github.com/protobi/js-xlsx) to create "valid" excel file from html table.
# Installation
In the browser, just add a script tag:
```html
<script type="text/javascript" src="../dist/xlsx_html.full.min.js"></script>
```
`xlsx_html.full.min.js` file contains
1. `lib/xlsx.core.min.js` from [protobi/js-xlsx](https://github.com/protobi/js-xlsx) - A fork of [SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx) which support styles. This library does all the heavy lifting.
2. `lib/FileSaver.min.js` from [eligrey/FileSaver.js](https://github.com/eligrey/FileSaver.js/) - Used to save the created excel file.
3. `src/xlsx_html_utils.js` - Utility functions for js-xlsx to create excel sheets from html tables.   
So, if you want, You can include each library separately.
```html
<script  type="text/javascript" src="../lib/xlsx.core.min.js"></script>
<script  type="text/javascript" src="../dist/xlsx_html_utils.min.js"></script>
<script  type="text/javascript" src="../lib/FileSaver.min.js"></script>
```
# Usage
Create your HTML table as normal.  
To export content of table  `#table1` run:
```javascript
XLSX.utils.html.save_table_as_excel(document.getElementById('table1'), {name: 'test.xlsx'})
```
See [samples/test1.html](https://github.com/linways/table-to-excel/blob/master/samples/test1.html) for working example.

# Cell Styling

#License
