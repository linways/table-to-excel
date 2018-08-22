# Table to Excel

Export HTML table to valid excel file effortlessly.
This is Utility methods for [protobi/js-xlsx](https://github.com/protobi/js-xlsx) to create excel from html table with styling options.

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
To export content of table `#table1` run:

```javascript
XLSX.utils.html.save_table_as_excel(document.getElementById("table1"), {
  name: "test.xlsx"
});
```

See [samples/test1.html](https://github.com/linways/table-to-excel/blob/master/samples/test1.html) or [this fiddle](https://jsfiddle.net/rohithb/e2h4mbc5/)for working example.

# Cell Types

Cell types can be set using `data-t` attribute in each `td` tag.  
Possible values: `b` Boolean, `n` Number, `e` error, `s` String, `d` Date
Example:

```html
<!-- for setting a cell type as number -->
<td data-t="n">2500</td>
<!-- for setting a cell type as date -->
<td data-t="d">05-23-2018</td>
```

# Cell Styling

All styles are set using `data` attributes on `td` tags.
There are 5 types of attributes: `data-f-*`, `data-a-*`, `data-b-*`, `data-fill-*` and `data-num-fmt` which corresponds to five top-level attributes `font`, `alignment`, `border`, `fill` and `numFmt` specified in [js-xlsx](https://github.com/protobi/js-xlsx).

| Category  | Attribute              | Description                            | Values                                                                                      |
| --------- | ---------------------- | -------------------------------------- | ------------------------------------------------------------------------------------------- |
| font      | `data-f-name`          | Font name                              | "Calibri" // default. Eg: "Arial"                                                           |
|           | `data-f-sz`            | Font size                              | "11" // font size in points                                                                 |
|           | `data-f-color`         | Font color                             | A hex ARGB value. Eg: FFFFOOOO for opaque red.                                              |
|           | `data-f-bold`          | Bold                                   | `true` or `false`                                                                           |
|           | `data-f-italic`        | Italic                                 | `true` or `false`                                                                           |
|           | `data-f-strike`        | Strike                                 | `true` or `false`                                                                           |
| Alignment | `data-a-h`             | Horizontal alignment                   | `left` or `center` or `right`                                                               |
|           | `data-a-v`             | Vertical alignment                     | `bottom` or `center` or `top`                                                               |
|           | `data-a-wrap`          | Wrap text                              | `true` or `false`                                                                           |
|           | `data-a-text-rotation` | Text rotation                          | Number from 0 to 180 or 255 (default is 0)                                                  |
|           |                        |                                        | `90` for rotating text 90 degrees                                                           |
|           |                        |                                        | `255` is special, aligned vertically                                                        |
| Border    | `data-b-t-s`           | Border top style                       | Refer `BORDER_STYLES`                                                                       |
|           | `data-b-b-s`           | Border top style                       | Refer `BORDER_STYLES`                                                                       |
|           | `data-b-l-s`           | Border top style                       | Refer `BORDER_STYLES`                                                                       |
|           | `data-b-r-s`           | Border top style                       | Refer `BORDER_STYLES`                                                                       |
| Fill      | `data-fill-color`      | Cell background color                  | A hex ARGB value.                                                                           |
| numFmt    | `data-num-fmt`         | Number Format                          | "0" // integer index to built in formats, see StyleBuilder.SSF property                     |
|           |                        |                                        | "0.00%" // string matching a built-in format, see StyleBuilder.SSF                          |
|           |                        |                                        | "0.0%" // string specifying a custom format                                                 |
|           |                        |                                        | "0.00%;\\(0.00%\\);\\-;@" // string specifying a custom format, escaping special characters |
| Exclude   | `data-exclude`         | Exclude this cell in the exported xlsx | `true`                                                                                      |

**`BORDER_STYLES:`** `thin`, `medium`, `thick`, `dotted`, `hair`, `dashed`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot` `mediumDashDotDot` `slantDashDot`

# Column Width

Column width's can be set by specifying `data-cols-width` in the `<table>` tag.
`data-cols-width` accepts comma separated column widths specified in character count ([refer](https://github.com/SheetJS/js-xlsx#column-properties)).
`data-cols-width="10,20"` will set width of first coulmn as width of 10 charaters and second column as 20 characters wide.
Example:

```html
<table data-cols-width="10,20,30">
...
</table>
```

# TODO

1. Add row height option
2. Option to set border color
3. Remove unnecessary parser methods from js-xlsx to reduce the size of the dist file

# License

Please consult the attached LICENSE file for details.  
Please consult licenses of the original repositories before using for any commercial purposes.

1. https://github.com/protobi/js-xlsx/blob/master/LICENSE
2. https://github.com/eligrey/FileSaver.js/blob/master/LICENSE.md
3. https://github.com/SheetJS/js-xlsx/blob/master/LICENSE
4. https://github.com/Stuk/jszip/blob/master/LICENSE.markdown - JSZip is dual licensed and is used according to the terms of the MIT License.
