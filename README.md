# Table to Excel 2

Export HTML table to valid excel file effortlessly.
This library uses [guyonroche/exceljs](https://github.com/guyonroche/exceljs) under the hood to create the excel.

# Installation

In the browser, just add a script tag:

```html
<script type="text/javascript" src="../dist/tableToExcel2.js"></script>
```

# Usage

Create your HTML table as normal.  
To export content of table `#table1` run:

```javascript
TableToExcel.convert(document.getElementById("table1"));
```

or

```javascript
TableToExcel.convert(document.getElementById("table1"), {
  name: "table1.xlsx",
  sheet: {
    name: "Sheet 1"
  }
});
```

See [samples/index.html]() or [this fiddle](https://jsfiddle.net/rohithb/e2h4mbc5/)for working example.

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
|           | `data-b-b-s`           | Border bottom style                    | Refer `BORDER_STYLES`                                                                       |
|           | `data-b-l-s`           | Border left style                      | Refer `BORDER_STYLES`                                                                       |
|           | `data-b-r-s`           | Border right style                     | Refer `BORDER_STYLES`                                                                       |
| Fill      | `data-fill-color`      | Cell background color                  | A hex ARGB value.                                                                           |
| numFmt    | `data-num-fmt`         | Number Format                          | "0"                                                                                         |
|           |                        |                                        | "0.00%"                                                                                     |
|           |                        |                                        | "0.0%" // string specifying a custom format                                                 |
|           |                        |                                        | "0.00%;\\(0.00%\\);\\-;@" // string specifying a custom format, escaping special characters |
| Exclude   | `data-exclude`         | Exclude this cell in the exported xlsx | `true`                                                                                      |
| Hyperlink | `data-hyperlink`       | To add hyperlink to a cell             |                                                                                             |

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
