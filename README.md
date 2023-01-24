# Excel Help

Helper + Middleware for exporting excel files based on [sheetjs](https://sheetjs.com)

## Installation

    npm install excel-help

### Include
Include into js app using common js.
```
const excelHelp = require("excel-help")
```

### Usage

Excel Help can be used in two ways.
1. *As an express js middleware* - To quickly export some data comig from an api as Excel File.
2.  *As a standalone helper* - To aid in creating worksheets and workspaces

### Middleware

#### Usage
```
app.use(excelHelp.middleware);
```

The middleware accepts data in following format to quickly export a *json array*)  or an *array of arrays* or *both* and provides a function *res.xlsx* to export it as a file / stream.

```
res.xlsx("export.xlsx",[
		{"
			type": "json", 
			"data": [{
                "foo":"bar",
				"bar":"foo",
				"abc":"def"
			},
			{
				"foo":"xyz",
				"bar":"lak",
				"abc":"mkx"
			}]
		},
		{
			"type": "columns",
			"data": [
                ["yqmxc", "kqyui", "zhasi", "kljhda"],
                ["yqmxl", "kqyuiads", "zhasida", null, new Date()]
			]
		}
], config?config:{})

```

![[output.png]]

It accepts an array of objects of folklowing format:

```
{
	"type" :  //"json" or "columns",
	"data": //the data //Array of objects //Array of arrays,
	"options" // options object (Not Mandatory)
}
```

###### options

* If type is "columns"
	[options](https://docs.sheetjs.com/docs/api/utilities/#array-of-arrays-input)
* If type is "json"
	We can use the in built sheetjs [options](https://docs.sheetjs.com/docs/api/utilities/#array-of-objects-input) along with these options
     * headerMap - headerMap is an object which can be used to transform the column names. By default the key of the data object is considered as the header.
       ```
       {headerMap: {"foo":"FOO","bar":"BAR"}}
	   ``` 

       * excludeHeaders - excludeHeaders is an array of headers that one wants to exclude from the export.

###### config

config may contain the workbook configuration as defined [config](https://docs.sheetjs.com/docs/api/write-options)


### Standalone

#### Usage

```
const ExcelHelp = require("excel-help");
let excelHelpWb = new ExcelHelp().addSheet(
[
		{
			"type": "json", 
			"data": [{
                     "foo":"bar",
                     "bar":"foo",
                     "abc":"def"
				},
				{
                    "foo":"xyz",
                    "bar":"lak",
                    "abc":"mkx"
				}]
		},
		{
			"type": "columns",
			"data": [
                   ["yqmxc", "kqyui", "zhasi", "kljhda"],
                   ["yqmxl", "kqyuiads", "zhasida", null, new Date()]
                ]
		}
],
"sheet1", //Name of the sheet (Not Mandatory)
config?config:{} // config object (NotMandatory)
).addSheet(
[
		{
			"type": "json", 
			"data": [
				{
					"foo":"xyz",
					"bar":"lak",
					"abc":"mkx"
				}]
		}
],
"sheet2", //Name of the sheet (Not Mandatory)
config?config:{} // config object (NotMandatory)
)
.build() 

```

The method *addSheet* can be chained any no of times to add muliple sheets to the workbook.
addSheet contains a config option which might be used to do operations on the worksheet like changing the width or merging columns. 

**Changing Column Widths**

```
{"!cols": [ { wch: 10 } ]}
```

Other examples can be found on [sheetjs](https://sheetjs.com)


The *build* functions returns the workbook object . The user can then proced to modify the workbook as he needs or [write](https://docs.sheetjs.com/docs/api/write-options) the workbook

Example

```
excelHelp.XLSX.writeFile(excelHelpWb, filename, write_opts)
```

