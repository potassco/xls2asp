# xls2asp

Convert excel spreadsheets to ASP facts

## Usage

```bash
python xls2asp.py --xls examples/instance.xlsm --template template_example.txt --output out.lp
```

**The name of a parsed sheet must follow [gringo syntax for constants](https://github.com/potassco/guide/releases/download/v2.2.0/guide.pdf#page=130)**

### Template

Each line of the template describes a sheet. Use `%` to comment out the rest of the line.
If a sheet is not described in the template it will be ignored.
If a sheet is described in the template, but is not found in the xls file, an error will be raised.  
A line of the template should have the format

```txt
sheetName, style, type1, type2, ...
```

#### Styles

4 different styles are available: `row`, `row_indexed`, `matrix_xy`, and `sparse_matrix_xy`.

##### row

Reading a sheet in `row` style output one predicate per row:  `sheetName(val1, val2,..., valn).`
The predicate name is the (transformed) name of the sheet. The predicate has as many arguments as the number of columns.  
An empty row will be ignored.  
An empty cell in a non empty row will raise an error, if no default value is fixed.
**THE FIRST ROW OF THE SHEET IS IGNORED !!!**
The first type corresponds to the type of the first column, the second type to the type of the second column...

##### row_indexed

Reading a sheet in `row_indexed` style outputs one predicate per row as in style `row`,
with an additional argument for the row index:  `sheetName(index,val1, val2,..., valn).`

##### matrix_xy

Reading a sheet in `matrix_xy` style will output one predicate per cell that is not in the first row or column:  `sheetName(x,y,val).`  
The predicate name is the (transformed) name of the sheet. The predicate has 3 arguments:

* `x` corresponding to the value of the first column of the row of the cell.
* `y` corresponding to the value of the first row of the column of the cell.
* `val` corresponding to the value of the cell

An empty cell in a non empty row will raise an error, if no default value is fixed
The first type corresponds to the type of the first column of the sheet.  
The second type corresponds to the type of the first row of the sheet.  
The third type corresponds to the type of the inner matrix.

##### sparse_matrix_xy

Reading a sheet in `sparse_matrix_xy` style outputs one predicate per cell as in style  `matrix_xy`, but empty cells will be ignored instead of triggering errors.

#### Types

8 different types are available:

* `int`
* `constant`
* `time`  converted to minutes
* `date`  converted to the tuple (dd,mm,yy)
* `datetime`  converted to the tuple (dd,mm,yy,hh,mm,ss)
* `string`
* `auto_detect` automatically detect if one of the above
* `skip` allows to skip a column in row style

##### Default value

Adding `=` and a value after the type in the template
The value will be written in the ASP facts exactly as they are in the template.
Any value can be use regardless the type, but the value must be an int, a valid gringo constant or written between quote.
Example:

```bash
tableName, styleName, int = none, int = -1, string = 4, constant= "les carottes sont cuites ☢ ☮"
```

**The following characters can't be used in the default value: `%`, `=`, `,`**

## Tests

Run the tests using `pytest` with command:

```shell
pytest tests/test.py
```