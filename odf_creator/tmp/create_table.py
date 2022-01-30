from odf.opendocument import OpenDocumentText, OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell, TableColumn
from odf.text import P
import json


def x_func(val: list, col_count=0, row_count=0, max_row_count=0, rows=[], last_elements_in_col=['']):
    if len(val) > 0:
        row_count += 1
        if len(rows) < row_count:
            rows.append([])
        n = len(rows)
        a = len(rows[n-1])
        if row_count > max_row_count:
            max_row_count = row_count

        for i in range(len(val)):
            el = dict(val[i])
            title = list(el.keys())[0]
            el_val = el.get(title)
            cell = TableCell(valuetype="string", numbercolumnsspanned=len(el_val))
            cell.addElement(P(text=title))
            rows[row_count-1].append(cell)
            last_elements_in_col[col_count]=[cell, row_count-1, len(rows[row_count-1])-1]
            vals = x_func(val=el_val, col_count=col_count, row_count=row_count,
                          max_row_count=max_row_count, rows=rows, last_elements_in_col=last_elements_in_col)
            col_count = vals["col_count"]
            max_row_count = vals["max_row_count"]
            rows = vals["rows"]
            last_elements_in_col = vals["last_elements_in_col"]
    else:
        col_count += 1
        last_elements_in_col.append([])

    return {"max_row_count": max_row_count, "col_count": col_count,
            "rows": rows, "last_elements_in_col": last_elements_in_col}

file_name = '../data/json_table.json'
with open(file_name, "r") as file:
    data = dict(json.load(file))

key = list(data.keys())[0]
result = x_func(val=data.get(key))

rows                = result["rows"]
col_count           = result["col_count"]
last_el_col         = result["last_elements_in_col"]
max_row_count       = result["max_row_count"]
table = Table()
table.addElement(TableColumn(numbercolumnsrepeated=col_count))
n = len(last_el_col)-1
for i in range(n):
    el      = last_el_col[i]
    cell    = el[0]
    row     = el[1]
    col     = el[2]
    print(i + 1, cell, row)
    rows[row][col].setAttribute("numberrowsspanned", max_row_count - row)

for row in rows:
    x = TableRow()
    for cell in row:
        x.addElement(cell)
    table.addElement(x)

file_name           = "output_file_4.odt"
spreadsheet         = OpenDocumentSpreadsheet()
spreadsheet.spreadsheet.addElement(table)
doc                 = OpenDocumentText()
doc.text.addElement(table)
doc.save(file_name)