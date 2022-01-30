from odf.opendocument import OpenDocumentText, OpenDocumentSpreadsheet
from odf.style import Style, TableProperties, TableColumnProperties, TableRowProperties, TableCellProperties
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.text import P


table_style         = Style(
    name            = "Таблица_2116_1",
    displayname     = "Таблица№1",
    family          = "table"
)
table_style.addElement(TableProperties())

table_style_column = Style(
    name            = "Таблица_2116_1.A",
    displayname     = "Таблица№1.A",
    family          = "table-column"
).addElement(TableColumnProperties(columnwidth="8"))

table_style_row     = Style(
    name            = "Таблица_2116_1.1",
    displayname     = "Таблица№1.1",
    family          = "table-row"
).addElement(TableRowProperties(minrowheight="4.207cm"))

table_style_cell    = Style(
    name            = "Таблица_2116_1.A1",
    displayname     = "Таблица№1.A1",
    family          = "table-cell"
).addElement(TableCellProperties(padding="0.049cm", borderleft="1.5pt solid #000000", borderright="none",
					bordertop="1.5pt solid #000000", borderbottom="1.5pt solid #000000"))


file_name           = "output_file_3.odt"
# spreadsheet         = OpenDocumentSpreadsheet()
doc                 = OpenDocumentText()
table               = Table(name="Таблица №1")

# spreadsheet.spreadsheet.addElement(table)
# spreadsheet.styles.addElement(table_style)
def add_cell(row, data):
    cell = TableCell(valuetype="string")
    cell.addElement(P(text=data))
    row.addElement(cell)


def add_rows(table, data):
    row = TableRow()
    table.addElement(TableColumn(numbercolumnsrepeated="1"))
    for i in data:
        add_cell(row, i)


    # cell = TableCell(valuetype="string")
    # cell.addElement(P(text="Учебный год"))
    # row.addElement(cell)

    table.addElement(row)


persons = [['Иванов', 'Петров', 'Журба'], ['Иванов2', 'Петров2', 'Журба2'], ['Иванов3', 'Петров3', 'Журба3']]

for i in persons:
    add_rows(table, i)
#first row
# row = TableRow()
#
# cell = TableCell(valuetype="string")
# cell.addElement(P(text="Ученики"))
# row.addElement(cell)
#
# cell = TableCell(valuetype="string")
# cell.addElement(P(text="Учебный год"))
# row.addElement(cell)
# table.addElement(TableColumn(numbercolumnsrepeated="1"))
# table.addElement(row)
#
# tr = TableRow()
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Раз"))
# tr.addElement(tc)
#
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Два"))
# tr.addElement(tc)
#
# tr = TableRow()
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Ячейка1"))
# tr.addElement(tc)
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Ячейка2"))
# tr.addElement(tc)
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Ячейка3"))
# tr.addElement(tc)
# table.addElement(tr)
#
# #second row
# tr = TableRow()
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Ячейка4"))
# tr.addElement(tc)
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Ячейка5"))
# tr.addElement(tc)
# tc = TableCell(valuetype="string")
# tc.addElement(P(text="Ячейка6"))
# tr.addElement(tc)
# table.addElement(tr)
# table.addElement(TableColumn(numbercolumnsrepeated="12"))

if __name__ == "__main__":
    doc.text.addElement(table)
    doc.save("new_file_333.odt")

    import zipfile
    z = zipfile.ZipFile(file=file_name)
    z.printdir()
