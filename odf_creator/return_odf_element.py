from odf.opendocument import OpenDocument, OpenDocumentChart
from odf.style import Style, TableCellProperties, TableRowProperties, TableProperties, TableColumnProperties
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.text import P, List, ListItem, Span


# from core.reports.odf_gen.modules.marmalades import create_chart

class Table_Cell_Style_Attr:
    def __init__(self, border_style):
        self.all_border_cell_style = {
            "verticalalign": "middle",
            "border": border_style
        }
        self.left_top_cs_attr = {"verticalalign": "middle", "bordertop": border_style, "borderleft": border_style}
        self.top_cs_attr = {"verticalalign": "middle", "bordertop": border_style}
        self.right_top_cs_attr = {"verticalalign": "middle", "bordertop": border_style, "borderright": border_style}
        self.left_cs_attr = {"verticalalign": "middle", "borderleft": border_style}
        self.right_cs_attr = {"verticalalign": "middle", "borderright": border_style}
        self.left_bottom_cs_attr = {"verticalalign": "middle", "borderleft": border_style, "borderbottom": border_style}
        self.bottom_cs_attr = {"verticalalign": "middle", "borderbottom": border_style}
        self.right_bottom_cs_attr = {"verticalalign": "middle", "borderright": border_style, "borderbottom": border_style}


def update_odf_element(element, data: dict) -> OpenDocument:
    data_type = data["type"]
    if data_type == "Table":
        return create_odf_table(odf_element=element, data=data)
    elif data_type == "List":
        return create_odf_list(data_list=data["data"])
    elif data_type == "String":
        return create_str_el(odf_element=element, text=data["data"])
    # elif data_type == "Chart":
    #     return create_odf_chart(chart_data=data)


# def create_odf_chart(chart_data) -> dict:
#     doc = OpenDocumentChart()
#     mychart = create_chart(chart_data)
#     mychart(doc)
#     res = {
#         "element": doc,
#         "style": None
#     }
#     return res

def return_table_size(data: list) -> dict:
    cols = 0
    rows = 0
    for row in data:
        rows += 1
        c = 0
        for cell in row:
            if cell.__contains__("colspan"):
                c += cell["colspan"]
            else:
                c += 1
            if c > cols:
                cols = c
    return {"cols": cols, "rows": rows}

def replace_odf_element(new_el, old_el):
    parent = old_el.parentNode
    parent.insertBefore(newChild=new_el, refChild=old_el)
    parent.removeChild(oldChild=old_el)


def create_str_el(odf_element, text) -> None:
    text = Span(text=text)

    if odf_element.tagName == "text:user-field-get":
        replace_odf_element(new_el=text, old_el=odf_element)
        # odf_element.setAttribute('stringvalue', text)


def create_odf_table(data, odf_element) -> dict:
    styles = []

    thin_border = "0.05pt solid #000000"
    bolt_border = "0.5pt solid #000000"

    cs_attr = Table_Cell_Style_Attr(border_style=thin_border)

    all_cells_style = Style(name="all_border_cell_style", family="table-cell")
    all_cells_content_style = TableCellProperties(attributes=cs_attr.all_border_cell_style)
    all_cells_style.addElement(all_cells_content_style)
    styles.append(all_cells_style)

    row_style = Style(name="row_style", family="table-row")
    row_style.addElement(TableRowProperties(keeptogether="always"))

    table_style = Style(name="table_style", family="table")
    table_style.addElement(TableProperties(bordermodel="collapsing"))
    styles.append(table_style)

    cell_style_name = all_cells_style
    content_cell_style_name = all_cells_content_style

    title = data["table_title"]
    data = data["data"]
    params = return_table_size(data=data)
    cols = params["cols"]
    rows = params["rows"]

    if odf_element.tagName == "text:user-field-get":
        data = title + data
        t = Table(stylename=table_style)
        t.addElement(TableColumn(numbercolumnsrepeated=cols))
        pn1 = odf_element.parentNode
        pn1.removeChild(odf_element)
        pn2 = pn1.parentNode
        pn2.insertBefore(t, pn1)
        pn2.removeChild(pn1)
        odf_element = t

    elif odf_element.tagName == "table:table":
        last_title_table_row = odf_element.childNodes[-1]
        row_style = last_title_table_row.getAttribute("stylename")
        last_title_table_cells_styles = []
        for cell in last_title_table_row.childNodes:
            cell_style_name = cell.getAttribute("stylename")
            content_cell_style_name = cell.childNodes[0].getAttribute("stylename")
            last_title_table_cells_styles.append([cell_style_name, content_cell_style_name, cell.childNodes[0]])
        del odf_element.childNodes[-1]

    else:
        return None

    styles.append(row_style)

    for row in range(len(data)):
        tr = TableRow(stylename=row_style)
        cur_cols_count = len(data[row])
        for col in range(cur_cols_count):
            try:
                cell_style_name, content_cell_style_name, c = last_title_table_cells_styles[(cols-cur_cols_count)+col]
            except:
                pass
            el = data[row][col]
            if el.__contains__("colspan"):
                if el['color'] == 'yellow':
                    tc = TableCell(stylename='tablex', valuetype="string")
                else:
                    tc = TableCell(stylename=cell_style_name, valuetype="string")
            else:
                tc = TableCell(stylename=cell_style_name, valuetype="string")

            if type(el)==dict:
                name = el["label"]

                if el.__contains__("colspan"):
                    tc.setAttribute("numbercolumnsspanned", el["colspan"])

                # if el.__contains__("color"):
                #     tc.setAttribute("style-name", el["tablex"])

                if el.__contains__("rowspan"):
                    tc.setAttribute("numberrowsspanned", el["rowspan"])
                if el.__contains__("_width"):
                    col_width_val = el["width"]
                    col_width_name = "col_width_"+col.__str__()
                    col_style = Style(name=col_width_name, family="table-column")
                    col_style.addElement(TableColumnProperties(columnwidth=col_width_val))
                    styles.append(col_style)

                    table_column = TableColumn(stylename=col_style)
                    odf_element.addElement(table_column)

            else:
                name = el

            p = P(stylename=content_cell_style_name, text=name)
            # p.firstChild.attributes = c.firstChild.attributes
            tc.addElement(p)
            tr.addElement(tc)
        odf_element.addElement(tr)


    res = {
        "element": odf_element,
        "styles": styles,
    }
    return res


def create_odf_list(data_list: list):
    textList = List()
    for el in data_list:
        p = P()
        sub_list = List()
        if el.__contains__("text"):
            p.addText(el["text"])
            el.__delitem__("text")
        if el.__contains__("cdata"):
            p.addCDATA(el["cdata"])
            el.__delitem__("cdata")
        if el.__contains__("data"):
            sub_list = create_odf_list(el["data"])["element"]
            el.__delitem__("data")
        for attr in el:
            p.setAttribute(attr, el.get(attr))
        item = ListItem()
        item.addElement(p)
        item.addElement(sub_list)
        textList.addElement(item)

    res = {
        "element": textList,
        "style": None
    }
    return res
