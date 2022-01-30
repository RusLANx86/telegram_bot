from odf.opendocument import load, OpenDocument
from return_odf_element import update_odf_element

import os

test_temp = os.path.abspath('templates/checks_template.odt')

data = {
    "table1":
        {
            "type": "Table",
            "table_title": [
                [
                    "номер по порядку",
                    "звание",
                    "фамилия, имя, отчество",
                    "вид справки",
                    "основание представления",
                    "поступила: (откуда, дата, номер документа)",
                    "отравлена: (куда, дата, номер документа)"
                ]
            ],
            "data": [[{'label':'хуй вам', "colspan": 7, 'color': "yellow"},],
                     ['х', 'у', 'й','-','в','а','м'],
                     ['х', 'у', 'й','-','в','а','м'],
                     ['х', 'у', 'й','-','в','а','м'],
                     ['х', 'у', 'й','-','в','а','м'],
                     ['х', 'у', 'й','-','в','а','м'],
                     ['х', 'у', 'й','-','в','а','м'],
                     ]
        },
}
# data = {
#             "famaly":
#                 {
#                     "type": "String",
#                     "data": 'Иван Фёдорофич Крузенштерн'
#                 },
#         }
#

def fill_template(template_file_path: str, json_data: dict) -> OpenDocument:

    file = load(template_file_path)
    text = file.text

    def update_user_fields(odf_element):
        for odf_elem in odf_element.childNodes:
            try:
                user_field_name = odf_elem.getAttribute("name")
                data = json_data.get(user_field_name)
                if data != None:
                    res = update_odf_element(element=odf_elem, data=data)
                    try:
                        styles = res["styles"]
                        for style in styles:
                            file.automaticstyles.addElement(style)
                    except:
                        pass
            except:
                update_user_fields(odf_elem)

        return odf_element

    update_user_fields(odf_element=text)


    from uuid import uuid4
    st = str(uuid4())
    output_file = os.path.abspath("report_files/{}.odt".format(st))

    # file.contentxml()

    file.save(output_file)
    # return file


if __name__ == '__main__':
    fill_template(test_temp, data)