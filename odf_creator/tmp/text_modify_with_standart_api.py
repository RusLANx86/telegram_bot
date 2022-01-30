from odf.userfield import UserFields


obj = UserFields(src='output_file.odt')
print(obj.list_fields())
upd_dict = {"name1": "NEW TEXT!!!!"}
obj.update(upd_dict)
