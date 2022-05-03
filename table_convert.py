import os
import re
import json
import xlrd
import argparse
from multiprocessing import Pool


def is_array_t(text):
    return re.match("(int|int16|int32|bool|float|double|string)[\[\]]+$", text)

def is_dict_t(text):
    return re.match("dict<(int|int16|int32|bool|float|double|string),[a-zA-Z0-1]+>$", text)

def is_value_t(text):
    return re.match("(int|int16|int32|bool|float|double|string)$", text)


def is_fieldname(text):
    return re.match("[a-zA-Z0-9_]+", text)

def get_filename_without_extension(filepath):
    return os.path.splitext(os.path.split(filepath)[1])[0]


class ExcelTableFieldInfo(object):
    def __init__(self) -> None:
        self.order = 0
        self.fieldname = ""
        self.fieldtype = ""
        self.fieldvalue = None

        self.is_value_type = False
        self.is_array_type = False
        self.is_dict_type = False

        # if dict
        self.key_type = ""
        self.element_type = ""
        self.elemnet_types = {}

    def init(self, fieldname_text, fieldtype_text, field_order, element_types):
        self.order = field_order
        self.fieldname = fieldname_text
        self.fieldtype = fieldtype_text
        self.elemnet_types = element_types

        if is_value_t(fieldtype_text):
            self.is_value_type = True
        elif is_array_t(fieldtype_text):
            self.is_array_type = True
            self.element_type = fieldtype_text[:-2]
        elif is_dict_t(fieldtype_text):
            self.is_dict_type = True
            kv = fieldtype_text[3:-1].split(',')
            self.key_type = kv[0]
            self.element_type = kv[1]

    def get_meta_info(self):
        return {
            "order" : self.order,
            "fieldname" : self.fieldname,
            "fieldtype" : self.fieldtype,
            "fieldvalue" : self.fieldvalue,
            "is_value_type" : self.is_value_type,
            "is_array_type" : self.is_array_type,
            "is_dict_type" : self.is_dict_type,
            "key_type" : self.key_type,
            "element_type" : self.element_type,
            "elemnet_types" : self.elemnet_types,
        }


class ExcelTableInfo(object):
    def __init__(self) -> None:
        self.excel_file = ""
        self.table_name = ""
        self.table_type = "" #array, map, object
        self.element_type = ""
        self.out_file = ""
        self.out_field = ""
        self.out_root = False
        self.key_type = ""
        self.fields = []
        self.datas = []

    def get_meta_info(self):
        return {
            "out_file":self.out_file,
            "out_field":self.out_field,
            "out_root":self.out_root,
            "excel_file":self.excel_file,
            "table_name":self.table_name,
            "table_type":self.table_type,
            "element_type":self.element_type,
            "key_type":self.key_type,
            "fields":[x.get_meta_info() for x in self.fields],
        }


def convert_value(value, type_text):
    result = None
    val_type = type(value)
    if type_text == "byte" or type_text == "int16" or type_text == "int":
        if value is None:
            result = 0
        elif val_type == int or val_type == float:
            result = int(value)
        elif val_type == str:
            if value:
                try:
                    result = int(float(value))
                except:
                    result = 0
            else:
                result = 0
        elif val_type == bool:
            result = 1 if value else 0
        else:
            result = 0
    elif type_text == "float" or type_text == "double":
        if value is None:
            result = 0.0
        elif val_type == float or val_type == int:
            result = float(value)
        elif val_type == bool:
            result = 1.0 if value else 0.0
        elif value == str:
            if value:
                try:
                    result = float(value)
                except:
                    result = 0.0
            else:
                result = 0.0
        else:
            result = 0.0
    elif type_text == "bool":
        result = True if value else False
    elif type_text == "string":
        if value is None:
            result = ""
        else:
            result = str(value)
    return result


def convert(value, field):
    if field.is_value_type:
        return convert_value(value, field.fieldtype)
    try:
        return json.loads(value) if value else None
    except:
        return None


def parse_tabale_array(tb):
    """
    row1 field desc
    row2 field name
    row3 field type
    """
    datas = []
    fields = []
    field_map = {}
    element_types = {}
    order = 0
    for c in range(0, tb.ncols):
        fieldname = str(tb.cell(1, c).value).strip()
        fieldtype = str(tb.cell(2, c).value).strip()
        if fieldtype and fieldtype and is_fieldname(fieldname):
            field = ExcelTableFieldInfo()
            field_map[c] = field
            order = order + 1
            field.init(fieldname, fieldtype, order, element_types)
            fields.append(field)

    # parse datas
    data = {}
    for r in range(3, tb.nrows):
        for c in range(0, tb.ncols):
            if c in field_map:
                field = field_map[c]
                fieldname = field.fieldname
                fieldtype = field.fieldtype
                fieldvalue = tb.cell(r, c).value
                if type(fieldvalue) == str and fieldvalue.startswith("#"):
                    break
                fieldvalue = convert(fieldvalue, field)
                data[fieldname] = fieldvalue
        if data:
            datas.append(data)
            data = {}
    return (fields, datas)


def parse_table_object(tb):
    """
    col_0 is desc
    col_1 is field name
    col_2 is field type
    col_3 is field value
    """
    data = {}
    fields = []
    element_types = {}
    order = 0
    for r in range(0, tb.nrows):
        fieldname = str(tb.cell(r, 1).value).strip()
        fieldtype = str(tb.cell(r, 2).value).strip()
        fieldvalue = tb.cell(r, 3).value
        if fieldtype and fieldtype and is_fieldname(fieldname):
            order = order + 1
            field = ExcelTableFieldInfo()
            field.init(fieldname, fieldtype, order, element_types)
            fields.append(field)
            fieldvalue = convert(fieldvalue, field)
            field.fieldvalue = fieldvalue
            data[fieldname] = fieldvalue
    return (fields, data)


def parse_table(excelfile):
    result = {}
    wb = xlrd.open_workbook(excelfile, encoding_override="utf-8")
    for tb in wb.sheets():
        if tb.nrows < 1 or tb.ncols < 1:
            continue

        filename = get_filename_without_extension(excelfile)
        tb_info = ExcelTableInfo()
        tb_info.excel_file = filename
        tb_info.out_file = filename
        tb_info.out_field = tb.name
        tb_info.out_root = False
        tb_info.table_name = tb.name
        tb_info.table_type = str(tb.cell(0,0).value)
        tb_info.key_type = ""
        tb_info.element_type = tb.name

        if tb_info.table_type == "array":
            fields, datas = parse_tabale_array(tb)
        elif tb_info.table_type.startswith("dict"):
            fields, datas = parse_tabale_array(tb)
            tb_info.key_type = fields[0].fieldtype
        elif tb_info.table_type == "object":
            fields, datas = parse_table_object(tb)
        elif tb_info.table_type == "group":
            fields, datas = parse_tabale_array(tb)
            tb_info.key_type = fields[0].fieldtype

        tb_info.fields = fields
        tb_info.datas = datas

        result[filename] = tb_info
    return result


def parse_init_table(excelpath):
    result = {}
    tb = xlrd.open_workbook(excelpath, encoding_override="utf-8").sheet_by_index(0)
    for r in range(2, tb.nrows):
        item = {
            "excel_name" : str(tb.cell(r,0).value),
            "table_name" : str(tb.cell(r,1).value),
            "out_file" : str(tb.cell(r,2).value),
            "out_field" : str(tb.cell(r,3).value),
            "out_root" : bool(tb.cell(r,4).value),
        }
        if not (item["excel_name"] in result):
            result[item["excel_name"]] = {}
        dic = result[item["excel_name"]]
        dic[item["table_name"]] = item
    return result


def gen_meta(result_dict):
    meta = [v.get_meta_info() for k,v in result_dict.items()]
    # for k,v in result_dict.items():
    #     if not (v.table_name in meta):
    #         meta[v.table_name] = {
    #             "filename": "",
    #             "table_name": "",
    #             "table_type": "object",
    #             "out_name": "",
    #             "fields":[]
    #         }
    #     fields = meta[v.table_name]["fields"]
    #     if v.table_name:
    #         field = {
    #             "order": len(fields) + 1,
    #             "fieldname": v.table_name,
    #             "fieldtype": "{0}[]".format(v.element_type) if v.table_type == "array" else "dict<{0},{1}>".format(v.key_type, v.element_type) ,
    #             "fieldvalue": None,
    #             "is_value_type": False,
    #             "is_array_type": v.table_type == "array",
    #             "is_dict_type": v.table_type == "dict",
    #             "key_type": v.key_type,
    #             "element_type": v.element_type,
    #             "elemnet_types": {}
    #         }
    #         fields.append(field)
    return meta


def gen_output_datas(result_dict):
    output_data = {}
    for k in result_dict:
        tb_info = result_dict[k]
        tb_type = tb_info.table_type
        tb_data = tb_info.datas
        if tb_type.startswith("dict"):
            tb_dict = {}
            key_name = tb_info.fields[0].fieldname
            for data in tb_data:
                key = data.get(key_name, None)
                if key in tb_dict:
                    print("[{0}] repeated".format(key))
                tb_dict[key] = data
            tb_data = tb_dict
        elif tb_type.startswith("group"):
            tb_dict = {}
            key_name = tb_info.fields[0].fieldname
            arr = []
            prev_key = None
            if len(tb_data) > 1:
                prev_key = tb_data[0].get(key_name)
                tb_dict[prev_key] = arr
            for data in tb_data:
                key = data.get(key_name)
                if not key == prev_key:
                    prev_key = key
                    arr = []
                    tb_dict[key] = arr
                else:
                    arr.append(data)
            tb_data = tb_dict

        out_file = tb_info.out_file
        out_field = tb_info.out_field
        out_root = tb_info.out_root
        if out_root:
            output_data[out_field] = tb_data
        else:
            if not (out_file in output_data):
                output_data[out_file] = {}
            output_data[out_file][out_field] = tb_data
    return output_data


def to_luastr(data, retract = 1):
    if isinstance(data, (int, float)):
        return str(data)
    elif isinstance(data, str):
        return "\"{}\"".format(data)
    elif isinstance(data, list):
        lst = []
        lst.append("{\n")
        for one_data in data:
            lst.append("  " * retract)
            lst.append("{},\n".format(to_luastr(one_data, retract + 1)))
        lst.append("  " * (retract - 1))
        lst.append("}")
        return "".join(lst)
    elif isinstance(data, dict):
        lst = []
        lst.append("{\n")
        for key in data.keys():
            lst.append("  " * retract)
            lst.append("[{}] = {},\n".format(to_luastr(key), to_luastr(data[key], retract + 1)))
        lst.append( "  " * (retract - 1))
        lst.append( "}")
        return "".join(lst)
    return "nil"


def output_lua(info):
    filepath = info[0]
    data = info[1]
    code = to_luastr(data)
    luastr = "module = {0}\n\nreturn module".format(code)
    with open(filepath, encoding="utf-8", mode="w") as f:
        f.write(luastr)

def output_json(info):
    filepath = info[0]
    data = info[1]
    with open(filepath, mode="w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

def output_js(info):
    filepath = info[0]
    data = info[1]
    js = "let config = {0};\n\nexport default config".format(data)
    with open(filepath, encoding="utf-8", mode="w") as f:
        f.write(js)


def output_ts(info):
    filepath = info[0]
    data = info[1]
    ts = "let config = {0};\n\nexport default config".format(data)
    with open(filepath, encoding="utf-8", mode="w") as f:
        f.write(ts)


def main(args):
    excel_files = []
    init_xlsx_path = None
    for f in os.listdir(args.dir_excel):
        if f == "__init__.xls" or f == "__init__.xlsx":
            init_xlsx_path = os.path.join(args.dir_excel, f)
        if not(f.startswith("~") or f.startswith("_")) and (f.endswith(".xlsx") or f.endswith(".xls")):
            filepath = os.path.join(args.dir_excel, f)
            excel_files.append(filepath)

    # __init__.xlsx
    out_infos = None
    if init_xlsx_path and os.path.exists(init_xlsx_path):
        out_infos = parse_init_table(init_xlsx_path)

    pool = Pool(args.process)
    results = pool.map(parse_table, excel_files)
    result_dict = {}
    for rst in results:
        for k,v in rst.items():
            result_dict[k]=v
            if not(out_infos and (v.excel_file in out_infos)):
                continue
            out_info = out_infos.get(v.excel_file, None)
            if not(out_info and (v.table_name in out_info)):
                continue
            dic = out_info[v.table_name]
            if dic:
                v.out_root = dic.get("out_root", False)
                v.out_file = dic.get("out_file") or v.out_file
                v.out_field = dic.get("out_field") or v.out_field


    # output meta
    meta_filepath = os.path.join(args.dir_meta, "__meta__.json")
    meta = gen_meta(result_dict)
    with open(meta_filepath, mode="w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=True)

    exts = {
        "json":"json",
        "lua":"lua",
        "js":"js",
        "ts":"ts"
    }

    ext = exts[args.out_type]
    output_data = gen_output_datas(result_dict)
    output_data_infos = [(os.path.join(args.dir_data, "{0}.{1}".format(k, ext)), v) for k,v in output_data.items()]

    output_data_funcs = {
        "json":output_json,
        "lua":output_lua,
        "js":output_js,
        "ts":output_ts
    }

    func = output_data_funcs[args.out_type]
    pool.map(func, output_data_infos)



if __name__ == "__main__":
    assert xlrd.__version__ == "1.2.0", "xlrd==1.2.0. The lastest version does not support xlsx."

    args_parser = argparse.ArgumentParser()
    args_parser.add_argument("dir_excel", nargs="?", default="./excel", help="Excel directoy")
    args_parser.add_argument("dir_data", nargs="?", default="./json", help="Data directoy")
    args_parser.add_argument("dir_meta", nargs="?", default="./meta", help="Meta info directory")
    args_parser.add_argument("out_type", nargs="?", default= "json", help="json/lua/js/ts")
    args_parser.add_argument("process", nargs="?", default=5)
    args = args_parser.parse_args()

    main(args)
