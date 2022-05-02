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

def is_fieldtype(text):
    return is_value_t(text) or is_array_t(text) or is_dict_t(text)

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
        self.filename = ""
        self.table_name = ""
        self.table_type = "" #array, map, object
        self.classname = ""
        self.out_name = ""
        self.out_file = ""
        self.out_root = False
        self.element_type = ""
        self.key_type = ""
        self.fields = []
        self.datas = []

    def get_meta_info(self):
        return {
            "filename":self.filename,
            "table_name":self.table_name,
            "table_type":self.table_type,
            "classname":self.classname,
            "out_name":self.out_name,
            "out_file":self.out_file,
            "element_type":self.element_type,
            "key_type":self.key_type,
            "out_root":self.out_root,
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


def convert_array(array_text, element_type, split_text=","):
    result = []
    if array_text:
        if len(array_text) >2 and array_text[0] == "[" and array_text[1] == "[":
            result = json.loads(array_text)
        else:
            if array_text[0] == "[" and array_text[-1] == "]":
                array_text = array_text[1:-1]
            result = [convert_value(x.strip(), element_type) for x in array_text.split(split_text)]
    return result


def convert_dict(dict_text, element_type_dict):
    result = {}
    if dict_text:
        if dict_text[0] == "{" and dict_text[-1] == "}":
            try:
                result = json.loads(dict_text)
            except:
                arr = [x.strip() for x in dict_text[1:-1].split(",")]
                for x in arr:
                    arr2 = [y.strip() for y in x.split(":")]
                    k = arr2[0]
                    element_type = element_type_dict.get(k, "string")
                    v = convert_value(arr2[1], element_type)
                    result[k] = v
    return result


def convert(value, field):
    if field.is_value_type:
        return convert_value(value, field.fieldtype)
    elif field.is_array_type:
        return convert_array(value, field.element_type)
    elif field.is_dict_type:
        return convert_dict(value, field.element_types)


def parse_tabale_array(tb):
    """
    row0 meta : {"type":"array/dict<key_fieldname,object/array>", "classname":"Test", "out_file":"Config", "out_name":"heroes", "out_root":false}
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
        fieldname = str(tb.cell(2, c).value).strip()
        fieldtype = str(tb.cell(3, c).value).strip()
        if not (is_fieldname(fieldname) and is_fieldtype(fieldtype)):
            continue
        field = ExcelTableFieldInfo()
        field_map[c] = field
        order = order + 1
        field.init(fieldname, fieldtype, order, element_types)
        fields.append(field)

    # parse datas
    data = {}
    for r in range(4, tb.nrows):
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
    row0 meta : {"type":"object/array/dict<key_fieldname,object/array>", "classname":"Test", "out_file":"Config", "out_root":true}
    col_0 is desc
    col_1 is field name
    col_2 is field type
    col_3 is field value
    """
    data = {}
    fields = []
    element_types = {}
    order = 0
    for r in range(1, tb.nrows):
        fieldname = str(tb.cell(r, 1).value).strip()
        fieldtype = str(tb.cell(r, 2).value).strip()
        fieldvalue = tb.cell(r, 3).value
        if is_fieldname(fieldname) and is_fieldtype(fieldtype):
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
        meta_text = str(tb.cell(0,0).value)
        meta = convert_dict(meta_text, {}) if meta_text else {}

        filename = get_filename_without_extension(excelfile)
        tb_info = ExcelTableInfo()
        tb_info.filename = filename
        tb_info.classname = meta.get("classname", tb.name)
        tb_info.table_name = tb.name
        tb_info.table_type = meta.get("type", "array")
        tb_info.out_file = meta.get("out_file", filename)
        tb_info.out_name = meta.get("out_name", tb.name)
        tb_info.out_root = meta.get("out_root", False)
        tb_info.key_type = ""
        tb_info.element_type = tb_info.classname

        if tb_info.table_type == "array":
            fields, datas = parse_tabale_array(tb)
        elif tb_info.table_type.startswith("dict"):
            fields, datas = parse_tabale_array(tb)
            idx = tb_info.table_type.index(",")
            keyid = tb_info.table_type[5:idx]
            for f in fields:
                if f.fieldname == keyid:
                    tb_info.key_type = f.fieldtype
                    break
        elif tb_info.table_type == "object":
            fields, datas = parse_table_object(tb)

        tb_info.fields = fields
        tb_info.datas = datas

        result[filename] = tb_info
    return result


proto_type_map = {
    "int":"int32",
    "int16":"int32",
    "float":"float",
    "double":"double",
    "string":"string",
    "bool":"bool"
}

def get_proto_type(text):
    if text in proto_type_map:
        return proto_type_map[text]
    if text.startswith("array<"):
        t = get_proto_type(text[6:-1])
        return "repeated {0}".format(t)
    if text.startswith("dict<"):
        text = text[5:-1]
        arr = text.split(",")
        kt = get_proto_type(arr[0])
        vt = get_proto_type(arr[1])
        return "map<{0},{1}>".format(kt, vt)
    return text



def gen_proto(meta, out_file):
    classnames = []
    exprs = []
    exprs.append('syntax = "proto3";\n\n')
    for k,v in meta.items():
        classname = v["classname"]
        if classname in classnames:
            continue
        classnames.append(classname)
        exprs.append("message {0}".format(classname))
        exprs.append("{\n")
        for f in v["fields"]:
            proto_type = get_proto_type(f["fieldtype"])
            exprs.append("    {0} {1} = {2};\n".format(proto_type, f["fieldname"], f["order"]))
        exprs.append("}\n\n")

    with open(out_file, mode="w", encoding="utf-8") as f:
        f.writelines(exprs)


def gen_meta(result_dict):
    meta = {k:v.get_meta_info() for k,v in result_dict.items()}
    for k,v in result_dict.items():
        if not (v.out_file in meta):
            meta[v.out_file] = {
                "filename": "",
                "table_name": "",
                "table_type": "object",
                "classname": v.out_file,
                "out_name": "",
                "out_file": v.out_file,
                "out_root": True,
                "fields":[]
            }
        fields = meta[v.out_file]["fields"]
        if v.out_name:
            field = {
                "order": len(fields) + 1,
                "fieldname": v.out_name,
                "fieldtype": "{0}[]".format(v.element_type) if v.table_type == "array" else "dict<{0},{1}>".format(v.key_type, v.element_type) ,
                "fieldvalue": None,
                "is_value_type": False,
                "is_array_type": v.table_type == "array",
                "is_dict_type": v.table_type == "dict",
                "key_type": v.key_type,
                "element_type": v.element_type,
                "elemnet_types": {}
            }
            fields.append(field)
    return meta


def gen_output_datas(result_dict):
    output_data = {}
    for k in result_dict:
        tb_info = result_dict[k]
        tb_type = tb_info.table_type
        tb_data = tb_info.datas
        tb_dict = {}
        out_file = tb_info.out_file
        out_name = tb_info.out_name
        out_root = tb_info.out_root
        if tb_type.startswith("dict"):
            # dict<id,object>
            dot_idx = tb_type.index(",")
            key_name = tb_type[5:dot_idx]
            val_type = tb_type[dot_idx+1:-1]
            if val_type == "object":
                for data in tb_data:
                    key = data.get(key_name, None)
                    if key in tb_dict:
                        print("[{0}] repeated".format(key))
                    tb_dict[key] = data
            else:
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

        if out_root or not out_name:
            output_data[out_file] = tb_data
        else:
            if not (out_file in output_data):
                output_data[out_file] = {}
            output_data[out_file][out_name] = tb_data
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


def gen_csharp(excelinfo, dir_template, out_dir):
    filepath = os.path.join(dir_template, "cs.txt")
    field_prefix_space = ""
    with open(filepath, mode='r', encoding="utf-8") as f:
        lines = f.readlines()
        for i in range(0, len(lines)):
            idx = lines[i].find("{fields}")
            if idx > -1:
                field_prefix_space = lines[i][:idx] + ""
                lines[i] = "{fields}"

    template = "".join(lines)

    root = {}
    for k in excelinfo:
        tbinfo = excelinfo[k]
        table_type = tbinfo.table_type
        out_file = tbinfo.out_file
        filename = tbinfo.filename
        out_name = tbinfo.out_name
        out_root = tbinfo.out_root
        classname = tbinfo.classname
        fields = tbinfo.fields

        if not (out_file in root) and out_root == False:
            root[out_file] = []
        if out_name and not out_root:
            if table_type == "array":
                root[out_file].append("{0}public {1}[] {2};//{3}\n".format(field_prefix_space, classname, out_name, filename))
            if table_type.startswith("dict"):
                idx = table_type.index(",")
                keyid = table_type[5:idx]
                for f in fields:
                    if f.fieldname == keyid:
                        key = f.fieldtype
                        root[out_file].append("{0}public Dictionary<{1}, {2}> {3};//{4}\n".format(field_prefix_space, key, classname, out_name, filename))
                        break

        # sub
        exprs = []
        for field in fields:
            fieldname = field.fieldname
            fieldtype = field.fieldtype
            fieldvalue = field.fieldvalue
            if field.is_array_type:
                # fieldtype = "{0}[]".format(field.element_type)
                if fieldvalue:
                    count = fieldtype.count("[")
                    if count > 2:
                        print("complex array not export")
                    elif count == 2:
                        #[[1,2,3], [1,2,3]]
                        text = str(fieldvalue)[1:-1]
                        arr = [x.strip() for x in text.split("], [")]
                        t = fieldtype[:-4]
                        val_arr = []
                        val_arr.append("new {0}".format(fieldtype))
                        val_arr.append("\n    {\n")
                        if len(arr) > 1:
                            arr[0] = arr[0][1:]
                            arr[-1] = arr[-1][:-1]
                        for x in arr:
                            val_arr.append("        new {0}{1}".format(t, "{"+x+"},\n"))
                        val_arr.append("\n    }")
                        fieldvalue = "".join(val_arr)

                    else:
                        fieldvalue = "new {0}{1}".format(fieldtype, "{"+str(fieldvalue)[1:-1]+"}")
            if field.is_dict_type:
                fieldtype = field.fieldtype.replace("dict", "Dictionary")
                if fieldvalue:
                    fieldvalue = None
                    print("cann't process dict")

            if fieldvalue is None:
                exprs.append("{0}public {1} {2};\n".format(field_prefix_space, fieldtype, fieldname))
            else:
                exprs.append("{0}public {1} {2} = {3};\n".format(field_prefix_space, fieldtype, fieldname, fieldvalue))

        cs_file = os.path.join(out_dir, classname+".cs")
        with open(cs_file, mode="w", encoding="utf-8") as f:
            code = template.replace("{classname}", classname).replace("{fields}", "".join(exprs))
            f.write(code)

    for out_file, arr in root.items():
        if arr:
            exprs = list(arr)
            cs_file = os.path.join(out_dir, out_file+".cs")
            with open(cs_file, mode="w", encoding="utf-8") as f:
                code = template.replace("{classname}", out_file).replace("{fields}", "".join(exprs))
                f.write(code)



def main(args):
    excel_files = []
    for f in os.listdir(args.dir_excel):
        if not f.startswith("~") and (f.endswith(".xlsx") or f.endswith(".xls")):
            filepath = os.path.join(args.dir_excel, f)
            excel_files.append(filepath)

    pool = Pool(args.process)
    results = pool.map(parse_table, excel_files)
    result_dict = {}
    for rst in results:
        for k,v in rst.items():
            result_dict[k]=v

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


    if args.script:
        output_code_funcs = {
            "c#":gen_csharp
        }
        scripts = [x.strip() for x in args.script.split(",")]
        for s in scripts:
            output_code_funcs[s](result_dict, args.dir_template, args.dir_script)





if __name__ == "__main__":
    assert xlrd.__version__ == "1.2.0", "xlrd==1.2.0. The lastest version does not support xlsx."

    args_parser = argparse.ArgumentParser()
    args_parser.add_argument("dir_excel", nargs="?", default="./excels", help="Excel directoy")
    args_parser.add_argument("dir_data", nargs="?", default="./json", help="Data directoy")
    args_parser.add_argument("dir_proto", nargs="?", default="./protos", help="Protos directoy")
    args_parser.add_argument("dir_script", nargs="?", default="./scripts", help="Scrips directoy")
    args_parser.add_argument("dir_template", nargs="?", default="./templates", help="Scrip templates directoy")
    args_parser.add_argument("dir_meta", nargs="?", default="./meta", help="Meta info directory")
    args_parser.add_argument("out_type", nargs="?", default= "json", help="json/lua")
    args_parser.add_argument("script", nargs="?", default="c#", help="c#")
    args_parser.add_argument("process", nargs="?", default=5)
    args = args_parser.parse_args()

    main(args)
