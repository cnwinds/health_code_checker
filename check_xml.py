import logging
import os
import time
from PIL import Image
import json

import xml_parse
from xml_parse import get_node_by_keyvalue, read_xml, read_xml_remove_ns
import yinshua
import re

import sys, os, zipfile

def unzip_single(src_file, dest_dir, password = None):
    ''' 解压单个文件到目标文件夹。
    '''
    if password:
        password = password.encode()
    zf = zipfile.ZipFile(src_file)
    try:
        zf.extractall(path=dest_dir, pwd=password)
    except RuntimeError as e:
        logging.error(e)
    zf.close()

""" 读取图片 """
def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

def parse_str(file):
    root = read_xml_remove_ns(file)

    str_list = []
    items = root.findall('./si')
    for child in items:
        i1 = child.findall("./t")
        if len(i1) > 0:
            str_list.append(i1[0].text.replace(" ", ""))
        else:
            str_list.append("")

    return str_list

img_pattern = re.compile(r'.*(".*").*')

def parse_sheet(sheet_file, share_str_list, imgs_list):

    root = read_xml_remove_ns(sheet_file)
    if root is None:
        logging.error("表格中有浮动图片，请转成单元格图片后处理！")
        os._exit(1)

    family_list = []
    family = None
    # family = {"stu" :{"name":"", "img":""}, "members": [{"name":"","img":""}]}
    # 获取家庭成员信息
    items = root.findall('./sheetData/row')
    for child in items: # row iter
        line_s = child.attrib['r']
        line = int(line_s)
        if line < 4:
            continue
        for c1 in child: # c iter
            if 't' in c1.attrib:
                
                if c1.attrib["r"][:1] == "C":  # 重置
                    if family:
                        family_list.append(family)
                    family = {"stu":{"name":"", "img":""}, "members":{"imgs":[],"names":[]}}
                if c1.attrib['t'] == 's':   # 文字处理
                    if c1.attrib["r"][:1] == "C":  # 学生名字
                        value = c1.find("./v")
                        family["stu"]["name"] = share_str_list[int(value.text)]
                    if c1.attrib["r"][:1] == "F":  # 成员姓名
                        value = c1.find("./v")
                        family["members"]["names"].append(share_str_list[int(value.text)])
                elif c1.attrib['t'] == 'str': # 图片处理
                    if c1.attrib["r"][:1] == "D":  # 学生图片
                        value = c1.find("./v")
                        search = img_pattern.search(value.text)
                        value = search.group(1)[1:-1]
                        value = imgs_list[value]
                        family["stu"]["img"] = value
                    if c1.attrib["r"][:1] >= "I" and c1.attrib["r"][:1] <= "N":  # 成员图片
                        value = c1.find("./v")
                        search = img_pattern.search(value.text)
                        value = search.group(1)[1:-1]
                        value = imgs_list[value]
                        family["members"]["imgs"].append(value)

    if family:
        family_list.append(family)
    
    return family_list


def parse_img(file, rels_file):
    root = read_xml(rels_file).getroot()
    ns = {  'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': "http://schemas.openxmlformats.org/drawingml/2006/main" }
    id_file_dict = {}
    for item in root:
        id_file_dict[item.attrib["Id"]] = item.attrib["Target"].replace("../","")
    
    root = read_xml(file).getroot()
    img_file_dict = {}
    for item in root:
        i1 = item.find("./xdr:pic/xdr:nvPicPr/xdr:cNvPr", ns)
        i2 = item.find("./xdr:pic/xdr:blipFill/a:blip", ns)
        file_name = id_file_dict[i2.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"]]
        img_file_dict[i1.attrib["name"]] = file_name

    return img_file_dict

def transform_filepath(f):
    return tmp_dir + "/" + f


def ocr_img(file):
    ocr_dict = {"name":"","date":""}
    result = json.loads(yinshua.get_content(file))
    if result['code'] == '0':
        last_value = ""
        items = result['data']['block'][0]['line']
        for i in items:
            value = i['word'][0]['content']

            # 判断是否当天时间
            if value[:3] == '更新于':
                ocr_dict["date"] = value
                ocr_dict["name"] = last_value
                break
            last_value = value
    return result['code'] == '0', result, ocr_dict


def del_name(err_names, name):
    for i in err_names.keys():
        if i.find(name) >= 0:
            err_names.pop(i)
            return True
    return False
    
if __name__ == '__main__':
    LOG_FORMAT = "[%(asctime)s] - %(levelname)s - %(message)s"
    logging.basicConfig(filename='check.log', level=logging.INFO, format=LOG_FORMAT)
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    logging.getLogger('').addHandler(console)

    today = time.strftime("%Y-%m-%d")
    tmp_dir = 'tmp'

    import shutil  
    if os.path.exists(tmp_dir) and shutil.rmtree(tmp_dir):
        pass

    if len(sys.argv) < 2:
        logging.error("请带上需要检查的文件")
        os._exit(0)
    else:
        xlsfilename = sys.argv[1]
    
    logging.info("### 开始处理文件[{0}] ###".format(xlsfilename))

    unzip_single(xlsfilename, tmp_dir)

    pic_success = 0
    pic_unknow = 0
    pic_total = 0
    pic_date_err = 0
    pic_name_err = 0

    # 解析文件集
    root = read_xml_remove_ns(transform_filepath("[Content_Types].xml"))
    items = root.findall("./Override")

    share_string_file_list = []
    share_string_node_list = get_node_by_keyvalue(items, {"ContentType":"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"})
    for item in share_string_node_list:
        share_string_file_list.append(item.attrib["PartName"])
    share_str_list = []
    for f in share_string_file_list:
        share_str_list = share_str_list + parse_str(transform_filepath(f))

    cell_image_file_list = []
    for item in get_node_by_keyvalue(items, {"ContentType":"application/vnd.wps-officedocument.cellimage+xml"}):
        cell_image_file_list.append(item.attrib["PartName"])
    for item in get_node_by_keyvalue(items, {"ContentType":"application/vnd.openxmlformats-officedocument.drawing+xml"}):
        cell_image_file_list.append(item.attrib["PartName"])

    imgs_list = {}
    for f in cell_image_file_list:
        pos = f.rfind('/')
        rels_file = f[:pos] + "/_rels" + f[pos:] + ".rels"
        imgs_list = dict(imgs_list, **parse_img(transform_filepath(f), transform_filepath(rels_file)))

    # logging.info("imgs_list",imgs_list)
    # logging.info(share_str_list)

    sheet_file_list = []
    sheet_node_list = get_node_by_keyvalue(items, {"ContentType":"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
    for item in sheet_node_list:
        sheet_file_list.append(item.attrib["PartName"])

    family_list = []
    for f in sheet_file_list:
        family_list = family_list + parse_sheet(transform_filepath(f), share_str_list, imgs_list)
        # 只查第一个sheet
        break

    # 得到所有家庭信息列表
    logging.info("共有学生{0}人。".format(len(family_list)))

    idx = 1
    for i in family_list:
       
        logging.info(">>> ({1}/{2}) 正在处理学生[{0}] <<<".format(i["stu"]['name'], idx, len(family_list)))
        idx = idx + 1

        err_names = {}
        err_imgs = {}
        for j in i["members"]["names"]:
            err_names[j] = "没有识别到二维码"
        
        # 处理学生二维码
        f = i["stu"]["img"]
        if f == '':
            err_names[i["stu"]['name']] = "没有上传二维码"
            err_imgs[fp] = ocr_dict['name']
        else:
            fp = transform_filepath("xl/" + f)
            s, r, ocr_dict = ocr_img(fp)
            if s:
                if ocr_dict['date'].find(today) < 0:
                    err_names[i["stu"]['name']] = "二维码日期错误"
                    err_imgs[fp] = ocr_dict['name']
                elif ocr_dict['name'] != i["stu"]['name']:
                    err_names[i["stu"]['name']] = "没有识别到二维码"
                    err_imgs[fp] = ocr_dict['name']
            else:
                err_names[i["stu"]['name']] = "图片识别接口调用错误"
                err_imgs[fp] = "图片识别接口调用错误:" + r

        # 处理成员二维码
        imgs = i["members"]["imgs"]
        for f in imgs:
            fp = transform_filepath("xl/" + f)
            s, r, ocr_dict = ocr_img(fp)
            if s:
                if ocr_dict['date'] == '' and ocr_dict['name'] == '':
                    err_imgs[fp] = "图片无法识别"
                else:
                    if ocr_dict['date'].find(today) < 0:
                        del_name(err_names, ocr_dict['name'])
                        err_names[ocr_dict['name']] = "二维码日期错误"
                        err_imgs[fp] = ocr_dict['name']
                    else:
                        if del_name(err_names, ocr_dict['name']) == False:
                            err_names[ocr_dict['name']] = "二维码姓名错误"
                            err_imgs[fp] = ocr_dict['name']
            else:
                err_imgs[fp] = "图片识别接口调用错误:" + r

        # 统一错误提示
        if len(err_names) > 0: 
            logging.error("问题：学生[{0}]的问题有{1}".format(i["stu"]['name'], err_names))
        
        if len(err_imgs) > 0:
            # logging.error("学生[{0}]不可识别的图片有{1}".format(i["stu"]['name'], err_imgs))
            for j in err_imgs:
                logging.error("识别的名称[{0}]".format(err_imgs[j]))
                img=Image.open(j)
                img.show()

    logging.info("### 结束处理文件[{0}] ###".format(xlsfilename))
