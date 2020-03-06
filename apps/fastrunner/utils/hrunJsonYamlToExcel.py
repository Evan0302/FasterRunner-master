import json
import io
import os

import xlwt
import yaml


class Hrun_JsonYaml:

    def __init__(self, infilename='', outfilename='', jsonObj=None):
        self.fileName = infilename
        self.excelfile = outfilename
        self.objContent = jsonObj
        self._rowIndex = 0
        self.excel = None
        self._listFields = []  # 字段列表
        self._fieldRow = 0  # 字段行号

    def load_file(self, sfilename=''):
        if sfilename != '': self.fileName = sfilename
        if not os.path.isfile(self.fileName):
            raise Exception("{} does not exist.".format(self.fileName))

        file_suffix = os.path.splitext(self.fileName)[1].lower()
        if file_suffix == '.json':
            self.objContent = self.loadJson(self.fileName)
        elif file_suffix in ['.yaml', '.yml']:
            self.objContent = self.loadYaml(self.fileName)
        elif file_suffix == ".csv":
            return None  # load_csv_file(self.fileName)
        else:
            # '' or other suffix
            err_msg = u"Unsupported file format: {}".format(self.fileName)

            return None

        return self.objContent

    def loadYaml(self, filename):
        with io.open(filename, 'r', encoding='utf-8') as yfile:
            return yaml.full_load(yfile)

    def loadJson(self, filename):
        with io.open(filename, 'r', encoding='utf-8') as jfile:
            return json.load(jfile)

    _caseItemIndex = 0

    def getListNextItem(self, iName):
        if self._caseItemIndex >= len(self.objContent):
            return None

        if isinstance(self.objContent, dict) and ('iName' not in self.objContent):
            return None

        oCase = self.objContent[self._caseItemIndex]
        if iName in oCase:
            self._caseItemIndex += 1
            return oCase
        else:
            return None

    def initSheet(self):
        from xlwt.Style import default_style
        self.excel = xlwt.Workbook()

        from xlwt import Alignment
        default_style.alignment.wrap = Alignment.WRAP_AT_RIGHT

        self.sheetname = os.path.basename(self.fileName)
        self.sheet = self.excel.add_sheet(self.sheetname)

    def toExcelWorkbook(self):
        if self.excel ==None :
            self.initSheet()

        # 生成sheet
        self.newSheet()

        return self.excel

    def toExcel(self):
        self.toExcelWorkbook()
        if os.path.exists(self.excelfile): os.remove(self.excelfile)
        self.excel.save(self.excelfile)

    def newSheetRow(self, cell, svalue):
        self.sheet.write(*cell, svalue)

    def newSheet(self):
        # self.sheetname = os.path.basename(self.fileName)
        # self.sheet = self.excel.add_sheet(self.sheetname)
        # 生成config
        self.newConfig(self.sheet)
        # 生成testcase
        self.newTestcase(self.sheet)

        self.autoExcelWidth(self.sheet)

    def newConfig(self, sheet):
        oconfig = self.getListNextItem('config')
        if oconfig is None:
            # 添加空的config
            sheet.write(self._rowIndex, 0, "config")
            self._rowIndex += 1
            sheet.write(self._rowIndex, 0, "=")
            sheet.write(self._rowIndex, 1, "name")
            sheet.write(self._rowIndex, 2, "base_url")
            self._rowIndex += 2
            return

        sheet.write(self._rowIndex, 0, "config")
        self._rowIndex += 1

        icol = 1
        sheet.write(self._rowIndex, 0, '=')
        for (k, v) in oconfig['config'].items():
            sheet.write(self._rowIndex, icol, k)
            if k in ['variables']:
                v = self.obj2Lines(v)
            sheet.write(self._rowIndex + 1, icol, str(v))
            icol += 1

        self._rowIndex += 2

    def _addField(self, fieldname):
        self.sheet.write(self._fieldRow, len(self._listFields)+1, fieldname)
        self._listFields.append(fieldname)

    def fieldCol(self, fieldName):
        """
        根据字段名，返回字段序号
        :param fieldName:
        :return:
        """
        if fieldName not in self._listFields:
            self._addField(fieldName)
        return self._listFields.index(fieldName) + 1

    def writeFieldRow(self, otestcase, sheet):

        icol = 1
        sheet.write(self._rowIndex, 0, '=')
        self._fieldRow = self._rowIndex

        for (k, v) in otestcase['test'].items():
            if k.startswith('request'):
                for (k, v) in v.items():
                    self._addField('request.' + str(k))
                    # self.sheet.write(self.sheetIndex, icol, 'request.' + str(k))
                    # self.listFields.append(k)
                    icol += 1
            else:
                self._addField(k)
                icol += 1
        self._rowIndex += 1

    def autoExcelWidth(self, sheet):
        # 获取字符串长度，一个中文的长度为2
        def len_byte0(value):
            length = len(value)
            utf8_length = len(value.encode('utf-8'))
            length = (utf8_length - length) / 2 + length
            return int(length)

        def len_byte(value):
            if value is None or value == "":
                return 10
            if value.find("\n") > 0:
                pass

            if type(value) != int:
                length = len(value)
                utf8_length = len(value.encode('utf-8'))
                length = (utf8_length - length) / 2 + length
            else:
                length = len(str(value))
            return int(length)

        for i, ofield in enumerate(self._listFields):
            if i + 1 in [1, 2, 3]:
                sheet.col(i + 1).width = 256 * 20
            elif ofield in ['extract', 'variables']:
                sheet.col(i + 1).width = 256 * 30
            elif ofield in ['request.headers', 'request.json', 'validate']:
                sheet.col(i + 1).width = 256 * 40

        # 设置行宽 col()中的参数为第几列
        # 256表示是一个字符的长度，256 * 30 就是四十个字符的长度
        # for i in range(0, len(self.sheet.cols)):
        #     exec('first_col_{} = sheet.col({})'.format(i + 1, i))
        #     exec('first_col_{}.width = 256 * 30'.format(i + 1))

        # # 确定栏位宽度
        # col_width = []
        # for i in range(len(result)):
        #     for j in range(len(result[i])):
        #         if i == 0:
        #             col_width.append(len_byte(result[i][j]))
        #         else:
        #             if col_width[j] < len_byte(str(result[i][j])):
        #                 col_width[j] = len_byte(result[i][j])
        #
        # # 设置栏位宽度，栏位宽度小于10时候采用默认宽度
        # for i in range(len(col_width)):
        #     if col_width[i] > 10:
        #         self.sheet.col(i).width = 256 * (col_width[i] + 1)

    def newTestcase(self, sheet):
        def writeFieldValue(fieldname, value):
            sheet.write(self._rowIndex, self.fieldCol(fieldname), str(value))

        self._rowIndex += 1
        sheet.write(self._rowIndex, 0, "teststeps")
        self._rowIndex += 1

        isFRow = False
        while True:
            otestcase = self.getListNextItem('test')
            if otestcase == None: break

            if not isFRow:  # 填写表头字段名  行
                self.writeFieldRow(otestcase, sheet)
                isFRow = not isFRow

            icol = 1
            for (k, v) in otestcase['test'].items():
                if k.startswith('request'):
                    for (k, v) in v.items():
                        if k in ['headers', 'json']:
                            v = self.obj2Lines(v)
                        writeFieldValue('request.' + k, v)
                        icol += 1
                else:
                    if k in ['validate', 'extract', 'variables']:
                        v = self.obj2Lines(v)

                    elif k in ['validate']:
                        v = self.obj2Lines(v)

                    writeFieldValue(k, v)
                    icol += 1
            self._rowIndex += 1



    def obj2Lines(self, ovalue):
        if isinstance(ovalue, list):
            return self.list2Lines(ovalue)
        elif isinstance(ovalue, dict):
            return self.dict2Lines(ovalue)

    def list2Lines(self, ovalue):
        xlist = []
        for o in ovalue:
            # tt = ','.join([k + ': ' + str(v) for k, v in o.items()])
            xlist.append(self.valueObj2Str(o))
        return '\n'.join(xlist)

    def dict2Lines(self, ovalue):
        return '\n'.join([k + ': ' + str(v) for k, v in ovalue.items()])

    def valueObj2Str(self, ovalue):
        if isinstance(ovalue, list):
            return str(ovalue)
        elif isinstance(ovalue, dict):
            return ','.join([k + ': ' + str(v) for k, v in ovalue.items()])
        else:
            return str(ovalue)


def loadfile2excel():
    infile = "e:/git_home/httprunner/tests/data/demo_testcase_hardcode.yml"
    # "e:/git_home/httprunner/docs/data/demo-quickstart.json"

    hj = Hrun_JsonYaml(infile, 'yy.xls')
    hj.load_file()

    hj.toExcel()

    print(hj.objContent)


def loadjsonstr2excel():
    jj = """[{'test': {'name': 'find place from text', 'request': {'url': '/maps/api/place/findplacefromtext/json', 'method': 'GET', 'verify': False, 'params': {'key': '$key', 'inputtype': 'textquery', 'input': '$input', 'fields': 'formatted_address,geometry,name,permanently_closed,place_id,plus_code,types', 'language': 'zh_CN'}, 'json': {'json': {}}}, 'extract': [{'place_id': 'content.candidates.0.place_id'}], 'validate': [{'eq': ['status_code', 200]}, {'eq': ['content.status', 'OK']}], 'variables': [{'input': '宝安'}, {'key': 'your_google_map_key'}]}}, {'test': {'name': 'get_groups', 'request': {'url': '/v1/transferairportadminsrv/area/getGroups', 'method': 'POST', 'verify': False, 'headers': {'Content-Type': 'application/x-www-form-urlencoded', 'Accept-language': 'zh_CN', 'cookie': '$admin_cookie'}, 'data': {'parentAreaId': '7037\xa0#\xa0广东省'}, 'json': {'json': {}}}}}, {'test': {'name': 'get_place_detail', 'request': {'url': 'https://maps.googleapis.com/maps/api/place/details/json', 'method': 'GET', 'verify': False, 'params': {'key': '$key', 'placeid': '$place_id', 'fields': 'address_component,formatted_address,geometry', 'language': 'zh-CN'}, 'json': {'json': {}}}, 'extract': [{'place_name': 'content.result.formatted_address'}], 'validate': [{'eq': ['status_code', 200]}, {'eq': ['content.status', 'OK']}], 'variables': [{'key': 'your_google_map_key'}, {'place_id': 'ChIJzyoujG6SAzQRRD3Jr26PFfM'}]}}, {'test': {'name': 'search_area_by_name', 'request': {'url': '/v1/transferairportadminsrv/area/search_area_by_name', 'method': 'GET', 'verify': False, 'headers': {'Content-Type': 'application/x-www-form-urlencoded', 'Accept-language': 'zh_CN', 'cookie': '$admin_cookie'}, 'params': {'areaName': '$in'}, 'json': {'json': {}}}}}, {'test': {'name': 'demo_api1', 'request': {'url': '/api/$var1', 'method': 'POST', 'verify': False, 'headers': {'Content-Type': '"application/json"'}, 'json': {'json': {'key': '$var2'}}}, 'validate': [{'eq': ['status_code', 200]}], 'variables': [{'var1': 'value1'}, {'var2': 'value2'}]}}]"""
    oj = eval(jj)
    # oj.pop('desc')
    # oj.pop('times')
    # oj = [{'test': oj}]

    print(oj)
    hj = Hrun_JsonYaml(jsonObj=oj, infilename='apiexport')
    hj.toExcelWorkbook().save('/mnt/d/xxx.xls')


if __name__ == "__main__":
    # loadfile2excel()

    loadjsonstr2excel()
