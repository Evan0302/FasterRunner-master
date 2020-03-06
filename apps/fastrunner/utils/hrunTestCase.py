import json

import xlrd


from fastrunner.utils.hrunutils import checkStr2Int_Bool, strStrip, json2File
from fastrunner.utils.parser import Format


class HrunTestCase:
    """
    config
    testcase

    from excel yaml json
    to excel json
    """

    def __init__(self, tc_content=None):
        self.ObjTestcase = tc_content

    def toJsonString(self):
        jstr = json.dumps(self.ObjTestcase, skipkeys=False, check_circular=True, allow_nan=True,
                          indent=4, separators=(',', ':'), ensure_ascii=False)
        return jstr

    def toJsonFile(self, outjsonfile):
        json2File(outjsonfile, self.ObjTestcase)

    def fromJson2Excel(self, sheet):
        pass

    def fromSheet2List(self, sheet):

        # 获取总行数
        rosw = sheet.nrows
        tmplist = []

        tcDict = {}
        # 根据总行数进行读取
        iRow: int = 0
        while (iRow < rosw):
            # 过滤空行
            if sheet.row_values(iRow, 0, 1)[0] != '' or sheet.row_values(iRow, 1, 2)[0] != '':
                tmpdict = {}
                # 解析config
                if sheet.row_values(iRow, 0, 1)[0] == 'config':
                    tcDict['config'] = self.parserConfig2Dict(sheet, iRow)
                    # tmplist.append(self.parserConfig(sheet, i))

                    iRow += 2
                elif sheet.row_values(iRow, 0, 1)[0] in ['teststeps','api']:
                    tcDict['teststeps'] = self.parserTestcase2List(sheet, iRow, sheet.row_values(iRow, 0, 1)[0])
                    # tmplist.extend(self.parserTestcase(sheet, i, sheet.row_values(i, 0, 1)[0]))
                    # tmplist.append(self.parserTestcase(sheet, i, sheet.row_values(i, 0, 1)[0]))
                    break
                elif sheet.row_values(iRow, 0, 1)[0] in ['testcases', 'apis', 'tests']:
                    tmplist.extend(self.parserTestcase2List(sheet, iRow, sheet.row_values(iRow, 0, 1)[0]))
                    # tmplist.append(self.parserTestcase(sheet, i, sheet.row_values(i, 0, 1)[0]))
                    break
                elif sheet.row_values(iRow, 0, 1)[0] == '=':
                    tmplist.extend(self.parserTestcase2List(sheet, iRow - 1, 'testcases'))
                    # tmplist.append(self.parserTestcase(sheet, i - 1, 'testcases'))
                    break
                else:
                    rowvalues = sheet.row_values(iRow)
                    # 讲每一行的内容添加进去
                    tmplist.append(rowvalues)
                    #     去除大标题第一行进行切割处理
            iRow += 1

        tmplist.append(tcDict)
        self.ObjTestcase = tmplist
        return tmplist

    def parserConfig2Dict(self, sheet, irow):
        dict_field = {}

        listFields = sheet.row_values(irow + 1)
        listValues = sheet.row_values(irow + 2)
        # 校验字段
        if listFields[0] == '=':
            dict_field = self._parserFields2Dict(listFields, listValues)

        else:
            print('字段行必须以 = 开头')
        return dict_field

    def parserTestcase2List(self, sheet, irow, name='testcases'):
        list_tc = []
        listFields = sheet.row_values(irow + 1)
        # 校验字段
        if listFields[0] == '=':
            for xrow in range(irow + 2, sheet.nrows):
                if sheet.row_values(xrow, 1, 2)[0] != '':
                    list_tc.append(self._parserFields2Dict(listFields, sheet.row_values(xrow)))
                    # list_tc.append({name:self.parserFields(listFields, sheet.row_values(xrow))})

        else:
            print('字段行必须以 = 开头')
        return list_tc

    def _parserFields2Dict(self, listfields, listvalues):
        dict_fields = {}

        def mergeRequest(sfield, svalue, dict_request={}):
            # 合并 request字段到 dict_request
            dict_request[sfield.split('.')[1]] = svalue
            return dict_request

        for icol in range(1, len(listfields)):  # 遍历字段
            xfield = listfields[icol]  # 字段名
            svalue = listvalues[icol]  # 字段值

            if svalue == '':  # 如果为空行即跳过
                pass
            else:
                # field 需要解析value
                if xfield in ['variables', 'parameters',
                              'request.headers', 'request.json', 'request.params', 'request.data']:
                    svalue = self.parserValueToDict(svalue)
                elif xfield in ['extract', 'validate']:
                    # list格式类型
                    svalue = self.parserValue2List(svalue)

                if xfield.startswith('request'):
                    # 合并request字段为 dict
                    dict_fields['request'] = mergeRequest(xfield, svalue)
                else:
                    dict_fields[xfield] = svalue
        return dict_fields

    def parserValue2List(self, svalue):
        valist = []
        tmplist = svalue.splitlines()
        i = 0
        while (i < len(tmplist)):
            s = tmplist[i]
            skey, svalue = self.parserValue2KeyValue(s.strip())

            if svalue == '':  # 处理多行value
                pass
            else:
                valist.append({skey: self.parserKvalue(svalue)})

            i += 1

        return valist

    def parserMultiLineToList(self, slist, index=0):
        tlist = []
        for i in range(index, len(slist)):
            tlist.append(self.parserKvalue(strStrip(slist[i])))
        return tlist

    def parserKvalue(self, svalue):
        if svalue.startswith('['):
            svalue = svalue.replace('[', '').replace(']', '').strip()
            xlist = svalue.split(',', 1)
            # tlist =[str(self.strip(s)) for s in xlist]
            # for s in xlist:
            #     tlist.append(self.strip(s) )

            return [checkStr2Int_Bool(s) for s in xlist]
        else:
            return checkStr2Int_Bool(svalue)

    def parserValueToDict(self, svalue):
        tmpdict = {}
        tmplist = svalue.splitlines()

        i = 0
        while (i < len(tmplist)):
            s = tmplist[i]
            skey, svalue = self.parserValue2KeyValue(s.strip())

            if svalue == '':  # 处理多行value
                tmpdict[skey.strip()] = self.parserMultiLineToList(tmplist, i + 1)
                break
            else:
                tmpdict[skey.strip()] = self.parserKvalue(svalue)
            i += 1

        return tmpdict

    def parserValue2KeyValue(self, svalue):
        """
        转化 key:value 为两个变量
        :param svalue:
        :return:  返回key、value 两个变量参数
        """
        xkeyvalue = svalue.split(':', 1)
        xkey = strStrip(xkeyvalue[0])
        return xkey, xkeyvalue[1].strip() if len(xkeyvalue) == 2 else ''



def addapi(apis, nodeid,projectid):
    from fastrunner.utils.hruntestcasefilter import fineFastTCBody
    # from fastrunner import models

    for xapi in apis['teststeps']:

        api = Format(fineFastTCBody(xapi))
        api.parse()

        api_body = {
            'name': api.name,
            'body': api.testcase,
            'url': api.url,
            'method': api.method,
            # 'project': models.Project.objects.get(id=projectid),
            'relation': nodeid
        }

        print(api_body)

        # models.API.objects.create(**api_body)


    return

def loadxls():
    wb = xlrd.open_workbook("/mnt/d/demo_httprunner.xls")
    jj = HrunTestCase().fromSheet2List(wb.sheet_by_name('API_list'))
    print(jj)
    addapi(jj[0], 4, 2)

if __name__ == "__main__":
    loadxls()