import xlrd
from django.http import HttpResponse
from rest_framework.decorators import api_view
from rest_framework.response import Response
from django.db import DataError
from fastrunner import models
from fastrunner.utils import response
from fastrunner.utils.decorator import request_log
from fastrunner.utils.hrunJsonYamlToExcel import Hrun_JsonYaml

from fastrunner.utils.hrunTestCase import HrunTestCase, HrunHarParser
from fastrunner.utils.hruntestcasefilter import fineToExcel
from fastrunner.utils.parser import Format

@api_view(['POST'])
@request_log(level='DEBUG')
def excel_export(request):
    '''
    用例导出为excel
    '''

    apilist = request.data['apilist']

    print("启动导出export===" + str(apilist))
    tclist = []
    for api in apilist:
        print(str(api['id'])+ '  ==  ' + api['name'])

        tc = models.API.objects.get(id=api['id'])
        new_body = fineToExcel(tc.body)

        tclist.extend(new_body)


    print(tclist)
    print('==================')
    hj = Hrun_JsonYaml(jsonObj=tclist, infilename='apiexport')
    wb= hj.toExcelWorkbook()
    wb.save('d:/xxx.xls')

    from io import BytesIO
    from PIL import Image
    f = BytesIO()
    wb.save(f)

    return HttpResponse(f.getvalue(), content_type="application/vnd.ms-excel")

    # return Response({
    #         'code': '0001',
    #         'success': True,
    #         'msg': 'excel文件导出成功'
    #     }
    #     )


@api_view(['POST'])
@request_log(level='DEBUG')
def file_import(request):
    '''
    用例导入
    token:
        验证token

    '''

    print("开始import")
    if request.method == 'POST':
        try:
            import os
            for k, f in request.FILES.items():
                paper_file = f
                nodeid = request.POST['nodeid']
                projectid = request.POST['projectid']

                fileext = os.path.splitext(paper_file.name)[1].lower()
                if fileext == '.xls':
                    # 判断文件是否excel
                    sheetname = request.POST['sheetname']
                    wb = xlrd.open_workbook(filename=None, file_contents=paper_file.read())

                    jj = HrunTestCase().fromSheet2List(wb.sheet_by_name(sheetname))

                    print(jj)
                    addapi(jj[0], nodeid,projectid )
                elif fileext == '.har':
                    # 处理har文件
                    file_contents = paper_file.read()
                    try:
                        har = HrunHarParser(file_contents)
                        jj = har.get_HrunTestcases()
                        print(jj)
                        addapi(jj, nodeid,projectid )

                    except (KeyError, TypeError):
                        return Response({
                                'code': '00322',
                                'success': False,
                                'msg': 'HAR file content error' + f.name
                            })
                else:
                    return Response({
                                'code': '00321',
                                'success': False,
                                'msg': '不支持的文件格式' + f.name
                            })


        except DataError:
            return Response(response.DATA_TO_LONG)

        return Response(response.API_ADD_SUCCESS)

    else:
        pass

    return Response(response.FILE_UPLOAD_SUCCESS)


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
            'project': models.Project.objects.get(id=projectid),
            'relation': nodeid
        }
        #
        # print(api_body)

        models.API.objects.create(**api_body)


    return

