import json


def list2Dict(ovalue):
    xDict = {}
    for o in ovalue:
        # tt = ','.join([k + ': ' + str(v) for k, v in o.items()])
        for k, v in o.items():
            xDict[k] = v
    return xDict


ofunc = {}


def filterFastrun(oContent, dpath):
    """
    改善Json结构，兼容httprunner2 ,输出到testcase run
    """

    def _filter(odict):
        oconfig = odict['config']
        # pop config.desc
        odict['config'].pop('desc')
        # get function
        ofunc = oconfig['refs']['debugtalk']['functions']
        # pop config.refs
        odict['config'].pop('refs')

        # request.base_url to key path
        orequest = oconfig.pop('request')
        for k, v in orequest.items():
            oconfig[k] = v
            # if k=='base_url':
            #     oconfig['path'] = v

        # config.variables to dict
        ov = odict['config']['variables']
        odict['config']['variables'] = list2Dict(ov)

        for tc in odict['teststeps']:
            if 'desc' in tc:
                # pop teststeps.desc
                tc.pop('desc')
            if 'variables' in tc:
                # test.variables to dict
                tc['variables'] = list2Dict(tc['variables'])
        # ots = odict.pop('teststeps')
        # odict['test']=ots

        return odict, ofunc

    if isinstance(oContent, dict):
        oContent, ofunc = _filter(oContent)
    elif isinstance(oContent, list):
        for odict in oContent:
            odict, ofunc = _filter(odict)

    return {"project_mapping": {
        "PWD": dpath,
        "functions": ofunc,
        "env": {}
    }, 'testcases': oContent
    }


def fineToExcel(tcString):
    """
    优化转换到excel的case ，
            提取自DB api 输出到excel
    """
    oj = eval(tcString)
    oj.pop('desc')
    oj.pop('times')
    return [{'test': oj}]


def fineFastTCBody(tccontent):
    """
    搞定 由excel导入的json，输出到DB保存 api
    """

    def handle_desc(sname, onode):
        """
        处理desc字段
        （extract，variables）list
        (params, data) dict
        """
        ovar = onode.pop(sname) if sname in onode else []
        olist = []
        odesc = {}
        if isinstance(ovar, list):
            olist = ovar
            for o in ovar:
                for k, v in o.items():
                    odesc[k] = ""
        else:  # 转换Dict to list
            for k, v in ovar.items():
                olist.append({k: v})
                odesc[k] = ""
        onode[sname] = {sname: olist, 'desc': odesc}
        return onode[sname]

    def handleDict2Desc(sname, onode):
        odata = onode.pop(sname) if sname in onode else {}
        odesc = {}
        for k, v in odata.items():
            odesc[k] = ''
        onode[sname] = {sname: odata, 'desc': odesc}
        return onode[sname]

    tccontent['url'] = tccontent['request']['url']
    tccontent['method'] = tccontent['request']['method']
    tccontent['times'] = 1

    tccontent['header'] = {}
    tccontent['header']['header'] = tccontent['request'].pop('headers') if 'headers' in tccontent['request'] else {}
    tccontent['header']['desc'] = {}
    for k, v in tccontent['header']['header'].items():
        tccontent['header']['desc'][k] = ''

    handle_desc('extract', tccontent)

    tccontent['validate'] = {'validate': tccontent.pop('validate') if 'validate' in tccontent else {}}

    handleDict2Desc('params', tccontent['request'])

    tccontent['request']['form'] = handleDict2Desc('data', tccontent['request'])
    # tccontent['request']['params'] = handleDict2Desc('params')
    # tccontent['request']['form'] = handleDict2Desc('data')
    tccontent['request']['json'] = tccontent['request'].pop('json') if 'json' in tccontent['request'] else {}
    tccontent['request']['files'] = {'files': {}, 'desc': {}}

    handle_desc('variables', tccontent)

    tccontent['hooks'] = {"setup_hooks": [], "teardown_hooks": []}

    # tccontent['desc']= {'header': {'Content-Type': '', 'User-Agent': '', 'app_version': '', 'device_sn': '', 'os_platform': ''},
    #          'data': {}, 'files': {}, 'params': {}, 'variables': {}, 'extract': {'token': ''}}
    return tccontent


from httprunner.parser import LazyString, LazyFunction
from datetime import date
from datetime import datetime


def objdump2json(json_data):
    class JsonCustomEncoder(json.JSONEncoder):

        def default(self, field):

            if isinstance(field, datetime):
                return field.strftime('%Y-%m-%d %H:%M:%S')
            elif isinstance(field, date):
                return field.strftime('%Y-%m-%d')
            elif isinstance(field, LazyString):
                ov = field.to_value(ofunc)
                # print(ov)
                return ov
                # return "LazyString"
            else:
                return json.JSONEncoder.default(self, field)

    class PythonObjectEncoder(json.JSONEncoder):
        def default(self, obj):
            if isinstance(obj, LazyString):
                return "LazyString"
            try:
                return super().default(self, obj)
            except TypeError:
                return str(obj)

    print(json_data)

    return json.dumps(json_data,
                      ensure_ascii=False, skipkeys=False, check_circular=True, allow_nan=True,
                      cls=JsonCustomEncoder
                      # PythonObjectEncoder
                      )

    #                         indent=4,
    # separators=(',', ':'),


def fastreportPretty(osummary):
    """
    修正报告summary结果，兼容Httprunner2，保存到DB report
    """
    # modify stat.testsRun  <--total
    try:
        ostat = osummary['stat']
        ostat['testsRun'] = ostat['teststeps']['total']
        ostat['successes'] = ostat['teststeps']['successes']
        ostat['failures'] = ostat['teststeps']['failures']
        ostat['errors'] = ostat['teststeps']['errors']
        ostat['skipped'] = ostat['teststeps']['skipped']
    except:
        pass

    # meta_datas.data  to  meta_datas
    for odetails in osummary['details']:
        for orecord in odetails['records']:
            try:
                orecord['start_timestamp'] = odetails['time']['start_at']

                odatas = orecord['meta_datas']['data'][0]
                orecord['meta_datas']['request'] = odatas['request']
                orecord['meta_datas']['response'] = odatas['response']

                # record.meta_data.request.start_timestamp
                orecord['meta_datas']['request']['start_timestamp'] = odetails['time']['start_at']

                # meta_data.response.elapsed_ms <-- meta_datas.stat.elapsed_ms
                orecord['meta_datas']['response']['elapsed_ms'] = orecord['meta_datas']['stat']['elapsed_ms']
                # meta_data.response.response_time_ms <-- meta_datas.stat.response_time_ms            
                orecord['meta_datas']['response']['response_time_ms'] = orecord['meta_datas']['stat'][
                    'response_time_ms']
                # meta_data.response.content_size <-- meta_datas.stat.content_size            
                orecord['meta_datas']['response']['content_size'] = orecord['meta_datas']['stat']['content_size']

                # status <-- response.status_code
                orecord['status'] = orecord['meta_datas']['response']['status_code']

                # validator in record.meta_data.validators
                orecord['meta_datas']['validators'] = orecord['meta_datas']['validators'].pop('validate_extractor')

                # meta_datas --> meta_data
                orecord['meta_data'] = orecord['meta_datas']

                # oexpand_data = orecord['meta_datas_expanded'].pop('data')[0]
                # orecord['meta_datas_expanded']['request'] = odatas['request']
                # orecord['meta_datas_expanded']['response'] = odatas['response']
            except:
                pass

        # 转换 LazyString 节点
        for k, v in odetails['in_out']['in'].items():
            if isinstance(v, LazyString):
                odetails['in_out']['in'][k] = v.to_value()
        for k, v in odetails['in_out']['out'].items():
            if isinstance(v, LazyString):
                odetails['in_out']['out'][k] = v.to_value()

    return osummary
