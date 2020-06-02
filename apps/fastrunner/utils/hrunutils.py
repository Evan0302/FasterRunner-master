import io
import json
import os


class HrunCall():
    """
      封装 Httprunner 的调用

      0、读取数据（excel、json or 其他）
      1、生成Testcase对象
      2、封装Hrun的对象
      3、调用Hrun测试方法
      4、输出报告
    """

    def __init__(self, workdir='') -> None:
        super().__init__()
        self.testcaseObj = None

        self.workDir = os.path.dirname(os.path.abspath(__file__)) if workdir == '' else workdir

    def callHrun(self):
        from httprunner.api import HttpRunner
        from httprunner import report
        # from httprunner.loader import buildup

        # load_debugtalk_functions()

        runner = HttpRunner(failfast=False, save_tests=True, log_level="debug")
        #     log_file="test.log"
        # )

        tcStructure = self._genHttpRunner()
        print(tcStructure)

        # 运行单个结构体
        summary = runner.run(tcStructure)

        print(summary)

        report_path = report.gen_html_report(
            summary,
            # report_template="/path/to/custom_report_template",
            report_dir=os.path.join(self.workDir, 'reports')
            # report_file="/path/to/report_file_path"
        )

    def _genHttpRunner(self):
        from httprunner import runner
        from httprunner.loader.buildup import load_debugtalk_functions

        return {"project_mapping": {
            "PWD": self.workDir,
            "functions": load_debugtalk_functions(),
            "env": {}
        }, 'testcases': self.testcaseObj
        }


def load_debugtalk_functions1():
    """ load project debugtalk.py module functions
        debugtalk.py should be located in project working directory.

    Returns:
        dict: debugtalk module functions mapping
            {
                "func1_name": func1,
                "func2_name": func2
            }

    """
    # load debugtalk.py module
    import importlib
    imported_module = importlib.import_module("debugtalk")
    from httprunner.loader.load import load_module_functions
    return load_module_functions(imported_module)


def checkStr2Int_Bool(svalue):
    """
    检查字符串是否 数字、Boolean类型
    :param svalue:
    :return:
    """
    svalue = strStrip(svalue)

    if svalue.isnumeric():
        return int(svalue)
    elif svalue.lower() in ['true', 'false']:
        return True if svalue.lower() == 'true' else False
    else:
        return svalue


def strStrip(instr):
    """
    字符串首尾 去除空格、'、"、-
    :param instr:
    :return:
    """
    instr = instr.strip()
    if instr.startswith('-'):  # 去除 -
        instr = instr.replace('-', '', 1).strip()
    instr = instr.strip("'").strip()  # 去除 " ' 引号

    return instr


def json2File(outfile, obj):
    # jstr = json.dumps(obj, skipkeys=False, check_circular=True, allow_nan=True,
    #                   indent=4, separators=(',', ':'), ensure_ascii=False)
    with io.open(outfile, 'w', encoding='utf-8') as out:
        json.dump(obj, out, indent=4, separators=(',', ': '), ensure_ascii=False)
        out.write('\n')
    # print(jstr)


def loadJson(filename):
    with io.open(filename, 'r', encoding='utf-8') as jfile:
        return json.load(jfile)
