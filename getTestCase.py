#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import multiprocessing
import xlsxwriter
import os
import json
import time
from threading import Thread
import wx
from multiprocessing import Pool


ip_server = "172.31.2.125"


def get_detail(id_sub, login_session):
    headers_data = {
        'Accept': "*/*",
        'Accept-Encoding': "gzip, deflate",
        'Accept-Language': "zh-CN,zh;q=0.9",
        'Connection': "keep-alive",
        'Content-Length': "32",
        'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
        'Host': "{}".format(ip_server),
        'Origin': "http://{}".format(ip_server),
        'Referer': "http://{}/iauto_acp/login.html".format(ip_server),
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
        'X-Requested-With': "XMLHttpRequest",
    }
    url_data = "http://{}/iauto_acp/itmsTestCaseN.do/getBaselineTestCaseByFirstLevel?firstLevel=2".format(ip_server)
    payload_data_sub = "id={}".format(id_sub)
    get_page = login_session.post(url_data, headers=headers_data, data=payload_data_sub)
    data_page = json.loads(get_page.text)
    print("Get detail info for id:{} with return code {}".format(id_sub, get_page.status_code))
    case_parentId_sub = []
    case_id_sub = []
    case_testCaseName_sub = []
    case_demandPoint_sub = []
    case_testCaseNumber_sub = []
    case_testPrecondition_sub = []
    case_testConfigLimit_sub = []
    case_testAttention_sub = []
    case_SOPRefer_sub = []
    case_testDataDemand_sub = []
    case_testProcedure_sub = []
    case_testExpect_sub = []
    case_projectProgress_sub = []
    case_testCaseVersion_sub = []
    case_testCaseVersionChange_sub = []
    case_testCaseAuthor_sub = []
    case_testCaseOwner_sub = []
    case_testCaseDemandNumber_sub = []
    case_testCaseDemandDescription_sub = []
    case_testLevelTestType_sub = []
    case_testCaseStatus_sub = []
    case_productionLimit_sub = []
    case_ifMultiExecution_sub = []
    case_multiExecutionDescription_sub = []
    case_manualTestTime_sub = []
    case_automatedTestTime_sub = []
    case_osLimit_sub = []
    case_osAndHardwareRelevance_sub = []
    case_ifAutomatedTest_sub = []
    case_nonautomatedCause_sub = []
    case_automatedNumber_sub = []
    case_ifReserveRecord_sub = []
    case_ifReserveLog_sub = []
    case_note_sub = []
    if len(data_page) != 0:
        for item_data in data_page:
            parentId = item_data["parentId"]
            testCaseId = item_data["id"]
            testCaseName = item_data["testCaseName"]
            demandPoint = item_data["demandPoint"]
            testCaseNumber = item_data["testCaseNumber"]
            testPrecondition = item_data["testPrecondition"]
            testConfigLimit = item_data["testConfigLimit"]
            testAttention = item_data["testAttention"]
            SOPRefer = item_data["SOPRefer"]
            testDataDemand = item_data["testDataDemand"]
            testProcedure = item_data["testProcedure"]
            testExpect = item_data["testExpect"]
            projectProgress = item_data["projectProgress"]
            testCaseVersion = item_data["testCaseVersion"]
            testCaseVersionChange = item_data["testCaseVersionChange"]
            testCaseAuthor = item_data["testCaseAuthor"]
            testCaseOwner = item_data["testCaseOwner"]
            testCaseDemandNumber = item_data["testCaseDemandNumber"]
            testCaseDemandDescription = item_data["testCaseDemandDescription"]
            testLevelTestType = item_data["testLevelTestType"]
            testCaseStatus = item_data["testCaseStatus"]
            productionLimit = item_data["productionLimit"]
            ifMultiExecution = item_data["ifMultiExecution"]
            multiExecutionDescription = item_data["multiExecutionDescription"]
            manualTestTime = item_data["manualTestTime"]
            automatedTestTime = item_data["automatedTestTime"]
            osLimit = item_data["osLimit"]
            osAndHardwareRelevance = item_data["osAndHardwareRelevance"]
            ifAutomatedTest = item_data["ifAutomatedTest"]
            nonautomatedCause = item_data["nonautomatedCause"]
            automatedNumber = item_data["automatedNumber"]
            ifReserveRecord = item_data["ifReserveRecord"]
            ifReserveLog = item_data["ifReserveLog"]
            note = item_data["note"]

            case_parentId_sub.append(parentId)
            case_id_sub.append(testCaseId)
            case_testCaseName_sub.append(testCaseName)
            case_demandPoint_sub.append(demandPoint)
            case_testCaseNumber_sub.append(testCaseNumber)
            case_testPrecondition_sub.append(testPrecondition)
            case_testConfigLimit_sub.append(testConfigLimit)
            case_testAttention_sub.append(testAttention)
            case_SOPRefer_sub.append(SOPRefer)
            case_testDataDemand_sub.append(testDataDemand)
            case_testProcedure_sub.append(testProcedure)
            case_testExpect_sub.append(testExpect)
            case_projectProgress_sub.append(projectProgress)
            case_testCaseVersion_sub.append(testCaseVersion)
            case_testCaseVersionChange_sub.append(testCaseVersionChange)
            case_testCaseAuthor_sub.append(testCaseAuthor)
            case_testCaseOwner_sub.append(testCaseOwner)
            case_testCaseDemandNumber_sub.append(testCaseDemandNumber)
            case_testCaseDemandDescription_sub.append(testCaseDemandDescription)
            case_testLevelTestType_sub.append(testLevelTestType)
            case_testCaseStatus_sub.append(testCaseStatus)
            case_productionLimit_sub.append(productionLimit)
            case_ifMultiExecution_sub.append(ifMultiExecution)
            case_multiExecutionDescription_sub.append(multiExecutionDescription)
            case_manualTestTime_sub.append(manualTestTime)
            case_automatedTestTime_sub.append(automatedTestTime)
            case_osLimit_sub.append(osLimit)
            case_osAndHardwareRelevance_sub.append(osAndHardwareRelevance)
            case_ifAutomatedTest_sub.append(ifAutomatedTest)
            case_nonautomatedCause_sub.append(nonautomatedCause)
            case_automatedNumber_sub.append(automatedNumber)
            case_ifReserveRecord_sub.append(ifReserveRecord)
            case_ifReserveLog_sub.append(ifReserveLog)
            case_note_sub.append(note)

    return case_parentId_sub, case_id_sub, case_testCaseName_sub, case_demandPoint_sub, case_testCaseNumber_sub, case_testPrecondition_sub, case_testConfigLimit_sub, case_testAttention_sub, case_SOPRefer_sub, case_testDataDemand_sub, case_testProcedure_sub, case_testExpect_sub, case_projectProgress_sub, case_testCaseVersion_sub, case_testCaseVersionChange_sub, case_testCaseAuthor_sub, case_testCaseOwner_sub, case_testCaseDemandNumber_sub, case_testCaseDemandDescription_sub, case_testLevelTestType_sub, case_testCaseStatus_sub, case_productionLimit_sub, case_ifMultiExecution_sub, case_multiExecutionDescription_sub, case_manualTestTime_sub, case_automatedTestTime_sub, case_osLimit_sub, case_osAndHardwareRelevance_sub, case_ifAutomatedTest_sub, case_nonautomatedCause_sub, case_automatedNumber_sub, case_ifReserveRecord_sub, case_ifReserveLog_sub, case_note_sub


class getTestCaseFrame(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TestCase信息抓取工具", pos=wx.DefaultPosition,
                          size=wx.Size(504, 471), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.m_panel1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_title1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"请在如下输入用户名和密码", wx.DefaultPosition, wx.DefaultSize,
                                         wx.ST_NO_AUTORESIZE)
        self.text_title1.Wrap(-1)

        self.text_title1.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_title1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer3, 0, wx.EXPAND, 5)

        gSizer2 = wx.GridSizer(2, 2, 0, 0)

        self.text_username = wx.StaticText(self.m_panel1, wx.ID_ANY, u"用户名", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_username.Wrap(-1)

        self.text_username.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_username.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_username = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        gSizer2.Add(self.input_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.text_password = wx.StaticText(self.m_panel1, wx.ID_ANY, u"密码", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_password.Wrap(-1)

        self.text_password.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_password.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_password = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        gSizer2.Add(self.input_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(gSizer2, 0, 0, 5)

        gSizer3 = wx.GridSizer(0, 2, 0, 0)

        bSizer10.Add(gSizer3, 0, 0, 5)

        bSizer21 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self.m_panel1, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer21.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self.m_panel1, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer21.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer10.Add(bSizer21, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.textctrl_display = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer91.Add(self.textctrl_display, 1, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer91, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer10)
        self.m_panel1.Layout()
        bSizer10.Fit(self.m_panel1)
        bSizer2.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    def newthread(self):
        Thread(target=self.run_all).start()

    def run_all(self):
        self.button_go.Disable()
        self.updatedisplay("开始抓取，请耐心等待...")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        # 获取用户名和密码
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()

        url_login = "http://{}/iauto_acp/login".format(ip_server)
        login_session = requests.session()
        headers_login = {
            'Accept': "*/*",
            'Accept-Encoding': "gzip, deflate",
            'Accept-Language': "zh-CN,zh;q=0.9",
            'Connection': "keep-alive",
            'Content-Length': "32",
            'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
            'Host': "{}".format(ip_server),
            'Origin': "http://{}".format(ip_server),
            'Referer': "http://{}/iauto_acp/login.html".format(ip_server),
            'User-Agent': "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36",
            'X-Requested-With': "XMLHttpRequest",
        }
        payload_login = "username={}&password={}".format(username, password)
        # 使用以上获取的信息post登录
        login_session.post(url_login, headers=headers_login, data=payload_login)

        # 开始获取数据
        data_id_name = {}
        data = {"0": {}}
        level_zero = ["0"]
        level_one = []
        level_two = []
        level_three = []
        level_four = []
        url_data = "http://{}/iauto_acp/itmsTestCaseN.do/getBaselineTestCaseByFirstLevel?firstLevel=2".format(ip_server)
        # get level one info, sample: SIT/BIOS/BMC
        for item_lv1_top in level_zero:
            payload_data_lv1 = "id={}".format(item_lv1_top)
            response_data_lv1 = json.loads(
                login_session.post(url_data, headers=headers_login, data=payload_data_lv1).text)
            for item_lv1 in response_data_lv1:
                id_lv1 = item_lv1["id"]
                name_lv1 = item_lv1["levelName"]
                parent_lv1 = item_lv1["parentId"]
                data_id_name["{}".format(id_lv1)] = name_lv1
                level_one.append(id_lv1)
                data["{}".format(parent_lv1)]["{}".format(id_lv1)] = {}

        # get level two info, examples: function/performance
        for item_lv2_top in level_one:
            payload_data_lv2 = "id={}".format(item_lv2_top)
            response_data_lv2 = json.loads(
                login_session.post(url_data, headers=headers_login, data=payload_data_lv2).text)
            if len(response_data_lv2) != 0:
                for item_lv2 in response_data_lv2:
                    id_lv2 = item_lv2["id"]
                    name_lv2 = item_lv2["levelName"]
                    parent_lv2 = item_lv2["parentId"]

                    data_id_name["{}".format(id_lv2)] = name_lv2
                    level_two.append(id_lv2)
                    for item_0 in data:
                        for item_1 in data["{}".format(item_0)]:
                            if str(item_1) == str(parent_lv2):
                                data["{}".format(item_0)]["{}".format(item_1)]["{}".format(id_lv2)] = {}

        # get level three info, example:cpu/mem/hdd
        for item_lv3_top in level_two:
            payload_data_lv3 = "id={}".format(item_lv3_top)
            response_data_lv3 = json.loads(
                login_session.post(url_data, headers=headers_login, data=payload_data_lv3).text)
            if len(response_data_lv3) != 0:
                for item_lv3 in response_data_lv3:
                    id_lv3 = item_lv3["id"]
                    name_lv3 = item_lv3["levelName"]
                    parent_lv3 = item_lv3["parentId"]

                    data_id_name["{}".format(id_lv3)] = name_lv3
                    level_three.append(id_lv3)
                    for item_0 in data:
                        for item_1 in data["{}".format(item_0)]:
                            for item_2 in data["{}".format(item_0)]["{}".format(item_1)]:
                                if str(item_2) == str(parent_lv3):
                                    data["{}".format(item_0)]["{}".format(item_1)]["{}".format(item_2)][
                                        "{}".format(id_lv3)] = {}
        # get level four info, maybe detail testcase
        for item_lv4_top in level_three:
            payload_data_lv4 = "id={}".format(item_lv4_top)
            response_data_lv4 = json.loads(
                login_session.post(url_data, headers=headers_login, data=payload_data_lv4).text)
            if len(response_data_lv4) != 0:
                if "testProcedure" not in response_data_lv4[0]:  # still level, not detail testcase
                    for item_lv4 in response_data_lv4:
                        id_lv4 = item_lv4["id"]
                        name_lv4 = item_lv4["levelName"]
                        parent_lv4 = item_lv4["parentId"]

                        data_id_name["{}".format(id_lv4)] = name_lv4
                        level_four.append(id_lv4)

                        for item_0 in data:
                            for item_1 in data["{}".format(item_0)]:
                                for item_2 in data["{}".format(item_0)]["{}".format(item_1)]:
                                    for item_3 in data["{}".format(item_0)]["{}".format(item_1)]["{}".format(item_2)]:
                                        if str(item_3) == str(parent_lv4):
                                            data["{}".format(item_0)]["{}".format(item_1)]["{}".format(item_2)][
                                                "{}".format(item_3)]["{}".format(id_lv4)] = {}
                else:
                    if item_lv4_top not in level_four:
                        level_four.append(item_lv4_top)

        # get detail test case
        case_lv1 = []
        case_lv2 = []
        case_lv3 = []
        case_lv4 = []
        case_parentId = []
        case_id = []
        case_testCaseName = []
        case_demandPoint = []
        case_testCaseNumber = []
        case_testPrecondition = []
        case_testConfigLimit = []
        case_testAttention = []
        case_SOPRefer = []
        case_testDataDemand = []
        case_testProcedure = []
        case_testExpect = []
        case_projectProgress = []
        case_testCaseVersion = []
        case_testCaseVersionChange = []
        case_testCaseAuthor = []
        case_testCaseOwner = []
        case_testCaseDemandNumber = []
        case_testCaseDemandDescription = []
        case_testLevelTestType = []
        case_testCaseStatus = []
        case_productionLimit = []
        case_ifMultiExecution = []
        case_multiExecutionDescription = []
        case_manualTestTime = []
        case_automatedTestTime = []
        case_osLimit = []
        case_osAndHardwareRelevance = []
        case_ifAutomatedTest = []
        case_nonautomatedCause = []
        case_automatedNumber = []
        case_ifReserveRecord = []
        case_ifReserveLog = []
        case_note = []

        temp_detail = []
        pool_detail = Pool(multiprocessing.cpu_count())
        for index_data_top, item_data_top in enumerate(level_four):
            temp_detail.append(pool_detail.apply_async(get_detail, args=(item_data_top, login_session)))
        pool_detail.close()
        pool_detail.join()
        '''
        return case_parentId_sub, case_id_sub, case_testCaseName_sub, case_demandPoint_sub,
        case_testCaseNumber_sub, case_testPrecondition_sub, case_testConfigLimit_sub, case_testAttention_sub,
        case_SOPRefer_sub, case_testDataDemand_sub, case_testProcedure_sub, case_testExpect_sub, 
        case_projectProgress_sub, case_testCaseVersion_sub, case_testCaseVersionChange_sub, case_testCaseAuthor_sub, 
        case_testCaseOwner_sub, case_testCaseDemandNumber_sub, case_testCaseDemandDescription_sub, case_testLevelTestType_sub,
        case_testCaseStatus_sub, case_productionLimit_sub, case_ifMultiExecution_sub, case_multiExecutionDescription_sub,
        case_manualTestTime_sub, case_automatedTestTime_sub, case_osLimit_sub, case_osAndHardwareRelevance_sub,
        case_ifAutomatedTest_sub, case_nonautomatedCause_sub, case_automatedNumber_sub, case_ifReserveRecord_sub,
        case_ifReserveLog_sub, case_note_sub
        '''
        for item_detail in temp_detail:
            data_detail_temp = item_detail.get()

            case_parentId.extend(data_detail_temp[0])
            case_id.extend(data_detail_temp[1])
            case_testCaseName.extend(data_detail_temp[2])
            case_demandPoint.extend(data_detail_temp[3])
            case_testCaseNumber.extend(data_detail_temp[4])
            case_testPrecondition.extend(data_detail_temp[5])
            case_testConfigLimit.extend(data_detail_temp[6])
            case_testAttention.extend(data_detail_temp[7])
            case_SOPRefer.extend(data_detail_temp[8])
            case_testDataDemand.extend(data_detail_temp[9])
            case_testProcedure.extend(data_detail_temp[10])
            case_testExpect.extend(data_detail_temp[11])
            case_projectProgress.extend(data_detail_temp[12])
            case_testCaseVersion.extend(data_detail_temp[13])
            case_testCaseVersionChange.extend(data_detail_temp[14])
            case_testCaseAuthor.extend(data_detail_temp[15])
            case_testCaseOwner.extend(data_detail_temp[16])
            case_testCaseDemandNumber.extend(data_detail_temp[17])
            case_testCaseDemandDescription.extend(data_detail_temp[18])
            case_testLevelTestType.extend(data_detail_temp[19])
            case_testCaseStatus.extend(data_detail_temp[20])
            case_productionLimit.extend(data_detail_temp[21])
            case_ifMultiExecution.extend(data_detail_temp[22])
            case_multiExecutionDescription.extend(data_detail_temp[23])
            case_manualTestTime.extend(data_detail_temp[24])
            case_automatedTestTime.extend(data_detail_temp[25])
            case_osLimit.extend(data_detail_temp[26])
            case_osAndHardwareRelevance.extend(data_detail_temp[27])
            case_ifAutomatedTest.extend(data_detail_temp[28])
            case_nonautomatedCause.extend(data_detail_temp[29])
            case_automatedNumber.extend(data_detail_temp[30])
            case_ifReserveRecord.extend(data_detail_temp[31])
            case_ifReserveLog.extend(data_detail_temp[32])
            case_note.extend(data_detail_temp[33])

        for index_parentid, item_parentid in enumerate(case_parentId):
            for item_0 in data:
                for item_1 in data["{}".format(item_0)]:
                    for item_2 in data["{}".format(item_0)]["{}".format(item_1)]:
                        for item_3 in data["{}".format(item_0)]["{}".format(item_1)]["{}".format(item_2)]:
                            if str(item_3) == str(item_parentid):
                                case_lv1.append(item_1)
                                case_lv2.append(item_2)
                                case_lv3.append(item_3)
                                case_lv4.append("None")
            for item_0 in data:
                for item_1 in data["{}".format(item_0)]:
                    for item_2 in data["{}".format(item_0)]["{}".format(item_1)]:
                        for item_3 in data["{}".format(item_0)]["{}".format(item_1)]["{}".format(item_2)]:
                            for item_4 in data["{}".format(item_0)]["{}".format(item_1)]["{}".format(item_2)]["{}".format(item_3)]:
                                if str(item_4) == str(item_parentid):
                                    case_lv1.append(item_1)
                                    case_lv2.append(item_2)
                                    case_lv3.append(item_3)
                                    case_lv4.append(item_4)

        # 如下是本地数据处理，与浏览器不再发生关系
        TitleItem = ['层级1', '层级2', '层级3', '层级4', '测试用例名称', '测试需求点', '测试用例编号', '测试准备-前提条件', '测试配置要求', '测试注意事项',
                     '参考SOP列表', '测试数据要求', '测试步骤', '预期结果', '项目进程', '测试用例版本', '版本变更记录', '用例作者', '用例归属', '用例设计需求编号',
                     '用例设计需求描述', '测试级别-测试类型', '用例状态', '适用产品', '是否重复执行', '重复执行描述', '手动测试时间', '自动测试时间', 'OS类别',
                     'OS与硬件相关系', '是否自动化用例', '不可自动化原因', '自动化编号', '是否保留系统&BMC日志', '是否保留测试数据/Log/截图', '备注']

        timestamp = time.strftime('%Y%m%d', time.localtime())
        WorkBook = xlsxwriter.Workbook("测试用例信息-{}.xlsx".format(timestamp))
        SheetOne = WorkBook.add_worksheet('测试用例')
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)
        # formatOne.set_bold(1)

        SheetOne.set_column('A:AJ', 8)
        SheetOne.set_column('C:D', 15)
        SheetOne.set_column('M:N', 40)

        # write row one
        SheetOne.merge_range(0, 0, 0, 3, "测试用例目录层级", formatOne)
        SheetOne.merge_range(0, 4, 0, 13, "测试执行", formatOne)
        SheetOne.merge_range(0, 14, 0, 18, "测试用例信息", formatOne)
        SheetOne.merge_range(0, 19, 0, 20, "需求跟踪", formatOne)
        SheetOne.merge_range(0, 21, 0, 27, "用例选择指导", formatOne)
        SheetOne.merge_range(0, 28, 0, 32, "自动化相关", formatOne)
        SheetOne.merge_range(0, 33, 0, 34, "日志", formatOne)
        SheetOne.write(0, 35, "备注", formatOne)

        already_write_id_list = []
        # write title
        for i in range(0, len(TitleItem)):
            SheetOne.write(1, i, TitleItem[i], formatOne)
        # write data

        for index_write, item_write in enumerate(case_id):
            if item_write not in already_write_id_list:
                already_write_id_list.append(item_write)

                SheetOne.write(2 + index_write, 0, data_id_name[case_lv1[index_write]], formatOne)
                SheetOne.write(2 + index_write, 1, data_id_name[case_lv2[index_write]], formatOne)
                SheetOne.write(2 + index_write, 2, data_id_name[case_lv3[index_write]], formatOne)
                if case_lv4[index_write] != "None":
                    SheetOne.write(2 + index_write, 3, data_id_name[case_lv4[index_write]], formatOne)
                SheetOne.write(2 + index_write, 4, case_testCaseName[index_write], formatOne)
                SheetOne.write(2 + index_write, 5, case_demandPoint[index_write], formatOne)
                SheetOne.write(2 + index_write, 6, case_testCaseNumber[index_write], formatOne)
                SheetOne.write(2 + index_write, 7, case_testPrecondition[index_write], formatOne)
                SheetOne.write(2 + index_write, 8, case_testConfigLimit[index_write], formatOne)
                SheetOne.write(2 + index_write, 9, case_testAttention[index_write], formatOne)
                SheetOne.write(2 + index_write, 10, case_SOPRefer[index_write], formatOne)
                SheetOne.write(2 + index_write, 11, case_testDataDemand[index_write], formatOne)
                SheetOne.write(2 + index_write, 12, case_testProcedure[index_write], formatOne)
                SheetOne.write(2 + index_write, 13, case_testExpect[index_write], formatOne)
                SheetOne.write(2 + index_write, 14, case_projectProgress[index_write], formatOne)
                SheetOne.write(2 + index_write, 15, case_testCaseVersion[index_write], formatOne)
                SheetOne.write(2 + index_write, 16, case_testCaseVersionChange[index_write], formatOne)
                SheetOne.write(2 + index_write, 17, case_testCaseAuthor[index_write], formatOne)
                SheetOne.write(2 + index_write, 18, case_testCaseOwner[index_write], formatOne)
                SheetOne.write(2 + index_write, 19, case_testCaseDemandNumber[index_write], formatOne)
                SheetOne.write(2 + index_write, 20, case_testCaseDemandDescription[index_write], formatOne)
                SheetOne.write(2 + index_write, 21, case_testLevelTestType[index_write], formatOne)
                SheetOne.write(2 + index_write, 22, case_testCaseStatus[index_write], formatOne)
                SheetOne.write(2 + index_write, 23, case_productionLimit[index_write], formatOne)
                SheetOne.write(2 + index_write, 24, case_ifMultiExecution[index_write], formatOne)
                SheetOne.write(2 + index_write, 25, case_multiExecutionDescription[index_write], formatOne)
                if len(case_manualTestTime[index_write]) == 0:
                    SheetOne.write(2 + index_write, 26, 0.0, formatOne)
                else:
                    try:
                        SheetOne.write(2 + index_write, 26, float(case_manualTestTime[index_write]), formatOne)
                    except (UnicodeEncodeError, ValueError):
                        SheetOne.write(2 + index_write, 26, 0.0, formatOne)

                if len(case_automatedTestTime[index_write]) == 0:
                    SheetOne.write(2 + index_write, 27, 0.0, formatOne)
                else:
                    try:
                        SheetOne.write(2 + index_write, 27, float(case_automatedTestTime[index_write]), formatOne)
                    except (UnicodeEncodeError, ValueError):
                        SheetOne.write(2 + index_write, 27, 0.0, formatOne)

                SheetOne.write(2 + index_write, 28, case_osLimit[index_write], formatOne)
                SheetOne.write(2 + index_write, 29, case_osAndHardwareRelevance[index_write], formatOne)
                SheetOne.write(2 + index_write, 30, case_ifAutomatedTest[index_write], formatOne)
                SheetOne.write(2 + index_write, 31, case_nonautomatedCause[index_write], formatOne)
                SheetOne.write(2 + index_write, 32, case_automatedNumber[index_write], formatOne)
                SheetOne.write(2 + index_write, 33, case_ifReserveLog[index_write], formatOne)
                SheetOne.write(2 + index_write, 34, case_ifReserveRecord[index_write], formatOne)
                SheetOne.write(2 + index_write, 35, case_note[index_write], formatOne)
        WorkBook.close()
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.updatedisplay(
            "抓到{}个结果！已经将结果写入《测试用例信息-{}.xlsx》，请自行查阅！请点击EXIT退出程序！".format(len(case_id), timestamp))
        time.sleep(1)
        self.updatedisplay("Finished")
        self.button_go.Enable()

    def close(self, event):
        self.Close()

    def onbutton(self, event):
        self.button_go.Disable()
        self.newthread()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("完成第{}页".format(t))
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText("{}".format(t))
        self.textctrl_display.AppendText(os.linesep)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = getTestCaseFrame(None)
    frame.Show()
    app.MainLoop()
