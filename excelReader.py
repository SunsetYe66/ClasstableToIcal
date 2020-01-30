# coding: utf-8
# Author: SunsetYe inherited from ChanJH
# Website: ChanJH <chanjh.com>, SunsetYe <github.com/sunsetye66>
# Contact: SunsetYe <sunsetye@me.com>
# todo:直接使用dict完成
# class info: classInfo: "" 时间的json
# design: class[0] = {className:"", startWeek:"", endWeek:"", weekStatus:,
# weekday:"", classTimeId:, classroom:"", teacher:""}

import sys
import json
import xlrd
import os
from random import randint


class ExcelReader:
    def __init__(self):
        # 读取时间表
        with open("conf_classTime.json", 'r', encoding='UTF-8') as f:
            self.class_timetable = json.loads(f.read())["classTime"]
            f.close()

        # 指定信息在 xls 表格内的列数
        self.config = dict()
        self.config["ClassName"] = 0
        self.config["StartWeek"] = 1
        self.config["EndWeek"] = 2
        self.config["Weekday"] = 3
        self.config["ClassTime"] = 4
        self.config["Classroom"] = 5
        self.config["WeekStatus"] = 6
        self.config["isClassSerialEnabled"] = [1, 7]
        self.config["isClassTeacherEnabled"] = [1, 8]
        # 信息列表 design: class[0] = {className:"", startWeek:"", endWeek:"", weekStatus:,
        # weekday:"", classTimeId:, classroom:"", teacher:""}
        # weekStatus: 0=Disabled 1=odd weeks 单周 2=even weeks 双周
        # 读取 excel 文件
        self.data = xlrd.open_workbook('classInfo.xlsx')
        self.table = self.data.sheets()[0]
        # 基础信息
        self.numOfRow = self.table.nrows  # 获取行数,即课程数
        self.numOfCol = self.table.ncols  # 获取列数,即信息量
        self.classList = list()

    def confirm_conf(self):
        # 与用户确定配置内容
        print("\n欢迎使用课程表生成工具·Excel 解析器。\n若自行修改过 Excel 表格结构，请检查。")
        print("ClassName: ", self.config["ClassName"])
        print("StartWeek: ", self.config["StartWeek"])
        print("EndWeek: ", self.config["EndWeek"])
        print("Weekday: ", self.config["Weekday"])
        print("ClassTime: ", self.config["ClassTime"])
        print("Classroom: ", self.config["Classroom"])
        print("WeekStatus: ", self.config["WeekStatus"])
        
        print("isClassSerialEnabled: ", self.config["isClassSerialEnabled"][0], end="\t")
        if self.config["isClassSerialEnabled"][0]:
            print("Serial: ", self.config["isClassSerialEnabled"][1])
        
        print("isClassTeacherEnabled: ", self.config["isClassTeacherEnabled"][0], end="\t")
        if self.config["isClassTeacherEnabled"][0]:
            print("Teacher: ", self.config["isClassTeacherEnabled"][1])

        option = input("回车继续，输入其他内容退出：")
        if option:
            return 1
        else:
            return 0

    def load_data(self):
        i = 1
        while i < self.numOfRow:
            _i = i - 1
            self.classList.append(dict())
            self.classList[_i].setdefault("ClassName", self.table.cell(i, self.config["ClassName"]).value)
            self.classList[_i].setdefault("StartWeek", self.table.cell(i, self.config["StartWeek"]).value)
            self.classList[_i].setdefault("EndWeek", self.table.cell(i, self.config["EndWeek"]).value)
            self.classList[_i].setdefault("WeekStatus", self.table.cell(i, self.config["WeekStatus"]).value)
            self.classList[_i].setdefault("Weekday", self.table.cell(i, self.config["Weekday"]).value)
            self.classList[_i].setdefault("ClassTimeId", self.table.cell(i, self.config["ClassTime"]).value)
            self.classList[_i].setdefault("Classroom", self.table.cell(i, self.config["Classroom"]).value)

            if self.config["isClassSerialEnabled"][0]:
                self.classList[_i].setdefault("ClassSerial",
                                              self.table.cell(i, self.config["isClassSerialEnabled"][1]).value)

            if self.config["isClassTeacherEnabled"][0]:
                self.classList[_i].setdefault("Teacher",
                                              self.table.cell(i, self.config["isClassTeacherEnabled"][1]).value)

            i += 1

    def write_data_test(self):
        if os.path.exists("conf_classInfo.json"):
            print("JSON File exists, use random filename.")
            filename = "conf_classInfo_" + str(randint(100, 999)) + ".json"
        else:
            filename = "conf_classInfo.json"
        with open(filename, 'w') as json_file:
            json_str = json.dumps(self.classList, ensure_ascii=False, indent=4)
            json_file.write(json_str)
            json_file.close()

    def main(self):
        if self.confirm_conf():
            sys.exit()
        self.load_data()
        self.write_data_test()


if __name__ == "__main__":
    p = ExcelReader()
    p.main()
