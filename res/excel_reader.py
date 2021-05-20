# coding: utf-8
# Author: SunsetYe inherited from ChanJH
# Website: ChanJH <chanjh.com>, SunsetYe <github.com/sunsetye66>
# Contact: SunsetYe <me # sunsetye.com>
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
        # 指定信息在 xls 表格内的列数，第一列是第 0 列。
        self.config = dict()
        self.config["ClassName"] = 0
        self.config["StartWeek"] = 1
        self.config["EndWeek"] = 2
        self.config["Weekday"] = 3
        self.config["ClassStartTime"] = 4
        self.config["ClassEndTime"] = 5
        self.config["Classroom"] = 6
        self.config["WeekStatus"] = 7
        self.config["isClassSerialEnabled"] = [0, 8]
        self.config["isClassTeacherEnabled"] = [0, 9]
        # weekStatus: 0=Disabled 1=odd weeks 单周 2=even weeks 双周
        # 读取 excel 文件
        try:
            self.data = xlrd.open_workbook('./data/classInfo.xls')
        except FileNotFoundError:
            print("文件不存在，请确认是否将课程信息前的 temp_ 去掉！")
            sys.exit()
        self.table = self.data.sheets()[0]
        # 基础信息
        self.numOfRow = self.table.nrows  # 获取行数,即课程数
        self.numOfCol = self.table.ncols  # 获取列数,即信息量
        self.classList = list()

    def confirm_conf(self):
        # 与用户确定配置内容
        print("\n欢迎使用课程表生成工具·Excel 解析器。\n若自行修改过 Excel 表格结构，请检查。")
        print("若要设定是否使用单双周、是否显示教师，请修改 excel_reader.py 中的 27, 28 行。")
        print("ClassName: ", self.config["ClassName"])
        print("StartWeek: ", self.config["StartWeek"])
        print("EndWeek: ", self.config["EndWeek"])
        print("Weekday: ", self.config["Weekday"])
        print("ClassStartTime: ", self.config["ClassStartTime"])
        print("ClassEndTime: ", self.config["ClassEndTime"])
        print("Classroom: ", self.config["Classroom"])
        print("WeekStatus: ", self.config["WeekStatus"])

        print("isClassSerialEnabled: ", self.config["isClassSerialEnabled"][0], end="")
        if self.config["isClassSerialEnabled"][0]:
            print(" ,", "Serial: ", self.config["isClassSerialEnabled"][1])

        print("isClassTeachserEnabled: ", self.config["isClassTeacherEnabled"][0], end="")
        if self.config["isClassTeacherEnabled"][0]:
            print(" ,", "Teacher: ", self.config["isClassTeacherEnabled"][1])

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
            self.classList[_i].setdefault("ClassStartTimeId", self.table.cell(i, self.config["ClassStartTime"]).value)
            self.classList[_i].setdefault("ClassEndTimeId", self.table.cell(i, self.config["ClassEndTime"]).value)
            self.classList[_i].setdefault("Classroom", self.table.cell(i, self.config["Classroom"]).value)
            if self.config["isClassSerialEnabled"][0]:
                try:
                    self.classList[_i].setdefault("ClassSerial",
                                                  str(int(self.table.cell(
                                                      i, self.config["isClassSerialEnabled"][1]).value)))
                except ValueError:
                    self.classList[_i].setdefault("ClassSerial",
                                                  str(self.table.cell(i, self.config["isClassSerialEnabled"][1]).value))
            if self.config["isClassTeacherEnabled"][0]:
                self.classList[_i].setdefault("Teacher",
                                              self.table.cell(i, self.config["isClassTeacherEnabled"][1]).value)
            i += 1

    def write_data(self):
        if os.path.exists("./data/conf_classInfo.json"):
            print("已存在 JSON 文件，使用随机文件名，请手动修改！")
            filename = "./data/conf_classInfo_" + str(randint(100, 999)) + ".json"
        else:
            filename = "./data/conf_classInfo.json"
        with open(filename, 'w', encoding='UTF-8') as json_file:
            json_str = json.dumps(self.classList, ensure_ascii=False, indent=4)
            json_file.write(json_str)
            json_file.close()

    def main(self):
        if self.confirm_conf():
            sys.exit()
        self.load_data()
        self.write_data()
        print("Excel 文件读取成功！")


if __name__ == "__main__":
    p = ExcelReader()
    p.main()
    print(p.classList)
