# coding: utf-8
# Author: SunsetYe
# Website: SunsetYe <github.com/sunsetye66>
# Contact: SunsetYe <sunsetye@me.com>
import sys


def gen_week():
    from week_generate_tool import GenerateWeeks
    process = GenerateWeeks()
    process.set_attribute()
    process.main_process()


def read_excel():
    from excel_reader import ExcelReader
    process = ExcelReader()
    process.main()


def gen_ical():
    from ical_generator import GenerateCal
    process = GenerateCal()
    process.set_attribute()
    process.main_process()


def main():
    inform_text = '''欢迎使用课程表生成工具！
使用前请参照 temp_classInfo.xlsx 内的格式、说明填写课程信息，并将其重命名为 classInfo.xlsx。
输入 1 进入「周数指示器」生成程序；
输入 2 进入「Excel 读取工具」；
输入 3 进入「iCal 生成工具」；
输入 0 退出，祝您使用愉快 ~
Copyright © 2020 Sunset Ye, Distributed under GPL Licence.'''
    print(inform_text)
    func = input("请输入要进入的功能：")
    if func == "0":
        sys.exit()
    elif func == "1":
        gen_week()
    elif func == "2":
        read_excel()
    elif func == "3":
        gen_ical()
    else:
        print("啥？")
        sys.exit()


if __name__ == '__main__':
    main()
    gen_week()
