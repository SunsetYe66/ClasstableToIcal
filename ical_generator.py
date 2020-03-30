# coding: utf-8
# Author: SunsetYe inherited from ChanJH
# Website: ChanJH <chanjh.com>, SunsetYe <github.com/sunsetye66>
# Contact: SunsetYe <sunsetye@me.com>

import json
import sys
from datetime import datetime, timedelta
from uuid import uuid4 as uid
import socket


class GenerateCal:
    def __init__(self):
        # 定义全局参数
        self.first_week = "20200224"  # 第一周周一的日期
        self.inform_time = 25  # 提前 N 分钟提醒
        self.g_name = f'{datetime.now().strftime("%Y.%m")} 课程表@{socket.gethostname()}'  # 全局课程表名
        self.g_color = "#ff9500"  # 预览时的颜色（可以在 iOS 设备上修改）
        self.a_trigger = ""

        # 读取文件，返回 dict(class_timetable) 时间表
        try:
            with open("conf_classTime.json", 'r', encoding='UTF-8') as f:
                self.class_timetable = json.loads(f.read())
                f.close()
        except:
            print("时间配置文件 conf_classTime.json 似乎有点问题")
            sys.exit()
        # 读取文件，返回 dict(class_info) 课程信息
        try:
            with open("conf_classInfo.json", 'r', encoding='UTF-8') as f:
                self.class_info = json.loads(f.read())
                f.close()
        except:
            print("课程配置文件 conf_classInfo.json 似乎有点问题")
            sys.exit()

    # 获取内网ip用于最后的提示
    def get_host_ip(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(('119.29.29.29', 80))
            ip = s.getsockname()[0]
        finally:
            s.close()
        return ip

    def set_attribute(self):
        self.first_week = input("请输入第一周周一的日期，格式为 YYYYMMDD，如 20200224：")  # 第一周周一的日期
        c = 0
        while c == 0:
            self.inform_time = input("请输入提前提醒时间，以分钟计；若不需要提醒请输入 N：")
            try:
                self.inform_time = int(self.inform_time)  # 提前 N 分钟提醒
                if self.inform_time <= 60:
                    self.a_trigger = f'-P0DT0H{self.inform_time}M0S'
                elif 60 < self.inform_time <= 1440:
                    minutes = self.inform_time % 60
                    hours = self.inform_time // 60
                    self.a_trigger = f'-P0DT{hours}H{minutes}M0S'
                else:
                    minutes = self.inform_time % 60
                    hours = (self.inform_time // 60) - 24
                    days = self.inform_time // 1440
                    self.a_trigger = f'-P{days}DT{hours}H{minutes}M0S'
                c = 1
            except ValueError:
                if self.inform_time in "nN":
                    self.a_trigger = ""
                    c = 1
                else:
                    print("输入数字有误！")

    def main_process(self):
        utc_now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        weekdays = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]

        # 开始操作，先写入头
        ical_begin_base = f'''BEGIN:VCALENDAR
VERSION:2.0
X-WR-CALNAME:{self.g_name}
X-APPLE-CALENDAR-COLOR:{self.g_color}
X-WR-TIMEZONE:Asia/Shanghai
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
X-LIC-LOCATION:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
'''
        try:
            with open(f"res-{str(utc_now)}.ics", "w", encoding='UTF-8') as f:  # 追加要a
                f.write(ical_begin_base)
                f.close()
        except:
            print("写入失败！可能是没有权限，请重试。")
            sys.exit()
        else:
            print("文件头写入成功！")

        initial_time = datetime.strptime(self.first_week, "%Y%m%d")  # 将开始时间转换为时间对象
        i = 1
        for obj in self.class_info:
            # 计算课程第一次开始的日期 first_time_obj，公式：7*(开始周数-1) （//把第一周减掉） + 周几 - 1 （没有周0，等于把周一减掉）
            try:
                delta_time = 7 * (obj['StartWeek'] - 1) + obj['Weekday'] - 1
            except TypeError:
                print("请检查 Excel 中是否有无用行，并删除 conf_classInfo.json 后重新运行 Excel 读取器及 iCal 生成器！")
                sys.exit()

            if obj['WeekStatus'] == 1:  # 单周
                if obj["StartWeek"] % 2 == 0:  # 若单周就不变，双周加7
                    delta_time += 7
            elif obj['WeekStatus'] == 2:  # 双周
                if obj["StartWeek"] % 2 != 0:  # 若双周就不变，单周加7
                    delta_time += 7
            first_time_obj = initial_time + timedelta(days=delta_time)  # 处理完单双周之后 first_time_obj 就是真正开始的日期
            if obj["WeekStatus"] == 0:  # 处理隔周课程
                extra_status = "1"
            else:
                extra_status = f'2;BYDAY={weekdays[int(obj["Weekday"] - 1)]}'  # BYDAY 是周 N，隔周重复需要带上

            try:  # 尝试处理纯数字的课程序号
                obj["ClassSerial"] = str(int(obj["ClassSerial"]))
                serial = f'课程序号：{obj["ClassSerial"]}'
            except ValueError:
                obj["ClassSerial"] = str(obj["ClassSerial"])
                serial = f'课程序号：{obj["ClassSerial"]}'
            except KeyError:  # 如果没有这个 key，直接略过
                serial = ""

            # 计算课程第一次开始、结束的时间，后面使用RRule重复即可，格式类似 20200225T120000
            final_stime_str = first_time_obj.strftime("%Y%m%d") + "T" + \
                              self.class_timetable[str(int(obj['ClassTimeId']))]["startTime"]
            final_etime_str = first_time_obj.strftime("%Y%m%d") + "T" + \
                              self.class_timetable[str(int(obj['ClassTimeId']))]["endTime"]
            delta_week = 7 * int(obj["EndWeek"] - obj["StartWeek"])
            stop_time_obj = first_time_obj + timedelta(days=delta_week + 1)
            stop_time_str = stop_time_obj.strftime("%Y%m%dT%H%M%SZ")  # 注意是utc时间，直接+1天处理
            # 教师可选，在此做判断
            try:
                teacher = f'教师：{obj["Teacher"]}\t'
            except KeyError:
                teacher = ""

            # 生成此次循环的 event_base
            if self.a_trigger:
                _alarm_base = f'''BEGIN:VALARM\nACTION:DISPLAY\nDESCRIPTION:This is an event reminder
TRIGGER:{self.a_trigger}\nX-WR-ALARMUID:{uid()}\nUID:{uid()}\nEND:VALARM\n'''
            else:
                _alarm_base = ""
            _ical_base = f'''\nBEGIN:VEVENT
CREATED:{utc_now}\nDTSTAMP:{utc_now}\nSUMMARY:{obj["ClassName"]}
DESCRIPTION:{teacher}{serial}\nLOCATION:{obj["Classroom"]}
TZID:Asia/Shanghai\nSEQUENCE:0\nUID:{uid()}\nRRULE:FREQ=WEEKLY;UNTIL={stop_time_str};INTERVAL={extra_status}
DTSTART;TZID=Asia/Shanghai:{final_stime_str}\nDTEND;TZID=Asia/Shanghai:{final_etime_str}
X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC\n{_alarm_base}END:VEVENT\n'''

            # 写入文件
            with open(f"res-{str(utc_now)}.ics", "a", encoding='UTF-8') as f:
                f.write(_ical_base)
                print(f"第{i}条课程信息写入成功！")
                i += 1
                f.close()

        # 拼合头尾
        with open(f"res-{str(utc_now)}.ics", "a", encoding='UTF-8') as f:
            f.write("\nEND:VCALENDAR")
            print(f"尾部信息写入成功！")
            f.close()

        final_inform = f'''
        最终文件 res-{str(utc_now)}.ics 已生成，可通过内网传输到 iOS Device 上使用。
        方法：
        \t1. 在放置 ics 的目录下打开终端。
        \t2. 输入 python -m http.server 8000 或 python3 -m http.server 8000 搭建 HTTP 服务器
        \t3. 在 iOS Device 的 Safari 浏览器中输入：
        \t\t\t\t\t\thttp://{self.get_host_ip()}:8000/
        \t4. 点击 res-{str(utc_now)}.ics，选择导入日历即可。'''

        print(final_inform)


if __name__ == "__main__":
    p = GenerateCal()
    p.set_attribute()
    p.main_process()
