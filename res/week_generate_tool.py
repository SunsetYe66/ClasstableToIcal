# coding: utf-8
# Author: SunsetYe
# Website: github.com/sunsetye66
# Contact: me # sunsetye.com
# Generate the Ordinal of Weeks 生成周数

import sys
from datetime import datetime, timedelta
from uuid import uuid4 as uid
from random import randint
import socket


class GenerateWeeks:
    def __init__(self):
        f_random = randint(10, 99)
        # 下方为需要修改的参数
        self.first_week = "20200224"  # 第一周的周一；若你习惯使用周日做一周的第一天，可改为第一周的周日
        self.total_weeks = 18  # 总周数
        self.g_color = "#68b88e"  # 预览时的颜色，可不改动
        self.indep_gen = 1  # 独立调用判断
        self.file_name = f"weeks-{str(f_random)}.ics"   # 非独立调用时传入的文件名

    def out_set_attribute(self, first_week, total_week,file_name):
        self.first_week = first_week
        self.total_weeks = total_week
        self.file_name = file_name
        self.indep_gen = 0

    def set_attribute(self):
        print("请输入第一周周一的日期，若使用周日为一周的第一天可输入第一周周日的日期。")
        self.first_week = input("格式为 YYYYMMDD，如 20200224：")
        self.total_weeks = int(input("请输入总周数："))

    def main_process(self):
        # 下方参数若要改动，请清楚自己在做什么
        g_name = f'{datetime.now().strftime("%Y.%m")} 周数指示器@{socket.gethostname()}'
        first_week_obj = datetime.strptime(self.first_week, "%Y%m%d")
        utc_now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        delta_7 = timedelta(days=7)

        # 写入文件头
        if self.indep_gen:
            try:
                with open(self.file_name, "w", encoding='UTF-8') as f:  # 追加要a
                    ical_begin_base = f'''BEGIN:VCALENDAR\nVERSION:2.0\nX-WR-CALNAME:{g_name}
    X-APPLE-CALENDAR-COLOR:{self.g_color}
                    '''
                    f.write(ical_begin_base)
                    f.close()
            except:
                print("写入失败！可能是没有权限，请重试。")
                sys.exit()
            else:
                print("文件头写入成功！")

        for i in range(self.total_weeks):
            curr_week = i + 1
            curr_week_obj = first_week_obj + i * delta_7  # 当前周第一天的 obj
            end_date_obj = curr_week_obj + delta_7  # 当前周最后一天的 obj
            # 转换为ical所用的字符串
            begin_date = curr_week_obj.strftime("%Y%m%d")
            end_date = end_date_obj.strftime("%Y%m%d")

            # 构造 ical 文件体
            _ical_base = f'''\nBEGIN:VEVENT
CREATED:{utc_now}\nDTSTAMP:{utc_now}\nTZID:Asia/Shanghai\nSEQUENCE:0
SUMMARY:第 {curr_week} 周\nDTSTART;VALUE=DATE:{begin_date}\nDTEND;VALUE=DATE:{end_date}\nUID:{uid()}
END:VEVENT\n'''

            # 写入当前段 VEVENT
            with open(self.file_name, "a", encoding='UTF-8') as f:
                f.write(_ical_base)
                print(f"第{curr_week}周写入成功！")
                f.close()

        # 写入尾部
        if self.indep_gen:
            with open(self.file_name, "a", encoding='UTF-8') as f:
                ical_end_base = "\nEND:VCALENDAR"
                f.write(ical_end_base)
                print(f"尾部信息写入成功！")
                f.close()

        # 输出文件名作为提示
        if self.indep_gen:
            final_inform = f'最终文件 {self.file_name} 已生成，可通过内网传输到 iOS Device 上使用，参见主文件。'
            print(final_inform)
        else:
            final_inform = f'周数指示器写入完成'
            print(final_inform)


if __name__ == '__main__':
    tool = GenerateWeeks()
    tool.set_attribute()
    tool.main_process()
