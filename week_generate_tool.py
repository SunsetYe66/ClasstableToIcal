# coding: utf-8
# Author: SunsetYe
# Website: github.com/sunsetye66
# Contact: sunsetye@me.com
# To generate the Ordinal of Weeks 生成周数
import sys
from datetime import datetime, timedelta
from uuid import uuid4 as uid
from random import randint
import socket

# 下方为需要修改的参数
first_week = "20200224"  # 第一周的周一；若你习惯使用周日做一周的第一天，可改为第一周的周日
total_weeks = 18  # 总周数
g_color = "#68b88e"  # 预览时的颜色，可不改动

# 下方参数若要改动，请清楚自己在做什么
f_random = randint(10, 99)
g_name = f'{datetime.now().strftime("%Y.%m")} 周数指示器@{socket.gethostname()}'
first_week_obj = datetime.strptime(first_week, "%Y%m%d")
utc_now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
delta_7 = timedelta(days=7)

# 写入文件头
try:
    with open(f"weeks-{str(f_random)}.ics", "w", encoding='UTF-8') as f:  # 追加要a
        ical_begin_base = f'''BEGIN:VCALENDAR
        VERSION:2.0
        X-WR-CALNAME:{g_name}
        X-APPLE-CALENDAR-COLOR:{g_color}
        '''
        f.write(ical_begin_base)
        f.close()
except:
    print("写入失败！可能是没有权限，请重试。")
    sys.exit()
else:
    print("文件头写入成功！")

for i in range(total_weeks):
    curr_week = i + 1
    curr_week_obj = first_week_obj + i * delta_7  # 当前周第一天的 obj
    end_date_obj = curr_week_obj + delta_7  # 当前周最后一天的 obj
    # 转换为ical所用的字符串
    begin_date = curr_week_obj.strftime("%Y%m%d")
    end_date = end_date_obj.strftime("%Y%m%d")

    # 构造 ical 文件体
    _ical_base = f'''
BEGIN:VEVENT\nCREATED:{utc_now}\nDTSTAMP:{utc_now}\nTZID:Asia/Shanghai\nSEQUENCE:0
SUMMARY:第 {curr_week} 周\nDTSTART;VALUE=DATE:{begin_date}\nDTEND;VALUE=DATE:{end_date}\nUID:{uid()}
END:VEVENT\n'''

    # 写入当前段 VEVENT
    with open(f"weeks-{str(f_random)}.ics", "a", encoding='UTF-8') as f:
        f.write(_ical_base)
        print(f"第{curr_week}周写入成功！")
        f.close()

# 写入尾部
with open(f"weeks-{str(f_random)}.ics", "a", encoding='UTF-8') as f:
    ical_end_base = "\nEND:VCALENDAR"
    f.write(ical_end_base)
    print(f"尾部信息写入成功！")
    f.close()

# 输出文件名作为提示
final_inform = f'最终文件 weeks-{str(f_random)}.ics 已生成，可通过内网传输到 iOS Device 上使用，参见主文件。'
print(final_inform)
