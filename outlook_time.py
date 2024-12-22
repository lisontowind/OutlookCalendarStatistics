import datetime as dt
import win32com.client
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import math
import pytz

def get_years():
    """
    获取年份列表，从 2000 到 2030。
    Returns:
        list[str]: ['2000', '2001', ..., '2030']
    """
    return [str(year) for year in range(2000, 2031)]


def is_leap_year(year):
    """
    判断是否为闰年。原始反编译代码有问题，这里做了典型的闰年判断修正:
    - 若年份能被 400 整除，则是闰年；
    - 否则若能被 100 整除，则不是闰年；
    - 否则若能被 4 整除，则是闰年；
    - 否则不是闰年。
    """
    if year % 400 == 0:
        return True
    elif year % 100 == 0:
        return False
    elif year % 4 == 0:
        return True
    else:
        return False


def get_days(year, month):
    """
    根据给定的年、月，返回该月天数。
    """
    # 针对不同月份返回天数
    if month in (1, 3, 5, 7, 8, 10, 12):
        return 31
    elif month in (4, 6, 9, 11):
        return 30
    elif month == 2:
        # 判断是否是闰年
        return 29 if is_leap_year(year) else 28
    else:
        # 如果给出的 month 无效，返回 0
        return 0


def update_days(event=None):
    """
    当开始日期的年份或月份改变时，更新“开始日”下拉列表。
    若当前选择的日大于该月天数，则自动将其重置为最后一天。
    """
    month = int(month_combobox.get())
    year = int(year_combobox.get())
    days = get_days(year, month)
    
    # 更新“日”下拉框的可选值
    day_combobox['values'] = [f"{day:02d}" for day in range(1, days + 1)]
    
    # 如果当前选中的日大于该月应有的天数，则重置为该月的最后一天
    current_day = int(day_combobox.get())
    if current_day > days:
        day_combobox.set(f"{days:02d}")


def update_days_end(event=None):
    """
    当结束日期的年份或月份改变时，更新“结束日”下拉列表。
    若当前选择的日大于该月天数，则自动将其重置为最后一天。
    """
    month = int(end_month_combobox.get())
    year = int(end_year_combobox.get())
    days = get_days(year, month)
    
    # 更新“日”下拉框的可选值
    end_day_combobox['values'] = [f"{day:02d}" for day in range(1, days + 1)]
    
    current_day = int(end_day_combobox.get())
    if current_day > days:
        end_day_combobox.set(f"{days:02d}")


# 获取 Outlook 日历数据
def get_calendar(begin, end):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        restriction = (
            "[END] >= '"
            + begin.strftime("%Y-%m-%d %H:%M")
            + "' AND [Start] <= '"
            + end.strftime("%Y-%m-%d %H:%M")
            + "'"
        )

        calendar = outlook.getDefaultFolder(9).Items
        calendar.IncludeRecurrences = True
        calendar.Sort("[Start]")
        calendar = calendar.Restrict(restriction)

        return calendar

    except Exception as e:
        messagebox.showerror("错误", f"无法读取Outlook日历数据：{str(e)}")
        return None

def get_appointments(calendar, begin, end):
    """
    从给定的 calendar 对象里取出在 [begin, end] 交集部分的日程，返回列表。
    每一个日程包含字段: start, end, duration(分钟), categories, subject
    注意：本函数需要传入正确的 begin, end，且 calendar 必须是可遍历的 Outlook 约会项列表。
    """
    appointments = []
    for app in calendar:
        # 将 Outlook 的 start、end 转换为带 pytz.UTC 时区的 datetime，并截断到 [begin, end] 范围
        app_start = app.start.replace(tzinfo=pytz.UTC)
        app_end = app.end.replace(tzinfo=pytz.UTC)
        app_start = max(app_start, begin)
        app_end = min(app_end, end)

        # 如果有效时间段大于 0，则加入统计
        if app_end > app_start:
            appointments.append({
                'start': app_start.strftime('%Y-%m-%d %H:%M:%S'),
                'end': app_end.strftime('%Y-%m-%d %H:%M:%S'),
                'duration': (app_end - app_start).total_seconds() / 60,  # 单位: 分钟
                'categories': app.categories,
                'subject': app.subject
            })
    return appointments


def process_data():
    """
    该函数原本打算:
      1) 读取用户在界面上选择的起止日期与时间 -> 生成 begin, end (datetime)
      2) 通过 get_calendar 获取该区间内的 Outlook 日历项
      3) 调用 get_appointments(calendar, begin, end) 获取最终的约会列表
      4) 将 appointments 存为全局变量，以备画图函数使用

    目前逻辑还未完整实现，仅保留空壳。
    可在此处进行时间合法性校验、报错提示等。
    """
    global begin, end, appointments
    
    # 例如从界面获取开始、结束的年、月、日、时、分
    year_s = int(year_combobox.get())
    month_s = int(month_combobox.get())
    day_s = int(day_combobox.get())
    hour_s = int(start_hour.get())
    minute_s = int(start_minute.get())
    
    year_e = int(end_year_combobox.get())
    month_e = int(end_month_combobox.get())
    day_e = int(end_day_combobox.get())
    hour_e = int(end_hour_combobox.get())
    minute_e = int(end_minute_combobox.get())

    # 组装为 datetime 对象
    begin = dt.datetime(year_s, month_s, day_s, hour_s, minute_s, tzinfo=pytz.UTC)
    end = dt.datetime(year_e, month_e, day_e, hour_e, minute_e, tzinfo=pytz.UTC)

    # 如果时间非法（结束时间早于开始时间），可以提示错误并返回 False
    if end <= begin:
        error_label.config(text="错误：结束时间不能早于开始时间！", foreground="red")
        return False
    else:
        error_label.config(text="", foreground="black")

    # 这里调用 get_calendar(begin, end)，获取 Outlook 日历对象
    calendar = get_calendar(begin, end)  
    
    # 调用 get_appointments(calendar, begin, end) 返回有效约会
    appointments = get_appointments(calendar, begin, end)

    # 若要测试绘图效果，可以手动模拟一部分 appointments，例如:
    # appointments = [
    #     {'start': '2024-01-01 09:00:00', 'end': '2024-01-01 10:00:00', 'duration': 60, 'categories': '工作', 'subject': '晨会'},
    #     {'start': '2024-01-01 14:00:00', 'end': '2024-01-01 15:30:00', 'duration': 90, 'categories': '休息, 阅读', 'subject': '午后读书'},
    #     ...
    # ]

    return True  # 若成功处理，则返回 True


def show_pie_chart():
    """
    显示“时间分配”的饼图。
    如果没有正确处理数据 (process_data 返回 False)，则不绘制。
    """
    global canvas

    # 先进行数据处理
    if not process_data():
        return
    
    # 如果已有画布对象，先销毁以免叠加
    if canvas is not None:
        canvas.destroy()

    # 计算总的分钟数
    total_duration = (end - begin).total_seconds() / 60

    # 用于统计不同类别总时长(分钟)
    category_duration = {}
    categorized_duration = 0

    # 遍历所有日程，累加各个类别的时长
    for app in appointments:
        # 如果 categories 为空，则视为 '未分类'
        categories = app['categories'].split(', ') if app['categories'] else ['未分类']
        for category in categories:
            category_duration[category] = category_duration.get(category, 0) + app['duration']
        categorized_duration += app['duration']

    # 统计“不在任何已知日程类别中的时间”（即未统计时间）
    no_stat_duration = total_duration - categorized_duration

    # 读取“是否隐藏未统计时间”的复选框状态
    hide_unstat = hide_unstat_var.get()
    if hide_unstat:
        # 如果隐藏未统计时间，则饼图只画有日程的部分
        total_duration = categorized_duration
    else:
        # 如果不隐藏且有未统计时间，加入一个“未统计”类别
        if no_stat_duration > 0:
            category_duration['未统计'] = no_stat_duration

    # 初始化画布
    canvas = tk.Canvas(root, width=400, height=400)
    canvas.grid(row=7, column=0, columnspan=2)

    # 饼图从 0 度开始画
    start_angle = 0

    # 定义一些常见颜色，绘制的饼图会按顺序循环使用
    colors = [
        'lightblue', 'lightgreen', 'lightcoral', 'lightsalmon',
        'lightpink', 'lightyellow', 'lightgrey'
    ]
    
    text_positions = []  # 用于存储文本标签的位置和内容
    i = 0               # 用于取颜色的索引

    # 遍历每个类别，绘制扇区
    for category, duration in category_duration.items():
        extent = 360 * (duration / total_duration) if total_duration > 0 else 0
        # 如果计算角度正好是 360，减去一点点，避免画不出饼图
        extent = 359.99 if abs(extent - 360) < 1e-5 else extent

        # 选定一种颜色
        color = colors[i % len(colors)]

        # 画圆弧：(50,50,350,350) 表示外接正方形的左上和右下坐标
        canvas.create_arc(
            50, 50, 350, 350,
            start=start_angle,
            extent=extent,
            fill=color,
            outline='black'
        )

        # 计算扇区中心角度，用于放置文本
        mid_angle = start_angle + extent / 2
        # 将极坐标转换为直角坐标
        x = 200 + 130 * math.cos(math.radians(mid_angle))
        y = 200 - 130 * math.sin(math.radians(mid_angle))

        # 计算该类别占总时长的百分比
        percentage = f"{(duration / total_duration) * 100:.1f}%" if total_duration > 0 else "0.0%"
        
        # 将“类别 + 百分比”文字及其坐标先存起来
        text_positions.append((x, y, f"{category} {percentage}"))
        
        # 下次扇区的起始角度就是当前扇区终止角度
        start_angle += extent
        i += 1

    # 将所有文本绘制到画布上
    for (x, y, text) in text_positions:
        canvas.create_text(x, y, text=text, font=('Arial', 10))

    # 在顶部放置一个标题
    canvas.create_text(200, 20, text='时间分配饼图', font=('Arial', 14))


def show_bar_chart():
    """
    显示“时间分配”的条形图。
    如果没有正确处理数据 (process_data 返回 False)，则不绘制。
    """
    global canvas

    # 先进行数据处理
    if not process_data():
        return
    
    # 如果已有画布对象，先销毁以免叠加
    if canvas is not None:
        canvas.destroy()

    # 计算总分钟数
    total_minutes = (end - begin).total_seconds() / 60

    # 统计不同类别的时长(转换为小时)
    category_duration = {}
    categorized_minutes = 0

    for app in appointments:
        categories = app['categories'].split(', ') if app['categories'] else ['未分类']
        for category in categories:
            # 注意，这里将 duration(分钟) / 60 变为小时
            category_duration[category] = category_duration.get(category, 0) + app['duration'] / 60
        categorized_minutes += app['duration']

    # 计算 “未统计” 时间（小时）
    no_stat_hours = (total_minutes - categorized_minutes) / 60

    # 检查是否隐藏“未统计”时间
    hide_unstat = hide_unstat_var.get()
    if hide_unstat:
        total_hours = categorized_minutes / 60
    else:
        total_hours = total_minutes / 60
        if no_stat_hours > 0:
            category_duration['未统计'] = no_stat_hours

    # 计算类别数量
    num_categories = len(category_duration)
    
    # 画布高度可根据类别数量伸展，保证能够容纳所有条形
    canvas_height = max(400, 50 + num_categories * 50)

    # 初始化画布
    canvas = tk.Canvas(root, width=400, height=canvas_height)
    canvas.grid(row=7, column=0, columnspan=2)

    # 找到最大的时长，用于确定条形最大宽度
    if len(category_duration) > 0:
        max_duration = max(category_duration.values())
    else:
        max_duration = 1  # 如果没有任何类别，避免除零
    
    # 可用的垂直高度大约是 300，一行一个 bar
    bar_height = 300 / num_categories if num_categories > 0 else 0

    colors = [
        'lightblue', 'lightgreen', 'lightcoral', 'lightsalmon',
        'lightpink', 'lightyellow', 'lightgrey'
    ]

    text_positions = []
    i = 0

    # 逐个类别画矩形条
    for category, duration in category_duration.items():
        # 根据与最大值的比率来决定条形的宽度(最多画到335)
        bar_length = 335 * (duration / max_duration) if max_duration != 0 else 0
        
        # 选择颜色
        color = colors[i % len(colors)]

        # 画一个矩形条形
        top_y = 50 + i * bar_height
        bottom_y = 50 + (i + 1) * bar_height
        canvas.create_rectangle(
            65, top_y, 65 + bar_length, bottom_y,
            fill=color
        )

        # 类别名显示在条形左侧
        text_positions.append((60, top_y + bar_height / 2, category, 'e'))

        # 时长数值显示在条形中间（可根据情况微调坐标）
        text_positions.append((65 + bar_length / 2, top_y + bar_height / 2, f"{duration:.2f} h", 'w'))

        i += 1

    # 将文字绘制到画布上
    for x, y, text, anchor in text_positions:
        canvas.create_text(x, y, text=text, anchor=anchor)

    # 在顶部放置一个标题
    canvas.create_text(200, 20, text='时间分配条形图', font=('Arial', 14))


if __name__ == '__main__':
    # 一些全局变量
    appointments = None      # 用于存放最终得到的日程项目
    begin = None             # 开始时间
    end = None               # 结束时间
    canvas = None            # 全局画布对象

    # 创建主窗口
    root = tk.Tk()
    root.title('选择起止日期和时间')
    root.geometry('450x550')

    today = dt.date.today()

    # -------------- 1) 日期时间输入区域 --------------
    input_frame = ttk.Frame(root)
    input_frame.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
    input_frame.grid_columnconfigure(1, weight=1)

    # ------------------ 开始日期 ------------------
    ttk.Label(input_frame, text='开始日期:').grid(row=0, column=0, padx=5, pady=5, sticky='e')
    year_combobox = ttk.Combobox(input_frame, values=get_years(), width=5)
    year_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    year_combobox.set(str(today.year))
    year_combobox.bind('<<ComboboxSelected>>', update_days)

    month_combobox = ttk.Combobox(
        input_frame,
        values=[f"{i:02d}" for i in range(1, 13)],
        width=3
    )
    month_combobox.grid(row=0, column=2, padx=2, sticky='ew')
    month_combobox.set(f"{today.month:02d}")
    month_combobox.bind('<<ComboboxSelected>>', update_days)

    day_combobox = ttk.Combobox(input_frame, width=3)
    day_combobox.grid(row=0, column=3, padx=2, sticky='ew')
    day_combobox.set(f"{today.day:02d}")
    # 手动更新一次，以保证“日”下拉可用
    update_days()

    start_hour = ttk.Combobox(input_frame, values=[f"{i:02d}" for i in range(24)], width=3)
    start_hour.grid(row=0, column=4, padx=2, sticky='ew')
    start_hour.set('00')

    start_minute = ttk.Combobox(input_frame, values=[f"{i:02d}" for i in range(60)], width=3)
    start_minute.grid(row=0, column=5, padx=2, sticky='ew')
    start_minute.set('00')

    # ------------------ 结束日期 ------------------
    ttk.Label(input_frame, text='结束日期:').grid(row=1, column=0, padx=5, pady=5, sticky='e')
    end_year_combobox = ttk.Combobox(input_frame, values=get_years(), width=5)
    end_year_combobox.grid(row=1, column=1, padx=5, pady=5, sticky='w')
    end_year_combobox.set(str(today.year))
    end_year_combobox.bind('<<ComboboxSelected>>', update_days_end)

    end_month_combobox = ttk.Combobox(
        input_frame,
        values=[f"{i:02d}" for i in range(1, 13)],
        width=3
    )
    end_month_combobox.grid(row=1, column=2, padx=2, sticky='ew')
    end_month_combobox.set(f"{today.month:02d}")
    end_month_combobox.bind('<<ComboboxSelected>>', update_days_end)

    end_day_combobox = ttk.Combobox(input_frame, width=3)
    end_day_combobox.grid(row=1, column=3, padx=2, sticky='ew')
    end_day_combobox.set(f"{today.day:02d}")
    update_days_end()

    # 多绑定一次以防止在某些情况下失效
    end_month_combobox.bind('<<ComboboxSelected>>', update_days_end)
    end_year_combobox.bind('<<ComboboxSelected>>', update_days_end)

    end_hour_combobox = ttk.Combobox(input_frame, values=[f"{i:02d}" for i in range(24)], width=3)
    end_hour_combobox.grid(row=1, column=4, padx=2, sticky='ew')
    end_hour_combobox.set('23')

    end_minute_combobox = ttk.Combobox(input_frame, values=[f"{i:02d}" for i in range(60)], width=3)
    end_minute_combobox.grid(row=1, column=5, padx=2, sticky='ew')
    end_minute_combobox.set('59')

    # 错误信息展示
    error_label = ttk.Label(input_frame, text="")
    error_label.grid(row=2, column=0, columnspan=6, sticky='ew')

    # “隐藏未统计时间” 的复选框
    hide_unstat_var = tk.IntVar()
    hide_unstat_checkbutton = ttk.Checkbutton(
        input_frame,
        text='隐藏未统计时间',
        variable=hide_unstat_var
    )
    hide_unstat_checkbutton.grid(row=3, column=0, columnspan=6, padx=5, pady=5, sticky='ew')

    # -------------- 2) 显示图表按钮 --------------
    button_frame = ttk.Frame(root)
    button_frame.grid(row=1, column=0, padx=5, pady=5, sticky='ew')
    button_frame.grid_columnconfigure(0, weight=1)
    button_frame.grid_columnconfigure(1, weight=1)

    ttk.Button(button_frame, text='显示饼图', command=show_pie_chart).grid(row=0, column=0, padx=5, pady=5, sticky='ew')
    ttk.Button(button_frame, text='显示条形图', command=show_bar_chart).grid(row=0, column=1, padx=5, pady=5, sticky='ew')

    # -------------- 3) 放置图表的画布区域 --------------
    canvas_frame = ttk.Frame(root)
    canvas_frame.grid(row=2, column=0, padx=5, pady=5, sticky='nsew')
    canvas_frame.grid_propagate(False)
    canvas_frame.config(width=400, height=450)

    canvas = tk.Canvas(canvas_frame, width=400, height=450)
    canvas.pack()

    # 使第三行可以伸缩
    root.grid_rowconfigure(2, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # 处理关闭窗口的方式
    def on_closing():
        """
        在关闭窗口前可以进行一些清理或询问操作。
        这里简单地直接退出。
        """
        root.quit()
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', on_closing)

    # 进入主事件循环
    root.mainloop()
