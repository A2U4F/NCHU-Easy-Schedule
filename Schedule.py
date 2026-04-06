import re
import sys
import json
import datetime
from datetime import timedelta
from datetime import datetime

try:
    import xlrd
except ImportError:
    print("错误：缺少 xlrd 库。请运行: pip install xlrd==1.2.0")
    sys.exit(1)


class ClassCalendar:
    """
    解析课表上一个位置的课程\n
    name: 课程名称; teacher: 教师; position:上课地址;\n
    classtime: list[datetime] 所有上课时间;\n
    continuetime: 一次课程持续时间
    """
    def __init__(self, format_str: str, weekday: int, beginwith: datetime):
        """
        解析单元格文本生成课程对象
        :param format_str: 单元格原始文本（包含换行符的多行字符串）
        :param weekday: 星期几 (1-7, 1=周一)
        :param beginwith: 学期第一天日期 (datetime 对象，默认 8:00)
        """
        self.name: str = ""
        self.teacher: str = ""
        self.position: str = ""
        self.classtime: list[datetime] = []
        self.continuetime = 100  # 默认课程时长 100 分钟（2节小课）
        
        # 1. 解析“周次”块，例如: "\n1-4([周])" 或 "\n3([周])"
        pattern_week = re.compile(r'\n\S+\(\[周\]\)')
        match = pattern_week.search(format_str)
        
        if match:
            info = match.group(0)
            # 提取数字范围，例如 "1-4" 或 "3"
            pattern_digits = re.compile(r'\d+(-\d+)?')
            for digit_match in pattern_digits.finditer(info):
                k = digit_match.group(0)
                if '-' in k:
                    # 范围周次，如 1-4
                    parts = k.split('-')
                    beginning = int(parts[0]) - 1
                    ending = int(parts[1])
                    for i in range(beginning, ending):
                        # 计算具体日期：学期第一天 + (星期偏移 + 周数偏移)
                        delta_days = (weekday - 1) + (7 * i)
                        self.classtime.append(beginwith + timedelta(days=delta_days))
                else:
                    # 单周，如 3
                    a = int(k)
                    delta_days = (weekday - 1) + (7 * (a - 1))
                    self.classtime.append(beginwith + timedelta(days=delta_days))
            
            # 从字符串中移除已解析的周次信息
            format_str = format_str.replace(info, "", 1)
        
        # 2. 解析“节次”块，例如: "[01-02节]"
        pattern_section = re.compile(r'\[\S+节\]')
        match = pattern_section.search(format_str)
        
        if match:
            info = match.group(0)
            # 匹配开始节次的数字（如 "1-" 中的 1）
            match_start = re.search(r'\d-', info)
            if match_start:
                d = match_start.group(0)[0]  # 获取节次数字字符
                beginsec = 0
                
                # 节次到分钟的映射（基于校历）
                # 第1节(8:00)=0*60, 第3节(10:00)=2*60, 第5节(14:00)=6*60...
                time_map = {
                    '1': 0 * 60,    # 08:00
                    '3': 2 * 60,    # 10:00
                    '5': 6 * 60,    # 14:00
                    '7': 8 * 60,    # 16:00
                    '9': 11 * 60    # 19:00
                }
                beginsec = time_map.get(d, 0)
                
                # 为所有已记录的上课时间添加分钟偏移
                for i in range(len(self.classtime)):
                    self.classtime[i] += timedelta(minutes=beginsec)
                
                # 判断是否为两节连上（包含"-"），决定时长是 140 还是 100 分钟
                if re.search(r"-\d+-",info):
                    self.continuetime = 140
                else:
                    self.continuetime = 100
            
            format_str = format_str.replace(info, "", 1)
        
        # 3. 解析地点（最后一行），例如: "\n教室101\n"
        pattern_line = re.compile(r'\n(\S+)\n$')
        match = pattern_line.search(format_str)
        if match:
            self.position = match.group(1)
            format_str = format_str[:match.start()]  # 截掉尾部
        
        # 4. 解析教师（新的最后一行），例如: "\n张老师\n"
        pattern_line = re.compile(r'\n(\S+)$')
        match = pattern_line.search(format_str)
        if match:
            self.teacher = match.group(1)
            format_str = format_str[:match.start()]
        
        # 5. 剩余部分为课程名（去除首尾空白和换行）
        self.name = format_str.replace('\n','')

    def disp(self):
        """在控制台显示课程详情"""
        print(f"课程名: {self.name}")
        print(f"教室: {self.position}")
        print(f"教师: {self.teacher}")
        print("上课时间:")
        for i, t in enumerate(self.classtime, 1):
            print(f"  第{i:2d}次课: {t.strftime('%Y/%m/%d %H:%M')}")


class OneClass:
    """一节课的信息"""
    def __init__(self, Class: ClassCalendar, N: int):
        self.name: str = Class.name
        self.teacher: str = Class.teacher
        self.position: str = Class.position
        self.time: dict[str, datetime] = {
            "begin": Class.classtime[N],
            "end": Class.classtime[N] + timedelta(minutes = Class.continuetime)
        }


class CourseCalender:
    """
    所有的历程的类:\n
    Course: 课程名称为键, OneClass列表为值\n
    basedatetime: 学期的第一个时刻\n
    holiday: 放假和调休\n
    CreatingTime: 创建对象时间
    """

    def __init__(self, data: dict):
        """
        存储所有的历程
        :param data: 配置文件, 由.json文件生成
        """
        self.Course: dict[str,list[OneClass]] = {}
        """课程名称为键, OneClass列表为值"""
        
        try:
            y, m, d = data["firstday"]
            # 默认从 8:00 开始
            self.basedatetime = datetime(y, m, d, 8, 0, 0)
            """学期的第一个时刻"""
        except ValueError:
            print("错误：日期格式不正确")
            input("按回车键退出...")
            return
        
        self.holiday: dict = {}
        """放假和调休"""
        for days in data["Holiday"].values():
            for date in days["Holiday period"]:
                self.holiday[tuple(date)]="NoAdjustment"
            
            for date in days["Holiday adjustment"]:
                self.holiday[tuple(date[0])]="ThisDayIsReplaced"
                self.holiday[tuple(date[1])]=tuple(date[0])
        
        self.CreatingTime = datetime.now()
        """创建对象时间"""
    
    def addCourse(self, Class: ClassCalendar):
        name:str = Class.name
        if Class.name not in self.Course:
            self.Course[name]=[]
        
        for N in range(Class.classtime.__len__()):
            self.Course[name].append(OneClass(Class, N))
        
        OrderKey =  lambda C : int(C.time["begin"].strftime("%Y%m%d%H%M"))
        """input是OneClass类, 作用是把OneClass的开始时间变成数字"""
        self.Course[name].sort(key = OrderKey)
        """排序课程"""
        
        #融合连在一起的课程
        Step = 0
        while Step < (self.Course[name].__len__() - 1):
            preTeacher = self.Course[name][Step].teacher
            nextTeacher = self.Course[name][Step + 1].teacher
            prePosition = self.Course[name][Step].position
            nextPosition = self.Course[name][Step + 1].position
            if (preTeacher == nextTeacher) and (prePosition == nextPosition):
                preend = int(self.Course[name][Step].time["end"].strftime("%Y%m%d%H%M"))
                nextbegin = OrderKey(self.Course[name][Step+1])
                if nextbegin - preend <=1800:
                    self.Course[name][Step].time["end"] = self.Course[name].pop(Step+1).time["end"]
            
            Step+=1

    def moveHoliday(self):
        def GetDate(C: OneClass) -> tuple[int]:
            date = C.time["begin"]
            return (date.year, date.month, date.day)
        
        for name in list(self.Course.keys()):
            new_course_list = []
            for cls in self.Course[name]:
                date_tuple = GetDate(cls)
                Condition = self.holiday.get(date_tuple)
                
                if Condition is None:
                    new_course_list.append(cls)
                elif type(Condition) is tuple:
                    Cy, Cm, Cd = Condition
                    Py, Pm, Pd = date_tuple
                    Delta = datetime(Cy, Cm, Cd) - datetime(Py, Pm, Pd)
                    cls.time["begin"] += Delta
                    cls.time["end"] += Delta
                    new_course_list.append(cls)
                # Condition == "NoAdjustment" or "ThisDayIsReplaced"时直接丢弃（不加入新列表）
            
            self.Course[name] = new_course_list
    
    def icsGen(self) -> str:
        """生成 iCalendar 格式的 VEVENT 字符串块"""
        
        lines: list[str] = []
        # 格式化为 iCalendar 标准时间格式：20240219T080000
        for cou in self.Course:
            lines.append(f"#{cou}")

            for t in self.Course[cou]:
                dtstart = t.time["begin"].strftime("%Y%m%dT%H%M%S")
                dtend   = t.time["end"].strftime("%Y%m%dT%H%M%S")
                uid = f"{dtstart}@stu.nchu.edu.cn"
                uid += self.CreatingTime.strftime("%Y%m%dT%H%M%S")
                
                lines.extend([
                    "BEGIN:VEVENT",
                    f"DTSTART:{dtstart}",
                    f"DTEND:{dtend}",
                    f"SUMMARY:{t.name}",
                    f"DESCRIPTION:{t.teacher}",
                    f"LOCATION:{t.position}",
                    f"UID:{uid}",
                    "END:VEVENT"
                ])
        
        return "\n".join(lines)
    
    def ShowInfo(self) -> str:
        lines: str = ""
        
        for Step,name in enumerate(self.Course,1):
            lines += f"第{Step:2d}门课: {name}, 共{self.Course[name].__len__():2d}节课\n"
        
        return lines


def main():
    # 设置标准输入输出编码（Windows CMD 可能需要 chcp 65001）
    if sys.platform == 'win32':
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # 打开配置文件
    try:
        with open("Configuration.json",'r') as f:
            Configuration= json.load(f)
    except:
        print("打开配置文件出错:",sys.exc_info()[0])
        return
    
    # 设置文件路径
    #print("请将文件拖入此处并回车: ", end="", flush=True)
    #file_in_addr = input().strip().strip('"')  # 去除 Windows 拖入时可能产生的引号
    file_in_addr: str = Configuration["Path"]
    file_out_addr: str = file_in_addr + ".ics"
    
    # 打开 .xls 文件
    try:
        book = xlrd.open_workbook(file_in_addr)
        schedule = book.sheet_by_index(0)  # 获取第一个工作表
    except Exception as e:
        print(f"\n错误：无法打开文件 '{file_in_addr}'")
        print(f"详细信息: {e}")
        input("按回车键退出...")
        return
    

    # 设置总日历变量
    Cal = CourseCalender(Configuration)
    
    # 识别 .xls 文件内容并生成 .ics 文件
    try:
        with open(file_out_addr, 'w', encoding='utf-8') as f:
            f.write("BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n")
            
            step = 0
            # 匹配课程详情块的正则：4-5 个换行开头的非空白行，最后以换行结束
            class_pattern = re.compile(r'(\n\S+){4,5}\n')
            
            # 遍历课表行列
            for col in range(1, 8):  # 1..7
                for row in range(3, 8):  # 3..7 (xlrd 0 索引，对应 Excel 第 4-8 行)
                    # 检查单元格类型是否为文本
                    if schedule.cell_type(row, col) == xlrd.XL_CELL_TEXT:
                        cell_value = schedule.cell_value(row, col)
                        if not cell_value:
                            continue
                        
                        # 定义被寻找的文本
                        search_text = str(cell_value)
                        
                        for match in class_pattern.finditer(search_text):
                            # 创建课程对象（weekday=col）
                            course = ClassCalendar(match.group(0), col, Cal.basedatetime)
                            Cal.addCourse(course)
            
            Cal.moveHoliday()
            print(Cal.ShowInfo())
            f.write(Cal.icsGen())
            f.write("END:VCALENDAR\n")
        
        print(f"\n已生成 .ics 文件，文件路径: {file_out_addr}")
        print("提示: 可直接导入 Outlook / Google Calendar / Apple 日历")
        
    except Exception as e:
        print(f"\n生成文件时出错: {e}")
    
    print("\n进程已退出...")


if __name__ == "__main__":
    main()
