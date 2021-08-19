# -*- coding:utf-8 -*-
import os
from enum import Enum, unique

import xlrd
import xlwt
import csv
from decimal import Decimal
from os import getcwd
# from tkinter import Tk, Frame, Label, Button
from tkinter import *
from tkinter.filedialog import askdirectory, askopenfilename
from tkinter.messagebox import showerror

property_file = "property.csv"


# TODO 列号沿用Excel的格式
# TODO 报错处理


def isChinese(c):
    if '\u4e00' <= c <= '\u9fff':
        return True
    return False


def getWidth(s):
    # 对于label和entry 一个字母或数字宽度为1 一个汉字宽度为2
    width = 0
    for c in s:
        if isChinese(c):
            width += 2
        else:
            width += 1

    return width


def handlerAdaptor(fun, **kwds):
    return lambda event, fun=fun, kwds=kwds: fun(event, **kwds)


def clearFrame(frame):
    for widget in frame.winfo_children():
        widget.destroy()


def getComponents(frame):
    res = [frame]
    for widget in frame.winfo_children():
        res += getComponents(widget)
    return res


class UI:
    def __init__(self, root, width, height):
        self.root = root
        self.root_width = width
        self.root_height = height
        self.nav_width = int(0.14 * self.root_width)
        padding = 40
        self.main_width = self.root_width - self.nav_width - 2 * padding

        self.nav_bar = Frame(self.root,
                             width=self.nav_width,
                             height=self.root_height)
        self.nav_bar.pack(side="left", fill="y")
        self.nav_bar.grid_propagate(0)
        self.nav_bar.pack_propagate(0)

        border_right = Label(self.root,
                             borderwidth=0.1,
                             relief=GROOVE,
                             bg="black")
        border_right.pack(side="left", fill="y")

        padding_left = Frame(self.root,
                             width=padding,
                             height=self.root_height)
        padding_left.pack(side="left", fill="y")

        self.main_frame = Frame(self.root,
                                width=self.main_width,
                                height=self.root_height)
        self.main_frame.pack(side="left", fill="y")
        self.main_frame.grid_propagate(0)
        self.main_frame.pack_propagate(0)

        padding_right = Frame(self.root,
                              width=padding,
                              height=self.root_height)
        padding_right.pack(side="left", fill="y")

        self.pages = {'0': "主页", '1': "设置"}
        self.page = '0'

        self.executor = Executor()

    def GUIManager(self):
        self.nav()
        self.mainPage()

    def mainPage(self):
        outer_frame = 0
        canvas = 0
        inner_frame = 0

        def scroll(event):
            d = int(-event.delta / 120)
            canvas.yview_scroll(d, "units")

        def bind_scroll():
            for widget in getComponents(outer_frame):
                widget.bind("<MouseWheel>", scroll)

        def scrollable_frame_before():
            nonlocal outer_frame, canvas, inner_frame
            outer_frame = Frame(self.main_frame,
                                highlightthickness=0)
            scrollbar = Scrollbar(outer_frame, orient="vertical")
            canvas = Canvas(outer_frame,
                            yscrollcommand=scrollbar.set,
                            width=int(0.85 * self.main_width),
                            height=int(0.75 * self.root_height),
                            highlightthickness=0)
            canvas.grid(row=0, column=0)
            scrollbar.config(command=canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")

            inner_frame = Frame(canvas)
            canvas.create_window((0, 0), window=inner_frame, anchor='nw')
            outer_frame.grid(row=1, column=0, columnspan=8)

        def scrollable_frame_after():
            nonlocal outer_frame, canvas
            outer_frame.update()
            canvas.config(scrollregion=canvas.bbox("all"))

        def selectDir():
            self.executor.selectDirAndExecute()

            clearFrame(outer_frame)
            scrollable_frame_before()
            i = 0
            for filename in self.executor.excelManager.filenames:
                label = Label(inner_frame, text=filename, font=("宋体", 12))
                label.grid(row=i, column=0)
                i += 1
            bind_scroll()
            scrollable_frame_after()

        def selectFile():
            self.executor.selectFileAndExecute()

            clearFrame(outer_frame)
            scrollable_frame_before()
            i = 0
            for filename in self.executor.excelManager.filenames:
                label = Label(inner_frame, text=filename, font=("宋体", 12))
                label.grid(row=i, column=0)
                i += 1
            bind_scroll()
            scrollable_frame_after()

        Button(self.main_frame,
               text="选择文件夹",
               font=("宋体", 12),
               width=10,
               command=selectDir).grid(row=0, column=0, columnspan=3, pady=25, sticky=W)
        Button(self.main_frame,
               text="选择文件",
               font=("宋体", 12),
               width=10,
               command=selectFile).grid(row=0, column=2, columnspan=3, pady=25, sticky=W)

        scrollable_frame_before()
        bind_scroll()
        scrollable_frame_after()

    def settings(self):
        outer_frame = 0
        canvas = 0
        inner_frame = 0
        property_label_list = []
        button_delete_list = []

        def scroll(event):
            d = int(-event.delta / 120)
            canvas.yview_scroll(d, "units")

        def bind_scroll():
            for widget in getComponents(outer_frame):
                widget.bind("<MouseWheel>", scroll)

        def scrollable_frame_before():
            nonlocal outer_frame, canvas, inner_frame
            outer_frame = Frame(self.main_frame,
                                width=self.main_width,
                                height=self.root_height,
                                highlightthickness=0)
            scrollbar_v = Scrollbar(outer_frame, orient="vertical")
            scrollbar_h = Scrollbar(outer_frame, orient="horizontal")
            canvas = Canvas(outer_frame,
                            yscrollcommand=scrollbar_v.set,
                            xscrollcommand=scrollbar_h.set,
                            width=self.main_width,
                            height=self.root_height - 18,
                            highlightthickness=0)
            canvas.grid(row=0, column=0, sticky="ns")
            scrollbar_v.config(command=canvas.yview)
            scrollbar_v.grid(row=0, column=1, sticky="ns")
            scrollbar_h.config(command=canvas.xview)
            scrollbar_h.grid(row=1, column=0, sticky="ew")

            inner_frame = Frame(canvas)
            canvas.create_window((0, 0), window=inner_frame, anchor='nw')
            outer_frame.grid(row=0, column=0, columnspan=6)

        def scrollable_frame_after():
            nonlocal outer_frame, canvas
            outer_frame.update()
            canvas.config(scrollregion=canvas.bbox("all"))

        def focusOut(event, label, entry, row, column):
            try:
                self.executor.isValidProperty(self.executor.property.rows[row])
            except ValueError:
                showerror("错误", "配置不合法：检查是否都为正整数")
                return
            except RangeError:
                showerror("错误", "配置不合法：检查 起始 <= 结束")
                return

            entry.destroy()
            self.executor.property.width[column] = max(self.executor.property.default_width[column],
                                                       getWidth(self.executor.property.rows[row][column].get()))
            label.config(width=self.executor.property.width[column])
            label.config(text=self.executor.property.rows[row][column].get())
            label.grid(row=row + 2, column=column, sticky=W + E)

            for widget in getComponents(self.main_frame):
                widget.unbind("<Button-1>")
            # 立即更新配置
            self.executor.updateRow()

            scrollable_frame_after()

        def editProperty(event, label):
            index = property_label_list.index(label)
            row = int(index / self.executor.property.attr_num)
            column = index % self.executor.property.attr_num

            label.grid_forget()
            print(self.executor.property.width)
            entry = Entry(inner_frame,
                          textvariable=self.executor.property.rows[row][column],
                          font=("宋体", "12"),
                          width=self.executor.property.width[column])
            entry.bind("<MouseWheel>", scroll)
            entry.focus_set()
            entry.grid(row=row + 2, column=column, sticky=W + E)
            for widget in getComponents(self.main_frame):
                if widget == entry:
                    continue
                widget.bind("<Button-1>", handlerAdaptor(focusOut, label=label, entry=entry, row=row, column=column))

        def confirmAdd(row, column, buffer, entry_list, button_add, button_confirm, button_cancel):
            try:
                self.executor.isValidProperty(buffer)
            except ValueError:
                showerror("错误", "配置不合法：检查是否都为正整数")
                return
            except RangeError:
                showerror("错误", "配置不合法：检查 起始 <= 结束")
                return

            for entry in entry_list:
                entry.destroy()

            self.executor.addProperty(buffer)

            button_confirm.grid_forget()
            button_cancel.grid_forget()

            i = 0
            for data in buffer:
                label = Label(inner_frame,
                              text=data.get(),
                              font=("宋体", "12"),
                              width=self.executor.property.width[i],
                              height=2)
                label.bind("<Double-Button-1>", handlerAdaptor(editProperty, label=label))
                label.bind("<MouseWheel>", scroll)
                property_label_list.append(label)
                label.grid(row=row, column=column + i, sticky=W + E)
                i += 1

            button_delete = Button(inner_frame,
                                   text="删除",
                                   font=("宋体", "12"),
                                   width=4,
                                   command=lambda index=row - 2: deleteRow(index, button_delete))
            button_delete.grid(row=row, column=column + i, sticky=W + E)
            button_delete_list.append(button_delete)

            button_add.config(command=lambda: addRow(row + 1, 0, button_add))
            button_add.grid(row=row + 1, column=column, pady=16, sticky=W + E)

            scrollable_frame_after()

        def cancelAdd(row, column, entry_list, button_add, button_confirm, button_cancel):
            for entry in entry_list:
                entry.destroy()

            button_confirm.grid_forget()
            button_cancel.grid_forget()

            button_add.grid(row=row + 1, column=column, pady=16, sticky=W + E)

            scrollable_frame_after()

        def addRow(row, column, button_add):
            button_add.grid_forget()
            r = row
            c = column
            buffer = [StringVar() for idx in range(0, self.executor.property.attr_num)]
            entry_list = []
            for i in range(0, self.executor.property.attr_num):
                entry = Entry(inner_frame,
                              textvariable=buffer[i],
                              font=("宋体", "12"),
                              width=self.executor.property.width[i])
                entry.bind("<MouseWheel>", scroll)
                entry.grid(row=r, column=c, sticky=W + E)
                entry_list.append(entry)
                c += 1

            button_confirm = Button(inner_frame,
                                    text="确认",
                                    font=("宋体", "12"),
                                    width=4,
                                    command=lambda: confirmAdd(row, column, buffer, entry_list, button_add,
                                                               button_confirm, button_cancel))
            button_confirm.grid(row=r, column=c, padx=10, sticky=W + E)

            c += 1
            button_cancel = Button(inner_frame,
                                   text="取消",
                                   font=("宋体", "12"),
                                   width=4,
                                   command=lambda: cancelAdd(row, column, entry_list, button_add, button_confirm,
                                                             button_cancel))
            button_cancel.grid(row=r, column=c, sticky=W + E)

            scrollable_frame_after()

        def deleteRow(row, button_delete):
            button_delete_list.remove(button_delete)
            button_delete.destroy()

            index = row * self.executor.property.attr_num
            for i in range(0, self.executor.property.attr_num):
                property_label_list[index].destroy()
                del property_label_list[index]

            self.executor.deleteProperty(row)

            displayPropertyLabels()
            displayButtonDelete()

            scrollable_frame_after()

        def displayButtonDelete():
            i = 0
            for button_delete in button_delete_list:
                button_delete.config(
                    command=lambda _row=i, _button_delete=button_delete: deleteRow(_row, _button_delete))
                button_delete.grid(row=i + 2, column=self.executor.property.attr_num, sticky=W + E)
                i += 1

        def displayPropertyLabels():
            i = 0
            j = 0
            for label in property_label_list:
                label.grid(row=i + 2, column=j, sticky=W + E)
                j += 1
                if j == self.executor.property.attr_num:
                    i += 1
                    j = 0

        def drawProperty():
            i = 1
            j = 0
            for attr in self.executor.property.headers:
                label = Label(inner_frame,
                              text=attr,
                              font=("宋体", "12"),
                              width=self.executor.property.width[j],
                              height=2)
                label.grid(row=i, column=j, sticky=W + E)
                j += 1

            i += 1
            j = 0
            # row 在 self.executor.property.rows 中的索引
            row_index = 0
            for row in self.executor.property.rows:
                for data in row:
                    label = Label(inner_frame,
                                  text=data.get(),
                                  font=("宋体", "12"),
                                  width=self.executor.property.width[j],
                                  height=2)
                    label.bind("<Double-Button-1>", handlerAdaptor(editProperty, label=label))
                    property_label_list.append(label)
                    j += 1

                button_delete = Button(inner_frame,
                                       text="删除",
                                       font=("宋体", "12"),
                                       width=4)
                button_delete_list.append(button_delete)

                row_index += 1
                i += 1
                j = 0

            button_add = Button(inner_frame,
                                text="新增",
                                font=("宋体", "12"),
                                width=8,
                                command=lambda: addRow(i, 0, button_add))
            button_add.grid(row=i, column=j, pady=16, sticky=W + E)

            displayPropertyLabels()
            displayButtonDelete()

        scrollable_frame_before()
        Label(inner_frame,
              text="模式",
              font=("宋体", "12"),
              width=6).grid(row=0, column=0, pady=25, sticky=W)
        rd1 = Radiobutton(inner_frame,
                          text="列",
                          font=("宋体", 12),
                          width=4,
                          anchor="w",
                          variable=self.executor.property.mode,
                          value=0,
                          command=self.executor.updateMode)
        rd2 = Radiobutton(inner_frame,
                          text="行",
                          font=("宋体", 12),
                          width=4,
                          anchor="w",
                          variable=self.executor.property.mode,
                          value=1,
                          command=self.executor.updateMode)
        rd1.grid(row=0, column=1, pady=25, sticky=W)
        rd2.grid(row=0, column=2, pady=25, sticky=W)
        drawProperty()
        bind_scroll()
        scrollable_frame_after()

    def switchPage(self, event, page_num):
        if self.page == page_num:
            return

        clearFrame(self.nav_bar)
        clearFrame(self.main_frame)
        if page_num == '0':
            self.mainPage()
        elif page_num == '1':
            self.settings()
        else:
            print(123456)

        self.page = page_num
        self.nav()

    def nav(self):
        for (key, value) in self.pages.items():
            font = ("宋体", 16)
            if key == self.page:
                font = ("宋体", 16, "bold")
            label = Label(self.nav_bar,
                          text=value,
                          font=font,
                          width=6)
            label.bind("<Button-1>", handlerAdaptor(self.switchPage, page_num=key))
            label.grid(row=key, column=0, ipadx=28, ipady=10)


class RangeError(Exception):
    pass


def precision(num):
    if num < 0:
        raise ValueError
    res = "0."
    for i in range(0, num):
        res += '0'

    return Decimal(res)


class Executor:
    def __init__(self):
        self.excelManager = ExcelManager()
        self.property = Property()

    def execute(self):
        @unique
        class Type(Enum):
            # 全空(一般不会)
            error = 0b00000
            # 正常
            normal = 0b11111
            # 只填了dst, dst_beg, template
            new = 0b00011

        def getType(lst):
            res = 0b0
            for item in lst:
                res <<= 1
                if item != "":
                    res += 1

            return res

        for filename in self.excelManager.filenames:
            self.excelManager.read(filename)

            # 有多个flt时，标记不符合条件的行或列，最后统一移除
            abandon = []

            for _property in self.property.rows:
                # pattern：(flt, condition) 或 (reg, reg_string) 或 (str, []) 或 (dec, round)
                src, src_beg, src_end, dst, dst_beg, pattern, template = [p.get() for p in _property]
                _type = getType(_property[0: 5])

                if _type == Type.error.value:
                    return

                elif _type == Type.new.value:
                    dst = int(dst) - 1
                    dst_beg = int(dst_beg) - 1

                    dst_row = 0
                    dst_column = 0
                    # 列
                    if self.property.mode.get() == 0:
                        dst_row = dst_beg
                        dst_column = dst
                    # 行
                    elif self.property.mode.get() == 1:
                        dst_row = dst
                        dst_column = dst_beg

                    value = template

                    while dst_row >= len(self.excelManager.dst_sheet):
                        self.excelManager.dst_sheet.append(list())
                    while dst_column >= len(self.excelManager.dst_sheet[dst_row]):
                        self.excelManager.dst_sheet[dst_row].append([None, None])
                    self.excelManager.dst_sheet[dst_row][dst_column] = [value, None]

                elif _type == Type.normal.value:
                    print("normal")
                    src = int(src) - 1
                    src_beg = int(src_beg) - 1

                    # 不填 src_end 默认到表格的最后一行/一列
                    if src_end == "":
                        # 列
                        if self.property.mode.get() == 0:
                            src_end = self.excelManager.nrows
                        # 行
                        elif self.property.mode.get() == 1:
                            src_end = self.excelManager.ncols
                    else:
                        src_end = int(src_end) - 1

                    dst = int(dst) - 1
                    dst_beg = int(dst_beg) - 1
                    for offset in range(0, src_end - src_beg + 1):
                        src_row = 0
                        src_column = 0
                        dst_row = 0
                        dst_column = 0
                        # 列
                        if self.property.mode.get() == 0:
                            src_row = src_beg + offset
                            src_column = src
                            dst_row = dst_beg + offset
                            dst_column = dst
                        # 行
                        elif self.property.mode.get() == 1:
                            src_row = src
                            src_column = src_beg + offset
                            dst_row = dst
                            dst_column = dst_beg + offset

                        value = str(self.excelManager.src_sheet.cell_value(src_row, src_column))
                        print(value)

                        if pattern != "":
                            value, parse_res = self.parser(pattern, value)
                        else:
                            parse_res = []

                        filt_res = True
                        for res in parse_res:
                            if isinstance(res, bool):
                                filt_res = filt_res and res

                        # 不满足筛选条件
                        if not filt_res:
                            abandon.append(src_row)
                            continue

                        if template != "":
                            place_holders = re.findall(r"{\s*\d+\s*}", template)
                            place_holders.sort(key=lambda idx: int(re.findall(r"\d+", idx)[0]))
                            _template = template
                            for i in range(0, min(len(parse_res), len(place_holders))):
                                _template = template.replace(place_holders[i], str(parse_res[i]))
                            value = _template

                        while dst_row >= len(self.excelManager.dst_sheet):
                            self.excelManager.dst_sheet.append(list())
                        while dst_column >= len(self.excelManager.dst_sheet[dst_row]):
                            self.excelManager.dst_sheet[dst_row].append([None, None])
                        self.excelManager.dst_sheet[dst_row][dst_column] = [value, None]

            res_buffer = []
            if len(abandon) != 0:
                for i in range(0, len(self.excelManager.dst_sheet)):
                    if i not in abandon:
                        res_buffer.append(self.excelManager.dst_sheet[i])
                self.excelManager.dst_sheet = res_buffer

            self.excelManager.write(filename)

    def parser(self, string, value):
        cmd_lst = string.split(";")
        cmd_lst = [s.strip() for s in cmd_lst]
        res = []
        for cmd_line in cmd_lst:
            cmds = cmd_line.split("->")
            cmds = [s.strip() for s in cmds]
            s = value
            for cmd in cmds:
                operator = re.findall(r"\(\s*(.*?)\s*,", cmd)[0]
                param = re.findall(r",\s*(.*?)\s*\)", cmd)[0]
                if operator == "flt":
                    words = re.findall(r"[^&|!()\s]+", param)
                    for word in words:
                        if word == s:
                            rpl = "1"
                        else:
                            rpl = "0"
                        param = param.replace(word, rpl, 1)
                    param = param.replace("!", " not ").replace("&", " and ").replace("|", " or ")
                    boolean = bool(eval(param))
                    res.append(boolean)
                elif operator == "reg":
                    obj = re.search(param, s)
                    if obj is None:
                        s = ""
                    else:
                        span = obj.span()
                        s = s[span[0]:span[1]]
                elif operator == "str":
                    beg = int(re.findall(r"\[\s*(-?\d*?)\s*:", param)[0])
                    end = int(re.findall(r":\s*(-?\d*?)\s*\]", param)[0])
                    s = s[beg: end]
                elif operator == "dec":
                    print(s)
                    d = Decimal(s)
                    r = int(param)
                    s = str(d.quantize(precision(r)))
            res.append(s)
        return [s, res]

    def selectDirAndExecute(self):
        self.excelManager.selectDir()

        if len(self.excelManager.filenames) == 0:
            showerror("错误", "文件夹中没有xls文件")
            return

        self.execute()

    def selectFileAndExecute(self):
        self.excelManager.selectFile()

        if len(self.excelManager.filenames) == 0:
            showerror("错误", "未选择xls文件")
            return

        self.execute()

    def isValidProperty(self, this_row):
        i = 0
        for data in this_row:
            if i == 5 or i == 6:
                i += 1
                continue
            if data.get() == "":
                continue
            if not data.get().isdigit():
                raise ValueError
            i += 1
        if this_row[1].get() > this_row[2].get():
            raise RangeError()

    def addProperty(self, buffer):
        self.property.rows.append(buffer)
        self.property.add()
        self.property.save()

    def deleteProperty(self, row):
        self.property.delete(row)
        self.property.save()

    def updateMode(self):
        self.property.updateMode()
        self.property.save()

    def updateRow(self):
        self.property.updateRow()
        self.property.save()


class Property:
    """[src row/column, src begin, src end, dst row/column, dst begin]"""

    def __init__(self):
        self.csvManager = CSVManager()

        try:
            self.csvManager.read(property_file)
        except FileNotFoundError:
            fp = open(property_file, 'w')
            fp.close()
            self.csvManager.headers = ["源文件", "起始", "终止", "目标文件", "起始", "命令", "模板", "0"]

        self.mode = IntVar()
        self.headers = []
        self.default_width = [8, 8, 8, 8, 8, 16, 16]
        self.width = [8, 8, 8, 8, 8, 16, 16]
        self.rows = []
        self.mode.set(self.csvManager.headers[-1])

        self.attr_num = len(self.csvManager.headers) - 1
        j = 0
        for attr in self.csvManager.headers:
            if j == self.attr_num:
                j = 0
                break
            self.headers.append(attr)
            j += 1

        for row in self.csvManager.rows:
            row_data = []
            j = 0
            for column in row:
                if j == self.attr_num:
                    break
                data = StringVar()
                data.set(column)
                row_data.append(data)
                # 计算数据的长度
                self.width[j] = max(self.width[j], max(self.default_width[j], getWidth(column)))
                j += 1
            self.rows.append(row_data)

    def add(self):
        for i in range(len(self.csvManager.rows), len(self.rows)):
            row = []
            for j in range(0, self.attr_num):
                row.append(self.rows[i][j].get())
            self.csvManager.rows.append(row)

    def delete(self, row):
        del self.rows[row]
        del self.csvManager.rows[row]

    def updateMode(self):
        self.csvManager.headers[-1] = self.mode.get()

    def updateRow(self):
        for i in range(0, len(self.rows)):
            for j in range(0, self.attr_num):
                self.csvManager.rows[i][j] = self.rows[i][j].get()

    def save(self):
        self.csvManager.write(property_file)


class FileManager:
    def __init__(self):
        self.filenames = []

    def selectDir(self):
        self.filenames.clear()
        file_dir = askdirectory()
        filename_list = os.listdir(file_dir)

        for filename in filename_list:
            file_path = file_dir + '/' + filename
            if os.path.isfile(file_path) and filename.endswith(".xls"):
                self.filenames.append(file_path)

    def selectFile(self):
        self.filenames.clear()
        filename = askopenfilename()

        if filename.endswith(".xls"):
            self.filenames.append(filename)


class ExcelManager(FileManager):
    def __init__(self):
        super().__init__()
        self.src_sheet = []
        self.nrows = 0
        self.ncols = 0
        """[[[label, style], ...], ...]"""
        self.dst_sheet = []
        self.xf_list = []
        self.font_list = []

    def getDefaultStyle(self, row, column):
        return self.xf_list[self.src_sheet.cell_xf_index(row, column)]

    def getDefaultFont(self, row, column):
        return self.font_list[self.src_sheet.cell_xf_index(row, column)]

    def read(self, filename):
        work_book = xlrd.open_workbook(filename, formatting_info=True)
        self.src_sheet = work_book.sheets()[0]
        self.nrows = self.src_sheet.nrows
        self.ncols = self.src_sheet.ncols
        print(self.nrows, self.ncols)
        self.xf_list = work_book.xf_list
        self.font_list = work_book.font_list

    def write(self, filename):
        workbook = xlwt.Workbook(encoding="utf-8")
        worksheet = workbook.add_sheet("sheet")

        _font = xlwt.Font()
        _style = xlwt.XFStyle()
        _font.name = u'宋体'
        _style.font = _font

        for row in range(0, len(self.dst_sheet)):
            for column in range(0, len(self.dst_sheet[row])):
                label, style = self.dst_sheet[row][column]
                if style is None:
                    # font = xlwt.Font()
                    # style = xlwt.XFStyle()
                    # font.name = self.getDefaultFont(row, column).name
                    # style.font = font
                    style = _style
                worksheet.write(row, column, label, style)

        workbook.save(filename + "的结果.xls")


class CSVManager(FileManager):
    def __init__(self):
        super().__init__()
        self.headers = []
        self.rows = []

    def read(self, filename):
        with open(filename) as f:
            f_csv = csv.reader(f)
            i = 0
            for row in f_csv:
                if i == 0:
                    self.headers = row
                else:
                    self.rows.append(row)
                i += 1

        print(self.headers)
        print(self.rows)

    def write(self, filename):
        with open(filename, "w", newline="") as f:
            f_csv = csv.writer(f)
            f_csv.writerow(self.headers)
            f_csv.writerows(self.rows)


def main():
    root = Tk()
    root.title("Excel转换工具0.0.1")
    root.geometry("980x500")
    # root.resizable(width=False, height=False)
    ui = UI(root, 980, 500)
    ui.GUIManager()
    root.mainloop()


if __name__ == "__main__":
    main()
