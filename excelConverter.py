# -*- coding:utf-8 -*-
import os

import xlrd
import xlwt
import csv
from decimal import Decimal
from os import getcwd
# from tkinter import Tk, Frame, Label, Button
from tkinter import *
from tkinter.filedialog import askdirectory, askopenfilename
from tkinter.messagebox import showerror

# font = xlwt.Font()
# style = xlwt.XFStyle()
# font.name = u"宋体"
# style.font = font
property_file = "property.csv"


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
        self.nav_width = int(0.2 * self.root_width)
        self.main_width = self.root_width - self.nav_width

        self.nav_bar = Frame(self.root,
                             width=self.nav_width,
                             height=self.root_height)
        self.nav_bar.pack(side="left", fill="y")

        border_right = Label(self.root,
                             borderwidth=0.1,
                             relief=GROOVE,
                             bg="black")
        border_right.pack(side="left", fill="y")

        self.main_frame = Frame(self.root,
                                width=self.main_width,
                                height=self.root_height)
        self.main_frame.pack(side="left", fill="y")
        self.main_frame.grid_propagate(0)
        self.main_frame.pack_propagate(0)

        self.pages = {'0': "主页", '1': "设置"}
        self.page = '0'

        self.executor = Executor()

    def GUIManager(self):
        self.nav()
        self.mainPage()

    def mainPage(self):
        padding_left = 40
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
                            width=400,
                            height=280,
                            highlightthickness=0)
            canvas.grid(row=0, column=0)
            scrollbar.config(command=canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")

            inner_frame = Frame(canvas)
            canvas.create_window((0, 0), window=inner_frame, anchor='nw')
            outer_frame.grid(row=1, column=0, columnspan=6, padx=padding_left)

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
                label = Label(inner_frame, text=filename)
                label.grid(row=i, column=0)
                i += 1
            bind_scroll()
            scrollable_frame_after()

        def selectFile():
            # TODO
            self.executor.selectFileAndExecute()
            text = ""
            for filename in self.executor.excelManager.filenames:
                text += (filename + '\n')

            file_info.config(text=text)

        Button(self.main_frame,
               text="选择文件夹",
               font=("宋体", 12),
               width=10,
               command=selectDir).grid(row=0, column=0, columnspan=3, padx=padding_left, pady=25, sticky=W)
        Button(self.main_frame,
               text="选择文件",
               font=("宋体", 12),
               width=10,
               command=selectFile).grid(row=0, column=2, columnspan=3, pady=25, sticky=W)

        scrollable_frame_before()
        bind_scroll()
        scrollable_frame_after()

    def settings(self):
        padding_left = 40
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
                                highlightthickness=0,
                                bg="red")
            scrollbar = Scrollbar(outer_frame, orient="vertical")
            canvas = Canvas(outer_frame,
                            yscrollcommand=scrollbar.set,
                            width=self.main_width,
                            height=self.root_height,
                            highlightthickness=0)
            canvas.grid(row=0, column=0, sticky="ns")
            scrollbar.config(command=canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")

            inner_frame = Frame(canvas)
            canvas.create_window((0, 0), window=inner_frame, anchor='nw')
            outer_frame.grid(row=1, column=0, columnspan=6, padx=padding_left)

        def scrollable_frame_after():
            nonlocal outer_frame, canvas
            outer_frame.update()
            canvas.config(scrollregion=canvas.bbox("all"))

        def focusOut(event, label, entry, row, column):
            try:
                self.executor.isValidProperty(row)
            except NullError:
                showerror("错误", "配置不合法：检查是否有空")
                return
            except ValueError:
                showerror("错误", "配置不合法：检查是否都为正整数")
                return
            except RangeError:
                showerror("错误", "配置不合法：检查 起始 <= 结束")
                return

            entry.destroy()
            label.config(text=self.executor.property.rows[row][column].get())
            label.grid(row=row + 2, column=column, sticky=W + E)

            for widget in getComponents(self.main_frame):
                widget.unbind("<Button-1>")
            # 立即更新配置
            self.executor.updateRow()

        def editProperty(event, label):
            index = property_label_list.index(label)
            row = int(index / self.executor.property.attr_num)
            column = index % self.executor.property.attr_num

            label.grid_forget()
            entry = Entry(inner_frame,
                          textvariable=self.executor.property.rows[row][column],
                          font=("宋体", "12"),
                          width=8)
            entry.bind("<MouseWheel>", scroll)
            entry.focus_set()
            entry.grid(row=row + 2, column=column, sticky=W + E)
            for widget in getComponents(self.main_frame):
                if widget == entry:
                    continue
                widget.bind("<Button-1>", handlerAdaptor(focusOut, label=label, entry=entry, row=row, column=column))

        def confirmAdd(row, column, buffer, entry_list, button_add, button_confirm, button_cancel):
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
                              width=8,
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
                              width=8)
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
                              width=8,
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
                                  width=8,
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


class NullError(Exception):
    pass


class Executor:
    def __init__(self):
        self.excelManager = ExcelManager()
        self.property = Property()

    def execute(self):
        for filename in self.excelManager.filenames:
            self.excelManager.read(filename)
            # TODO
            self.excelManager.write(filename)

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

    def isValidProperty(self, row):
        this_row = self.property.rows[row]
        for data in this_row:
            if data.get() == "":
                raise NullError()
            if not data.get().isdigit():
                raise ValueError
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
    """[src row/column, src begin, src end, dest row/column, dest begin]"""

    def __init__(self):
        self.csvManager = CSVManager()
        self.csvManager.read(property_file)
        self.mode = IntVar()
        self.headers = []
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
        """[[(row, column, label, style), ...], ...]"""
        self.sheet = []

    def read(self, filename):
        data = xlrd.open_workbook(filename)
        table = data.sheets()[0]

    def write(self, filename):
        workbook = xlwt.Workbook(encoding="utf-8")
        worksheet = workbook.add_sheet("sheet")
        for row in self.sheet:
            for params in row:
                row, column, label, style = params
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
    root.geometry("650x400")
    root.resizable(width=False, height=False)
    ui = UI(root, 650, 400)
    ui.GUIManager()
    root.mainloop()


if __name__ == "__main__":
    main()
