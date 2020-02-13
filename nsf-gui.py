import os
import re
import sys
import tkinter
import traceback
from tkinter import messagebox, ttk

import xlwings as xl

VERSION = "v1.1-rc"

# 检查环境
if getattr(sys, 'frozen', False):  # 运行于 |PyInstaller| 二进制环境
    basedir = sys._MEIPASS  # pylint: disable=no-member
else:  # 运行于一般Python 环境
    basedir = os.path.dirname(__file__)


#Tools Function
def name_formatter(nameList: list, number=2):
    """
    在两个字的名字中间加两个空格，返回经过处理的列表。列表中三个字的名字不会被处理。

    nameList: 一个包含待处理名字的列表

    示例：
    IN: ['张三', '李四', '王老五']
    OUT: ['张三', '李  四', '王老五']
    """
    r = []
    for i in nameList:
        if isinstance(i, str) and len(i) == 2:
            #TODO: add a setting entry to set the space number
            i = i[0] + ' ' * number + i[1]
        r.append(i)
    return r


def remove_all_space(nameList: list):
    "Return a list without any space"
    r = []
    for i in nameList:
        if isinstance(i, str):
            i = i.replace(' ', '')
        r.append(i)
    return r


class Application(tkinter.Frame):
    def __init__(self, master: tkinter.Tk = None):
        #init xl objects
        self.currentSelectRange = None
        #init vars
        self.rangeText = tkinter.StringVar()
        self.currentBook = tkinter.StringVar()
        self.currentSheet = tkinter.StringVar()
        #init settings vars
        self.transposeBool = tkinter.BooleanVar(value=True)
        self.bypassRegExCheck = tkinter.BooleanVar(value=False)
        #init methon vars
        # select = getRangeBySelect()
        # entry = getRangeByEntry()
        self.methodToGetRange = tkinter.StringVar(value='select')
        self.methodToGetRange.trace_add('write', self.refreshEntryState)

        #INIT FUNCTIONS
        super().__init__(master)
        self.master = master
        self.grid()
        self._createWidget()
        self.refreshExcelState()
        self.master.bind('<FocusIn>',
                         lambda *keywords: self.refreshExcelState())

    def _createWidget(self):
        #Fathers
        self.pW = ttk.PanedWindow(orient=tkinter.VERTICAL)
        self.statusLF = ttk.LabelFrame(self.pW, text='状态')
        self.settingsLF = ttk.LabelFrame(self.pW, text='参数设置', padding=5)
        self.rangeSelectorLF = ttk.LabelFrame(self.pW,
                                              text='选择单元格范围',
                                              padding=5)
        self.pW.add(self.statusLF)
        self.pW.add(self.settingsLF)
        self.pW.add(self.rangeSelectorLF)
        self.pW.grid(column=0, row=0, columnspan=3, padx=10, pady=10)

        #CurrentBook
        self.l2 = ttk.Label(self.statusLF, text='当前工作簿：')
        self.l2.grid(column=0, columnspan=1, row=0)
        self.bookLabel = ttk.Label(self.statusLF,
                                   textvariable=self.currentBook)
        self.bookLabel.grid(column=1, columnspan=3, row=0, sticky='w')
        #CurrentSheet
        self.l3 = ttk.Label(self.statusLF, text='当前工作表：')
        self.l3.grid(column=0, columnspan=1, row=1)
        self.sheetLabel = ttk.Label(self.statusLF,
                                    textvariable=self.currentSheet)
        self.sheetLabel.grid(column=1, columnspan=3, row=1, sticky='w')

        #Settings
        self.descriptionLabel = ttk.Label(self.settingsLF,
                                          text='高级设置，建议保持默认，谨慎修改。')
        self.descriptionLabel.grid(column=0, row=0, columnspan=2)
        self.transposeCBox = ttk.Checkbutton(self.settingsLF,
                                             text='垂直写入（Transpose）',
                                             variable=self.transposeBool)
        self.transposeCBox.grid(column=0, row=1)
        self.bypassRegExCheckBtn = ttk.Checkbutton(
            self.settingsLF, text='跳过正则表达式检查', variable=self.bypassRegExCheck)
        self.bypassRegExCheckBtn.grid(column=1, row=1)

        #rangeSelector
        #self.l1 = ttk.Label(self.rangeSelectorLF, text="输入待处理单元格范围（例：A1:A10）:")
        #self.l1.grid(column=0, columnspan=1, row=1, padx=5, sticky=tkinter.E)
        self.methodEntryRB = ttk.Radiobutton(self.rangeSelectorLF,
                                             text="输入待处理单元格范围",
                                             value='entry',
                                             variable=self.methodToGetRange)
        self.methodEntryRB.grid(column=0, row=0, sticky=tkinter.W)
        self.e1 = ttk.Entry(self.rangeSelectorLF,
                            textvariable=self.rangeText,
                            state='disabled')
        self.e1.grid(column=1,
                     columnspan=1,
                     row=0,
                     rowspan=2,
                     padx=5,
                     pady=5,
                     sticky=tkinter.E)
        self.methodSelectRB = ttk.Radiobutton(self.rangeSelectorLF,
                                              text='处理当前选中单元格',
                                              value='select',
                                              variable=self.methodToGetRange,
                                              command=self.refreshExcelState)
        self.methodSelectRB.grid(column=0, row=1, sticky=tkinter.W)
        #Command
        self.refreshBtn = ttk.Button(text='刷新', command=self.refreshExcelState)
        self.refreshBtn.grid(column=0, row=2, padx=10, sticky='w')
        self.removeSpaceBtn = ttk.Button(text='删除空格',
                                         command=self.removeAllSpace)
        self.removeSpaceBtn.grid(column=1, row=2, sticky='e')
        self.btn1 = ttk.Button(text='人名添加空格', command=self.doFormatJobs)
        self.btn1.grid(column=2, row=2, padx=10, pady=5, sticky='e')

    def refreshEntryState(self, *keywords):
        'When methon is select, disable the entry to prevent writ.'
        t = self.methodToGetRange.get()
        if t == 'select':
            self.e1.state(['disabled'])
        else:
            self.e1.state(['!disabled'])

    @staticmethod
    def checkRangeText(rText):
        "Check the Range text as A1:A10, return bool"
        regex = re.compile('^[A-Z]{1,2}[0-9]+:[A-Z]{1,2}[0-9]+$')
        result = regex.match(rText)
        if result:
            return True
        else:
            return False

    def getRangeBySelect(self):
        "Get current selected range, return None if have problem"
        return xl.apps.active.selection.options(
            transpose=self.transposeBool.get())

    def getRangeByEntry(self):
        "Get current range by the entry's value, return None if have problem"
        rT = self.rangeText.get()
        if not self.bypassRegExCheck.get() and not self.checkRangeText(rT):
            messagebox.showerror('rangeText Error', '单元格范围格式不正确')
            return None
        sR = xl.Range(rT).options(transpose=self.transposeBool.get())
        return sR

    def getExcelRange(self):
        "Get current range, return None if have problem"
        try:
            if self.methodToGetRange.get() == 'select':
                return self.getRangeBySelect()
            else:
                return self.getRangeByEntry()
        except AttributeError:
            messagebox.showerror('AttributeError', '无法获取当前工作簿，请检查Excel是否正在运行？')
            raise
        except:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            tbL = traceback.format_exception(exc_type, exc_value,
                                             exc_traceback)
            tbS = ''
            for i in tbL:
                tbS = tbS + '\n' + i
            if exc_type.__module__:
                eTitle = f'{exc_type.__module__}.{exc_type.__name__}'
            else:
                eTitle = str(exc_type.__name__)
            # May replace the tbS as traceback.format_exception_only()
            messagebox.showerror(eTitle, tbS)
            raise

    def replaceRangeValue(self, target: xl.Range, out):
        origin = target.value
        exeBool = messagebox.askyesno(
            '执行确认', f'IN: {origin}\nOUT: {out}\n是否写入？此操作无法撤销', icon='question')
        if exeBool:
            target.value = out
            messagebox.showinfo('Successful?', '写入完成，请至Excel中查看效果。')
            xl.apps.active.activate(steal_focus=True)

    def doFormatJobs(self):
        targetRange = self.getExcelRange()
        if targetRange == None:
            return
        rST = name_formatter(targetRange.value)
        self.replaceRangeValue(targetRange, rST)

    def removeAllSpace(self):
        targetRange = self.getExcelRange()
        if targetRange == None:
            return
        nL = targetRange.value
        result = remove_all_space(nL)
        self.replaceRangeValue(targetRange, result)

    def refreshExcelState(self):
        try:
            self.currentBook.set(xl.books.active.name)
            self.currentSheet.set(xl.books.active.sheets.active.name)
            if self.methodToGetRange.get() == 'select':
                sR = self.getRangeBySelect()
                self.rangeText.set(sR.get_address(False, False))
        except AttributeError:
            messagebox.showerror('ERROR', '无法检测到打开的工作簿，请检查Excel是否正在运行？')
            return


#INIT
root = tkinter.Tk()

# set ICON
root.iconbitmap(os.path.join(basedir, 'icon.ico'))

root.title(f"人名空格补齐实用程序 {VERSION}")
root.resizable(False, False)
app = Application(root)
app.mainloop()
