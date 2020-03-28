# import tkinter
# from Tkinter import *
import tkinter as tk
from tkinter.filedialog import *
from tkinter import filedialog
import os

application_window = tk.Tk()
file_types = [('excel文件', '.xls')]
file_types1 = [('excel文件', '.xls'), ('excel文件', '.xlsx')]
# f = askopenfilename(title='askopenfilename', initialdir="D:", filetypes=[('所有文件', '*.*'), ('Python源文件', '.py')])
# f2 = askopenfilename(title='选择源文件', initialdir="c:", filetypes=file_types1)
answer = filedialog.askopenfilenames(parent=application_window,
                                     initialdir=os.getcwd(),
                                     title="选择一个或多个源文件",
                                     filetypes=file_types1)
if len(answer) != 0:
    string_filename = ""
    for i in range(0, len(answer)):
        string_filename += str(answer[i]) + "\n"
    print("您选择的文件是：" + string_filename)
else:
    print("您没有选择任何文件")
# f1 = asksaveasfilename(title='asksaveasfilename', initialdir="E:", filetypes=[('所有文件', '*.*'), ('Python源文件', '.py')])
# askopenfiles()
'''

class Application(tkinter.Frame):
    def say_hi(self):
        print("hi there, everyone!")

    def createWidgets(self):
        self.QUIT = tkinter.Button(self)
        self.QUIT["text"] = "QUIT"
        self.QUIT["fg"]   = "red"
        self.QUIT["command"] =  self.quit

        self.QUIT.pack({"side": "left"})

        self.hi_there = tkinter.Button(self)
        self.hi_there["text"] = "Hello",
        self.hi_there["command"] = self.say_hi

        self.hi_there.pack({"side": "left"})

    def __init__(self, master=None, **kw):
        tkinter.Frame.__init__(self, master)
        super().__init__(master, **kw)
        self.pack()
        self.createWidgets()

root = tkinter.Tk()
app = Application(master=root)
app.mainloop()
root.destroy()
'''
