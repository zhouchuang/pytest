import tkinter as tk

class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('餐补生成工具')
        w = 565
        h = 365
        self.iconbitmap('tool.ico')  # 加图标
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))


        # 程序参数/数据
        self.name = '张三'
        self.age = 30
        # 弹窗界面
        self.setupUI()

    def setupUI(self):
        row1 = tk.Frame(self)
        row1.pack(fill="x")
        ttk.Label(self, text="文件路径").grid(column=0, row=0)
        ttk.Entry(self, width=50, textvariable=pathtext, state='readonly').grid(column=1, row=0)

if __name__ == '__main__':
    app = MyApp()
    app.mainloop()