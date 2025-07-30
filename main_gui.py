import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from pathlib import Path
from PIL import Image, ImageTk
import logging
from queue import Queue
from report_worker import Report, TASK_FINISH, CRITICAL_ERROR

# 全局变量，控制日志写入的到滚动文本框中：
logger = logging.getLogger("report")

# 程序的执行目录
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    EXE_DIR = sys._MEIPASS
else:
    EXE_DIR = Path.cwd()


class QueueHandler(logging.Handler):
    """Class to send logging records to a queue
    It can be used from different threads
    The GUI class polls this queue to display records in a ScrolledText widget
    """

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)


class GUI:
    def __init__(self, root, version=''):
        # 初始化变量
        self.root = root
        self.version = version

        # 初始化图像，此变量需要与 root.mainloop 同级别
        potin_pic = str(Path(EXE_DIR, r'templates', 'potin.png'))
        self.logo = ImageTk.PhotoImage(Image.open(potin_pic).resize((190, 20), Image.LANCZOS))

        # 原始记录xlsm文件路径
        self.xlsm_file = tk.StringVar()

        # 生成文件的类型：
        self._type_lst = ["原始记录", "检验报告", "报告+记录"]
        self.task_type = tk.StringVar()
        self.task_type.set(self._type_lst[2])

        # 是否打开原始记录的修订模式
        self.is_revision_mode = tk.BooleanVar(value=True)

        # 最终生成的输出文件（包含全路径的字符串）
        self.output_name = ''

        # Create a logging handler using a queue
        self.log_queue = Queue()
        self.queue_handler = QueueHandler(self.log_queue)

        # 设置日志级别（level） 和 输出格式Formatters（日志格式器）
        logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter(
            '%(asctime)s [%(levelname)s]: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
            # datefmt='%H:%M:%S',
            style='%'
        )
        self.queue_handler.setFormatter(formatter)
        logger.addHandler(self.queue_handler)

        # 创建界面 Widgets
        self.create_ui()

        # Start polling messages from the queue
        self.root.after(100, self.poll_log_queue)

        # 开始窗口循环：
        # self.root.mainloop()

    def create_ui(self):
        # 开始设置主窗体
        self.root.title(f"报告自动化生成工具({self.version})    by liugang@caict.ac.cn")
        width_screen = self.root.winfo_screenwidth()  # 获取屏幕宽
        height_screen = self.root.winfo_screenheight()  # 获取屏幕高
        width_root = 620  # 指定当前窗体宽
        height_root = 600  # 指定当前窗体高
        # 设置窗体大小及居中
        self.root.geometry("%dx%d+%d+%d" % (
            width_root, height_root, (width_screen - width_root) / 2, (height_screen - height_root) / 2))
        self.root.resizable(False, False)

        # 更换窗口的小图标：
        app_pic = str(Path(EXE_DIR, r'templates', 'app.ico'))
        self.root.iconbitmap(app_pic)

        # 主框架
        main_frame = ttk.Frame(self.root, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame,
            text="检测报告自动化生成工具",
            font=("微软雅黑", 16, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(pady=(5, 15))

        # #######################################
        # 文件夹选择部分 file_frame
        #########################################
        file_frame = ttk.Frame(main_frame, padding=5)
        file_frame.pack(fill=tk.X, pady=5)

        file_entry = ttk.Entry(file_frame, textvariable=self.xlsm_file, state="readonly")
        file_entry.pack(side=tk.LEFT, fill=tk.X, padx=(0, 10), expand=tk.TRUE)

        browse_btn = tk.Button(
            file_frame,
            text="打开记录",
            command=self.on_get,
            bg="#4f7fDD",
            fg="white",
            relief=tk.RAISED,
            borderwidth=2,
            width=12,
            # activebackground="orange"
        )
        browse_btn.pack(side=tk.RIGHT)
        # file_frame.columnconfigure(1, weight=1)  # 让输入框可以扩展

        # #######################################
        # 选项部分  option_frame
        #########################################
        # 输出格式选择
        option_frame = ttk.Frame(main_frame, padding=5)
        option_frame.pack(anchor=tk.W, fill=tk.X, pady=5)
        # 选择输出格式：
        ttk.Label(option_frame, text="输出类型：").pack(side=tk.LEFT)
        # ttk.OptionMenu(root, 绑定的变量,初始值，值列表，command=处理函数）
        self.output_format = ttk.OptionMenu(option_frame, self.task_type, self.task_type.get(), *self._type_lst)
        self.output_format.configure(width=10)
        self.output_format.pack(side=tk.LEFT, padx=(0, 10))

        # 选择是否打开记录修订模式：
        ttk.Checkbutton(option_frame, variable=self.is_revision_mode, text="打开修订模式").pack(side=tk.LEFT,
                                                                                                padx=(20, 40))

        # 生成按钮
        self.generate_btn = tk.Button(
            option_frame,
            text="开始生成",
            command=self.on_generate,
            state="disabled",
            bg="#4f7fDD",
            fg="white",
            relief=tk.RAISED,
            borderwidth=2,
            width=12,
            disabledforeground="darkgray"
        )
        self.generate_btn.pack(side=tk.RIGHT, anchor=tk.E)
        clear_log_btn = tk.Button(
            option_frame,
            text="清空日志",
            command=self.on_clear,
            bg="#4f7fDD",
            fg="white",
            relief=tk.RAISED,
            borderwidth=2,
            width=12
        )
        clear_log_btn.pack(side=tk.RIGHT, padx=10)

        #########################################
        # 日志框部分 log_frame
        #########################################
        log_frame = ttk.LabelFrame(main_frame, text=" 操作日志", padding=1)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=(3, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            width=40,
            height=4,
            font=('微软雅黑', 9),
            state="disabled"
        )
        # 日志框默认样式
        self.log_text.configure(
            # background="lightyellow",
            background="ivory",
            foreground="blue",
            insertbackground="black"
        )
        # 日志框根据日志的级别配置不同的颜色
        self.log_text.tag_config('INFO', foreground='darkgreen')
        self.log_text.tag_config('DEBUG', foreground='gray')
        self.log_text.tag_config('WARNING', foreground='orange')
        self.log_text.tag_config('ERROR', foreground='red')
        self.log_text.tag_config('CRITICAL', foreground='red', underline=True)

        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 日志框插入 Copyright 信息：
        self.log_text.configure(state='normal')
        self.log_text.insert(1.0, "版权声明： 本工具仅限【博鼎实华（北京）技术有限公司】内部员工使用\n")
        self.log_text.configure(state='disabled')

        # #######################################
        # 状态栏部分
        #########################################
        status_logo = ttk.Label(main_frame, image=self.logo)
        status_logo.pack(side=tk.LEFT)
        status_logo.bind("<Double-1>", self.open_dir)
        status_text = tk.Label(main_frame, text="Copyright © 2022，刘刚, All rights reserved.  ", foreground='darkblue')
        status_text.pack(side=tk.RIGHT)

        # 注册窗口退出程序：
        # self.root.protocol('WM_DELETE_WINDOW', self.quit)

    # 让滚动文本框及时显示内容
    def log_display(self, record):
        # 未被格式化的原始log信息：record.msg
        # 判断一下任务是否已经完成：
        if TASK_FINISH in record.msg:
            # print(record.msg)
            temp_msg = record.msg[len(TASK_FINISH):]
            if temp_msg != CRITICAL_ERROR:
                self.output_name = temp_msg
                new_name = Path(self.output_name).stem
                output_excel = Path(self.xlsm_file.get()).parent / (new_name + '.xlsm')
                self.xlsm_file.set(str(output_excel))
                if (messagebox.askyesno("查看生成结果", "   任务已完成，是否立即查看生成的文档？\n\n（后续也可通过双击左下角博鼎Logo查看）")):
                    self.open_dir(Path(output_excel.parent))
            # 修改“生成” 按钮的文字和状态
            self.generate_btn.configure(text="开始生成", state=tk.NORMAL)

        else:
            # 格式化化后的log信息：
            msg = self.queue_handler.format(record)
            self.log_text.configure(state='normal')
            self.log_text.insert(tk.END, msg + '\n', record.levelname)
            self.log_text.configure(state='disabled')
            # Autoscroll to the bottom
            self.log_text.see(tk.END)

    # 轮询queue队列，如果有消息就调用log_display 在滚动文本框中显示：
    def poll_log_queue(self):
        # Check every 100ms if there is a new message in the queue to display
        while not self.log_queue.empty():
            self.log_display(self.log_queue.get(block=False))

        self.root.after(100, self.poll_log_queue)

    # 获取xlsm格式的原始记录路径：
    def on_get(self):
        file_get = filedialog.askopenfile(title="请选择Excel版本的原始记录",
                                          # filetypes=[("Excel 文件", "*.xls*"), ("全部文件", "*")],
                                          filetypes=[("Excel 文件", "*.xls*")],
                                          initialdir="D:\\Report")
        if file_get and Path(file_get.name).is_file():
            self.xlsm_file.set(file_get.name)
            logger.info(f"已选择文件: {file_get.name}")
            self.generate_btn.configure(state='normal')
        else:
            logger.warning("未选择有效的记录文件！")
            if not Path(self.xlsm_file.get()).is_file():
                self.generate_btn.configure(state='disabled')

    def on_generate(self):
        self.generate_btn.configure(state=tk.DISABLED)
        task_type = self._type_lst.index(self.task_type.get())

        self.report_worker = Report(xlsm_file=self.xlsm_file.get(),
                                    task_type=task_type,
                                    is_revision_mode=self.is_revision_mode.get()
                                    )
        # 修改自动出报告模块中的日志输出为全局变量logger
        # clock_test.report_logger = logger
        self.report_worker.start()

    def on_clear(self):
        """清空日志框"""
        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(1.0, "版权声明： 本工具仅限【博鼎实华（北京）技术有限公司】内部员工使用\n")
        self.log_text.update()
        self.log_text.configure(state="disabled")

    def open_dir(self, event):
        try:
            if self.output_name and Path(self.output_name).exists():
                os.startfile(Path(self.output_name).parent)
            else:
                os.startfile(Path(self.xlsm_file.get()).parent)
        except Exception as e:
            messagebox.showerror("Error", f"无法打开目录 {Path(self.output_name).parent},报错: {e}")


def main():
    root = tk.Tk()
    logger.setLevel(level=logging.DEBUG)
    app = GUI(root, version='V20250708')
    root.mainloop()


if __name__ == "__main__":
    main()
