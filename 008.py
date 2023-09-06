import tkinter as tk
from tkinter import messagebox, filedialog
import openpyxl
import xlrd

class ImageSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片搜索与复制工具")

        self.excel_path = tk.StringVar()
        self.source_path = tk.StringVar()
        self.dest_path = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # 主窗口标题
        title_label = tk.Label(self.root, text="图片搜索与复制工具", font=("Helvetica", 16))
        title_label.pack(pady=10)

        # 选择Excel文件
        excel_frame = tk.Frame(self.root)
        tk.Label(excel_frame, text="Excel文件路径：").pack(side="left")
        tk.Entry(excel_frame, textvariable=self.excel_path, width=40).pack(side="left")
        tk.Button(excel_frame, text="选择Excel文件", command=self.select_excel_file).pack(side="left")
        excel_frame.pack(pady=10)

        # 选择源文件夹
        source_frame = tk.Frame(self.root)
        tk.Label(source_frame, text="源路径：").pack(side="left")
        tk.Entry(source_frame, textvariable=self.source_path, width=40).pack(side="left")
        tk.Button(source_frame, text="选择源文件夹", command=self.select_source_folder).pack(side="left")
        source_frame.pack(pady=10)

        # 选择目标文件夹
        dest_frame = tk.Frame(self.root)
        tk.Label(dest_frame, text="目标路径：").pack(side="left")
        tk.Entry(dest_frame, textvariable=self.dest_path, width=40).pack(side="left")
        tk.Button(dest_frame, text="选择目标文件夹", command=self.select_dest_folder).pack(side="left")
        dest_frame.pack(pady=10)

        # 运行按钮
        run_button = tk.Button(self.root, text="运行", command=self.run_copy)
        run_button.pack()

    def select_excel_file(self):
        excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        self.excel_path.set(excel_path)

    def select_source_folder(self):
        source_folder = filedialog.askdirectory()
        self.source_path.set(source_folder)

    def select_dest_folder(self):
        dest_folder = filedialog
