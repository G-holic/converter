#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文件格式转换工具 - 图形化界面版本
支持格式: Excel (.xlsx), CSV (.csv), JSON (.json), Markdown表格 (.md)
依赖: pandas, openpyxl, tabulate
安装命令: pip install pandas openpyxl tabulate
"""

import os
import sys
import json
import threading
import warnings
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox

# 抑制openpyxl的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 预导入tabulate，确保pandas能正确使用
try:
    import tabulate
except ImportError:
    pass

try:
    import pandas as pd
except Exception as e:
    print(f"pandas导入失败: {e}")
    pd = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


def excel_to_markdown(src, dst, sheet_name=0, encoding='utf-8'):
    try:
        df = pd.read_excel(src, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        raise Exception(f"读取Excel文件失败: {e}")
    
    # 清理单元格内的换行符，避免Markdown表格格式错乱
    df = df.map(lambda x: str(x).replace('\n', ' ').replace('\r', ' ') if pd.notna(x) else x)
    
    try:
        markdown_content = df.to_markdown(index=False)
    except Exception as e:
        # 如果to_markdown失败，手动生成Markdown表格
        markdown_content = _manual_df_to_markdown(df)
    
    with open(dst, 'w', encoding=encoding) as f:
        f.write(markdown_content)
    return True

def excel_to_csv(src, dst, sheet_name=0, encoding='utf-8'):
    df = pd.read_excel(src, sheet_name=sheet_name, engine='openpyxl')
    # 清理单元格内的换行符
    df = df.map(lambda x: str(x).replace('\n', ' ').replace('\r', ' ') if pd.notna(x) else x)
    df.to_csv(dst, index=False, encoding=encoding)
    return True

def csv_to_markdown(src, dst, encoding='utf-8'):
    try:
        df = pd.read_csv(src, encoding=encoding)
    except Exception as e:
        raise Exception(f"读取CSV文件失败: {e}")
    
    # 清理单元格内的换行符
    df = df.map(lambda x: str(x).replace('\n', ' ').replace('\r', ' ') if pd.notna(x) else x)
    
    try:
        markdown_content = df.to_markdown(index=False)
    except Exception as e:
        # 如果to_markdown失败，手动生成Markdown表格
        markdown_content = _manual_df_to_markdown(df)
    
    with open(dst, 'w', encoding=encoding) as f:
        f.write(markdown_content)
    return True

def csv_to_excel(src, dst, encoding='utf-8'):
    df = pd.read_csv(src, encoding=encoding)
    df.to_excel(dst, index=False, engine='openpyxl')
    return True

def json_to_markdown(src, dst, orient='records', encoding='utf-8'):
    try:
        with open(src, 'r', encoding=encoding) as f:
            data = json.load(f)
    except Exception as e:
        raise Exception(f"读取JSON文件失败: {e}")
    
    try:
        if orient == 'records':
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(data)
    except Exception as e:
        raise Exception(f"解析JSON数据失败: {e}")
    
    # 清理单元格内的换行符
    df = df.map(lambda x: str(x).replace('\n', ' ').replace('\r', ' ') if pd.notna(x) else x)
    
    try:
        markdown_content = df.to_markdown(index=False)
    except Exception as e:
        # 如果to_markdown失败，手动生成Markdown表格
        markdown_content = _manual_df_to_markdown(df)
    
    with open(dst, 'w', encoding=encoding) as f:
        f.write(markdown_content)
    return True

def markdown_to_csv(src, dst, encoding='utf-8'):
    df = _parse_markdown_table(src, encoding)
    df.to_csv(dst, index=False, encoding=encoding)
    return True

def markdown_to_excel(src, dst, encoding='utf-8'):
    df = _parse_markdown_table(src, encoding)
    df.to_excel(dst, index=False, engine='openpyxl')
    return True

def _parse_markdown_table(filepath, encoding):
    with open(filepath, 'r', encoding=encoding) as f:
        lines = f.readlines()
    
    # 过滤出表格行
    table_lines = [line.strip() for line in lines if line.strip() and line.strip().startswith('|')]
    
    if len(table_lines) < 2:
        raise ValueError("未找到有效表格")
    
    # 解析表头
    headers = [cell.strip() for cell in table_lines[0].strip('|').split('|')]
    
    # 解析数据行（跳过第2行分隔线）
    data_rows = []
    for row_line in table_lines[2:]:
        cells = [cell.strip() for cell in row_line.strip('|').split('|')]
        
        # 补齐或截断列数以匹配表头
        if len(cells) > len(headers):
            cells = cells[:len(headers)]
        elif len(cells) < len(headers):
            cells.extend([''] * (len(headers) - len(cells)))
        
        data_rows.append(cells)
    
    return pd.DataFrame(data_rows, columns=headers)


def _manual_df_to_markdown(df):
    """手动将DataFrame转换为Markdown表格格式"""
    # 获取列宽
    col_widths = {}
    for col in df.columns:
        max_width = len(str(col))
        for val in df[col]:
            max_width = max(max_width, len(str(val)))
        col_widths[col] = max_width
    
    # 生成表头
    header = '| ' + ' | '.join([str(col).ljust(col_widths[col]) for col in df.columns]) + ' |'
    separator = '| ' + ' | '.join(['-' * col_widths[col] for col in df.columns]) + ' |'
    
    # 生成数据行
    rows = []
    for _, row in df.iterrows():
        row_str = '| ' + ' | '.join([str(row[col]).ljust(col_widths[col]) for col in df.columns]) + ' |'
        rows.append(row_str)
    
    return '\n'.join([header, separator] + rows)



class FileConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件格式转换工具 v1.0")
        self.root.geometry("700x550")
        self.root.minsize(650, 500)
        
        self.input_path = StringVar()
        self.output_path = StringVar()
        self.input_format = StringVar(value="auto")
        self.output_format = StringVar(value="auto")
        self.sheet_name = StringVar(value="0")
        self.json_orient = StringVar(value="records")
        self.encoding = StringVar(value="utf-8")
        
        self.create_widgets()
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(file_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5, sticky=EW)
        ttk.Button(file_frame, text="浏览...", command=self.select_input).grid(row=0, column=2, padx=(5, 0))
        
        ttk.Label(file_frame, text="输出文件:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5, sticky=EW)
        ttk.Button(file_frame, text="浏览...", command=self.select_output).grid(row=1, column=2, padx=(5, 0))
        
        file_frame.columnconfigure(1, weight=1)
        
        format_frame = ttk.LabelFrame(main_frame, text="格式设置", padding="10")
        format_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(format_frame, text="输入格式:").grid(row=0, column=0, sticky=W, pady=5)
        input_combo = ttk.Combobox(format_frame, textvariable=self.input_format, 
                                   values=["auto", "xlsx", "csv", "json", "md"], 
                                   width=12, state="readonly")
        input_combo.grid(row=0, column=1, padx=5, sticky=W)
        ttk.Label(format_frame, text="(auto=从扩展名自动识别)", foreground="gray").grid(row=0, column=2, sticky=W, padx=5)
        
        ttk.Label(format_frame, text="输出格式:").grid(row=1, column=0, sticky=W, pady=5)
        output_combo = ttk.Combobox(format_frame, textvariable=self.output_format, 
                                    values=["auto", "xlsx", "csv", "md"], 
                                    width=12, state="readonly")
        output_combo.grid(row=1, column=1, padx=5, sticky=W)
        ttk.Label(format_frame, text="(auto=从输出文件名识别)", foreground="gray").grid(row=1, column=2, sticky=W, padx=5)
        
        advanced_frame = ttk.LabelFrame(main_frame, text="高级选项", padding="10")
        advanced_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(advanced_frame, text="Excel工作表:").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(advanced_frame, textvariable=self.sheet_name, width=15).grid(row=0, column=1, padx=5, sticky=W)
        ttk.Label(advanced_frame, text="(索引从0开始 或 工作表名称)", foreground="gray").grid(row=0, column=2, sticky=W, padx=5)
        
        ttk.Label(advanced_frame, text="JSON结构:").grid(row=1, column=0, sticky=W, pady=5)
        json_combo = ttk.Combobox(advanced_frame, textvariable=self.json_orient, 
                                  values=["records", "columns"], width=12, state="readonly")
        json_combo.grid(row=1, column=1, padx=5, sticky=W)
        
        ttk.Label(advanced_frame, text="文件编码:").grid(row=2, column=0, sticky=W, pady=5)
        enc_combo = ttk.Combobox(advanced_frame, textvariable=self.encoding, 
                                 values=["utf-8", "gbk", "gb2312", "utf-16"], 
                                 width=12, state="readonly")
        enc_combo.grid(row=2, column=1, padx=5, sticky=W)
        
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=X, pady=(0, 5))
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.pack(fill=X)
        
        log_frame = ttk.LabelFrame(main_frame, text="转换日志", padding="5")
        log_frame.pack(fill=BOTH, expand=True, pady=(5, 0))
        
        self.log_text = Text(log_frame, height=10, wrap=WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="开始转换", command=self.start_conversion, 
                   style="Accent.TButton").pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="清空日志", command=self.clear_log).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="退出程序", command=self.root.quit).pack(side=RIGHT, padx=5)
    
    def select_input(self):
        filetypes = [
            ("所有支持文件", "*.xlsx *.csv *.json *.md"),
            ("Excel文件", "*.xlsx"),
            ("CSV文件", "*.csv"),
            ("JSON文件", "*.json"),
            ("Markdown表格", "*.md"),
            ("所有文件", "*.*")
        ]
        filename = filedialog.askopenfilename(title="选择输入文件", filetypes=filetypes)
        if filename:
            self.input_path.set(filename)
            input_file = Path(filename)
            default_output = str(input_file.with_suffix("")) + "_converted.md"
            self.output_path.set(default_output)
            self.log(f"已选择输入文件: {input_file.name}")
    
    def select_output(self):
        if self.output_format.get() == "auto":
            ext = ".md"
        else:
            ext = "." + self.output_format.get()
        filename = filedialog.asksaveasfilename(title="保存为", defaultextension=ext,
                                                 filetypes=[("所有文件", "*.*")])
        if filename:
            self.output_path.set(filename)
            self.log(f"已设置输出文件: {Path(filename).name}")
    
    def clear_log(self):
        self.log_text.delete(1.0, END)
    
    def log(self, message):
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(END, f"[{timestamp}] {message}\n")
        self.log_text.see(END)
        self.root.update_idletasks()
    
    def start_conversion(self):
        if pd is None:
            messagebox.showerror("依赖缺失", 
                               "未安装 pandas！\n\n请运行以下命令安装依赖:\npip install pandas openpyxl tabulate")
            return
        
        if openpyxl is None:
            messagebox.showerror("依赖缺失", 
                               "未安装 openpyxl！\n\n请运行以下命令安装依赖:\npip install openpyxl")
            return
        
        input_file = self.input_path.get().strip()
        output_file = self.output_path.get().strip()
        
        if not input_file or not output_file:
            messagebox.showwarning("提示", "请选择输入文件和输出文件")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("错误", f"输入文件不存在:\n{input_file}")
            return
        
        in_fmt = self.input_format.get().lower()
        out_fmt = self.output_format.get().lower()
        
        if in_fmt == "auto":
            ext = Path(input_file).suffix.lower()
            map_ext = {'.xlsx': 'xlsx', '.csv': 'csv', '.json': 'json', '.md': 'md'}
            in_fmt = map_ext.get(ext, None)
            if not in_fmt:
                messagebox.showerror("错误", f"无法识别输入文件格式: {ext}")
                return
        
        if out_fmt == "auto":
            ext = Path(output_file).suffix.lower()
            map_ext = {'.xlsx': 'xlsx', '.csv': 'csv', '.md': 'md'}
            out_fmt = map_ext.get(ext, None)
            if not out_fmt:
                messagebox.showerror("错误", 
                                   f"无法识别输出文件格式: {ext}\n\n请手动选择输出格式")
                return
        
        supported = [('xlsx','md'), ('xlsx','csv'), ('csv','md'), ('csv','xlsx'),
                     ('json','md'), ('md','csv'), ('md','xlsx')]
        
        if (in_fmt, out_fmt) not in supported:
            messagebox.showerror("不支持的转换", 
                               f"不支持的转换组合:\n{in_fmt} -> {out_fmt}\n\n支持的转换:\n" +
                               "• xlsx -> md, csv\n• csv -> md, xlsx\n• json -> md\n• md -> csv, xlsx")
            return
        
        sheet = self.sheet_name.get().strip()
        try:
            if sheet.isdigit():
                sheet = int(sheet)
        except:
            pass
        
        orient = self.json_orient.get()
        encoding = self.encoding.get()
        
        self.progress.start(10)
        self.log("=" * 50)
        self.log(f"开始转换: {in_fmt.upper()} -> {out_fmt.upper()}")
        self.log(f"输入文件: {input_file}")
        self.log(f"输出文件: {output_file}")
        
        thread = threading.Thread(target=self._convert_worker,
                                  args=(in_fmt, out_fmt, input_file, output_file,
                                        sheet, orient, encoding))
        thread.daemon = True
        thread.start()
    
    def _convert_worker(self, in_fmt, out_fmt, src, dst, sheet, orient, enc):
        try:
            if in_fmt == 'xlsx' and out_fmt == 'md':
                excel_to_markdown(src, dst, sheet, enc)
            elif in_fmt == 'xlsx' and out_fmt == 'csv':
                excel_to_csv(src, dst, sheet, enc)
            elif in_fmt == 'csv' and out_fmt == 'md':
                csv_to_markdown(src, dst, enc)
            elif in_fmt == 'csv' and out_fmt == 'xlsx':
                csv_to_excel(src, dst, enc)
            elif in_fmt == 'json' and out_fmt == 'md':
                json_to_markdown(src, dst, orient, enc)
            elif in_fmt == 'md' and out_fmt == 'csv':
                markdown_to_csv(src, dst, enc)
            elif in_fmt == 'md' and out_fmt == 'xlsx':
                markdown_to_excel(src, dst, enc)
            
            self.log("✓ 转换成功！")
            self.log(f"输出文件: {dst}")
            self.log("=" * 50)
            
            self.root.after(0, lambda: messagebox.showinfo("转换完成", 
                        f"转换成功！\n\n输出文件:\n{dst}"))
        
        except Exception as e:
            error_msg = f"✗ 转换失败: {str(e)}"
            self.log(error_msg)
            self.log("=" * 50)
            self.root.after(0, lambda: messagebox.showerror("转换失败", str(e)))
        
        finally:
            self.root.after(0, self.progress.stop)


def main():
    root = Tk()
    
    try:
        from tkinter import font
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(size=10)
        root.option_add("*Font", default_font)
    except:
        pass
    
    app = FileConverterGUI(root)
    
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("Accent.TButton", foreground="white", background="#2196F3", 
                    font=('Microsoft YaHei', 10, 'bold'))
    style.map("Accent.TButton", 
              background=[('active', '#1976D2'), ('pressed', '#1565C0')])
    
    root.mainloop()


if __name__ == "__main__":
    main()
