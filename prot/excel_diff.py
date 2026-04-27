import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import xml.etree.ElementTree as ET
import os

class ExcelParser:
    """標準ライブラリのみを使用してxlsxを解析するクラス"""
    
    NAMESPACE = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    @staticmethod
    def get_data(filepath):
        if not os.path.exists(filepath):
            return {}

        data = {} # {(row, col): {'val': ..., 'formula': ...}}
        
        try:
            with zipfile.ZipFile(filepath, 'r') as z:
                # 1. 共有文字列の読み込み
                shared_strings = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    with z.open('xl/sharedStrings.xml') as f:
                        tree = ET.parse(f)
                        for t in tree.findall('.//ns:t', ExcelParser.NAMESPACE):
                            shared_strings.append(t.text if t.text else "")

                # 2. シート1の読み込み
                if 'xl/worksheets/sheet1.xml' in z.namelist():
                    with z.open('xl/worksheets/sheet1.xml') as f:
                        tree = ET.parse(f)
                        for cell in tree.findall('.//ns:c', ExcelParser.NAMESPACE):
                            addr = cell.get('r')
                            c_type = cell.get('t')
                            
                            formula = cell.find('ns:f', ExcelParser.NAMESPACE)
                            formula_text = formula.text if formula is not None else ""
                            
                            value_node = cell.find('ns:v', ExcelParser.NAMESPACE)
                            value = ""
                            if value_node is not None:
                                val_idx = value_node.text
                                if c_type == 's': # 共有文字列参照
                                    value = shared_strings[int(val_idx)]
                                else:
                                    value = val_idx
                            
                            data[addr] = {'val': value, 'fml': formula_text}
            return data
        except Exception as e:
            messagebox.showerror("解析エラー", f"{filepath} の解析に失敗しました:\n{e}")
            return {}

class DiffApp:
    def __init__(self, root):
        root.title("Excel XML Diff Prototype")
        root.geometry("800x500")

        # ファイル選択エリア
        frame = tk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.X)

        self.path1 = tk.StringVar()
        tk.Entry(frame, textvariable=self.path1, width=60).grid(row=0, column=0, padx=5)
        tk.Button(frame, text="ファイル1を選択", command=lambda: self.select_file(self.path1)).grid(row=0, column=1)

        self.path2 = tk.StringVar()
        tk.Entry(frame, textvariable=self.path2, width=60).grid(row=1, column=0, padx=5, pady=5)
        tk.Button(frame, text="ファイル2を選択", command=lambda: self.select_file(self.path2)).grid(row=1, column=1)

        tk.Button(root, text="差分を抽出", command=self.run_diff, bg="#e1e1e1").pack(pady=5)

        # 結果表示エリア
        self.tree = ttk.Treeview(root, columns=("Cell", "File1_Val", "File2_Val", "File1_Fml", "File2_Fml"), show='headings')
        self.tree.heading("Cell", text="セル")
        self.tree.heading("File1_Val", text="値1")
        self.tree.heading("File2_Val", text="値2")
        self.tree.heading("File1_Fml", text="数式1")
        self.tree.heading("File2_Fml", text="数式2")
        
        self.tree.column("Cell", width=50)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def select_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            var.set(path)

    def run_diff(self):
        data1 = ExcelParser.get_data(self.path1.get())
        data2 = ExcelParser.get_data(self.path2.get())

        for i in self.tree.get_children():
            self.tree.delete(i)

        all_cells = sorted(list(set(data1.keys()) | set(data2.keys())))

        for addr in all_cells:
            c1 = data1.get(addr, {'val': '', 'fml': ''})
            c2 = data2.get(addr, {'val': '', 'fml': ''})

            if c1['val'] != c2['val'] or c1['fml'] != c2['fml']:
                self.tree.insert("", tk.END, values=(addr, c1['val'], c2['val'], c1['fml'], c2['fml']))

if __name__ == "__main__":
    root = tk.Tk()
    app = DiffApp(root)
    root.mainloop()