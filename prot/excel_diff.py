import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import xml.etree.ElementTree as ET
import os
import hashlib
import webbrowser

import html_report

class ExcelParser:
    """標準ライブラリのみを使用してxlsx/xlsmを解析するクラス"""
    
    NS_MAIN = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    NS_RELS = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    def __init__(self, filepath):
        self.filepath = filepath
        self.shared_strings = []
        self.sheet_mapping = {}
        self.has_macro = False
        self.macro_hash = None
        self.valid = False
        
        if os.path.exists(filepath):
            try:
                self._load_metadata()
                self.valid = True
            except Exception as e:
                messagebox.showerror("解析エラー", f"{filepath} のメタデータ読み込みに失敗しました:\n{e}")

    def _load_metadata(self):
        with zipfile.ZipFile(self.filepath, 'r') as z:
            # 1. 共有文字列の読み込み
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    for t in tree.findall('.//ns:t', self.NS_MAIN):
                        self.shared_strings.append(t.text if t.text else "")

            # 2. シート名とXMLファイルパスの紐付け
            sheet_rids = {}
            if 'xl/workbook.xml' in z.namelist():
                with z.open('xl/workbook.xml') as f:
                    tree = ET.parse(f)
                    for sheet in tree.findall('.//ns:sheet', self.NS_MAIN):
                        name = sheet.get('name')
                        r_id = sheet.get(f'{{{self.NS_R}}}id')
                        if name and r_id:
                            sheet_rids[r_id] = name

            if 'xl/_rels/workbook.xml.rels' in z.namelist():
                with z.open('xl/_rels/workbook.xml.rels') as f:
                    tree = ET.parse(f)
                    for rel in tree.findall('.//rels:Relationship', self.NS_RELS):
                        r_id = rel.get('Id')
                        if r_id in sheet_rids:
                            target = rel.get('Target')
                            path = target if target.startswith('xl/') else f'xl/{target}'
                            if path.startswith('/'):
                                path = path[1:]
                            self.sheet_mapping[sheet_rids[r_id]] = path

            # 3. マクロ (vbaProject.bin) の存在確認とハッシュ計算
            if 'xl/vbaProject.bin' in z.namelist():
                self.has_macro = True
                with z.open('xl/vbaProject.bin') as f:
                    self.macro_hash = hashlib.md5(f.read()).hexdigest()

    def get_sheet_data(self, sheet_name):
        if sheet_name not in self.sheet_mapping:
            return {}
            
        path = self.sheet_mapping[sheet_name]
        data = {}
        with zipfile.ZipFile(self.filepath, 'r') as z:
            if path in z.namelist():
                with z.open(path) as f:
                    tree = ET.parse(f)
                    for cell in tree.findall('.//ns:c', self.NS_MAIN):
                        addr = cell.get('r')
                        c_type = cell.get('t')
                        
                        formula = cell.find('ns:f', self.NS_MAIN)
                        formula_text = formula.text if formula is not None else ""
                        
                        value_node = cell.find('ns:v', self.NS_MAIN)
                        value = ""
                        if value_node is not None and value_node.text is not None:
                            val_idx = value_node.text
                            if c_type == 's' and val_idx.isdigit():
                                idx = int(val_idx)
                                if idx < len(self.shared_strings):
                                    value = self.shared_strings[idx]
                            else:
                                value = val_idx
                        
                        data[addr] = {'val': value, 'fml': formula_text}
        return data

class DiffApp:
    def __init__(self, root):
        root.title("Excel XML Diff Tool (with HTML Report)")
        root.geometry("900x500")

        frame = tk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.X)

        self.path1 = tk.StringVar()
        tk.Entry(frame, textvariable=self.path1, width=70).grid(row=0, column=0, padx=5)
        tk.Button(frame, text="ファイル1を選択", command=lambda: self.select_file(self.path1)).grid(row=0, column=1)

        self.path2 = tk.StringVar()
        tk.Entry(frame, textvariable=self.path2, width=70).grid(row=1, column=0, padx=5, pady=5)
        tk.Button(frame, text="ファイル2を選択", command=lambda: self.select_file(self.path2)).grid(row=1, column=1)

        tk.Button(root, text="差分を抽出 & ブラウザで開く", command=self.run_diff, bg="#e1e1e1", width=30).pack(pady=5)

        columns = ("Sheet", "Cell", "File1_Val", "File2_Val", "File1_Fml", "File2_Fml")
        self.tree = ttk.Treeview(root, columns=columns, show='headings')
        self.tree.heading("Sheet", text="シート名")
        self.tree.heading("Cell", text="セル")
        self.tree.heading("File1_Val", text="値1")
        self.tree.heading("File2_Val", text="値2")
        self.tree.heading("File1_Fml", text="数式1")
        self.tree.heading("File2_Fml", text="数式2")
        
        self.tree.column("Sheet", width=120)
        self.tree.column("Cell", width=50)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def select_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            var.set(path)

    def run_diff(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        p1, p2 = self.path1.get(), self.path2.get()
        if not p1 or not p2:
            messagebox.showwarning("警告", "比較する2つのファイルを選択してください。")
            return

        if os.path.abspath(p1) == os.path.abspath(p2):
            messagebox.showinfo("確認", "全く同じファイルが選択されています。\n異なるファイルを選択してください。")
            return

        parser1 = ExcelParser(p1)
        parser2 = ExcelParser(p2)

        if not parser1.valid or not parser2.valid:
            return

        if parser1.has_macro and parser2.has_macro:
            if parser1.macro_hash != parser2.macro_hash:
                messagebox.showwarning("マクロ変更検知", "注意: 両ファイル間でマクロ本体に変更が加えられています。")

        # --- シート構成の比較ロジック ---
        sheets1 = set(parser1.sheet_mapping.keys())
        sheets2 = set(parser2.sheet_mapping.keys())
        
        common_sheets = sheets1 & sheets2      # 両方にあるシート（積集合）
        deleted_sheets = sheets1 - sheets2     # File1のみにあるシート（差集合）
        added_sheets = sheets2 - sheets1       # File2のみにあるシート（差集合）

        # 共通シートが1つもない場合はエラー
        if not common_sheets:
            messagebox.showerror("比較エラー", "比較対象となる共通の同名シートが1つも存在しません。")
            return

        diff_count = 0
        report_data = {} 

        # --- 追加・削除されたシートのGUI表示 ---
        for sheet in sorted(list(deleted_sheets)):
            self.tree.insert("", tk.END, values=(sheet, "ALL", "[シート削除]", "-", "-", "-"))
            diff_count += 1
            
        for sheet in sorted(list(added_sheets)):
            self.tree.insert("", tk.END, values=(sheet, "ALL", "-", "[シート追加]", "-", "-"))
            diff_count += 1

        # --- 共通シートのセル差分抽出 ---
        for sheet_name in sorted(list(common_sheets)):
            data1 = parser1.get_sheet_data(sheet_name)
            data2 = parser2.get_sheet_data(sheet_name)

            all_cells = sorted(list(set(data1.keys()) | set(data2.keys())))
            sheet_diffs = []

            for addr in all_cells:
                c1 = data1.get(addr, {'val': '', 'fml': ''})
                c2 = data2.get(addr, {'val': '', 'fml': ''})

                if c1['val'] != c2['val'] or c1['fml'] != c2['fml']:
                    self.tree.insert("", tk.END, values=(sheet_name, addr, c1['val'], c2['val'], c1['fml'], c2['fml']))
                    diff_count += 1
                    
                    sheet_diffs.append({
                        "cell": addr,
                        "v1": c1['val'], "v2": c2['val'],
                        "f1": c1['fml'], "f2": c2['fml']
                    })
            
            if sheet_diffs:
                report_data[sheet_name] = sheet_diffs

        if diff_count == 0:
            messagebox.showinfo("比較完了", "シート構成およびデータに差分は見つかりませんでした。")
        else:
            # HTMLレポートの生成とブラウザ表示
            base_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(base_dir, "..", "output")
            
            try:
                # 修正: added_sheets と deleted_sheets を引数に追加
                html_path = html_report.generate_html_report(
                    p1, p2, report_data, added_sheets, deleted_sheets, output_dir
                )
                
                file_uri = f"file:///{os.path.abspath(html_path).replace(os.sep, '/')}"
                webbrowser.open(file_uri)
                
            except Exception as e:
                messagebox.showerror("HTML生成エラー", f"レポートの生成中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DiffApp(root)
    root.mainloop()