import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import xml.etree.ElementTree as ET
import os

class ExcelParser:
    """標準ライブラリのみを使用してxlsxを解析し、シート名ベースでデータを取得するクラス"""
    
    NS_MAIN = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    NS_RELS = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    def __init__(self, filepath):
        self.filepath = filepath
        self.shared_strings = []
        self.sheet_mapping = {} # {シート名: 内部XMLパス}
        self.valid = False
        
        if os.path.exists(filepath):
            try:
                self._load_metadata()
                self.valid = True
            except Exception as e:
                messagebox.showerror("解析エラー", f"{filepath} のメタデータ読み込みに失敗しました:\n{e}")

    def _load_metadata(self):
        """共有文字列とシート名マッピングを初期読み込みする"""
        with zipfile.ZipFile(self.filepath, 'r') as z:
            # 1. 共有文字列の読み込み (ブック全体で1つ)
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    for t in tree.findall('.//ns:t', self.NS_MAIN):
                        self.shared_strings.append(t.text if t.text else "")

            # 2. シート名とXMLファイルパスの紐付けを構築
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
                            # ZIP内のパス形式に合わせて正規化
                            path = target if target.startswith('xl/') else f'xl/{target}'
                            if path.startswith('/'):
                                path = path[1:]
                            self.sheet_mapping[sheet_rids[r_id]] = path

    def get_sheet_data(self, sheet_name):
        """指定されたシート名のセルデータを抽出して返す"""
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
                        
                        # 数式の取得
                        formula = cell.find('ns:f', self.NS_MAIN)
                        formula_text = formula.text if formula is not None else ""
                        
                        # 値の取得
                        value_node = cell.find('ns:v', self.NS_MAIN)
                        value = ""
                        if value_node is not None and value_node.text is not None:
                            val_idx = value_node.text
                            if c_type == 's' and val_idx.isdigit(): # 共有文字列
                                idx = int(val_idx)
                                if idx < len(self.shared_strings):
                                    value = self.shared_strings[idx]
                            else:
                                value = val_idx
                        
                        data[addr] = {'val': value, 'fml': formula_text}
        return data


class DiffApp:
    def __init__(self, root):
        root.title("Excel XML Diff Tool")
        root.geometry("900x500")

        # ファイル選択エリア
        frame = tk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.X)

        self.path1 = tk.StringVar()
        tk.Entry(frame, textvariable=self.path1, width=70).grid(row=0, column=0, padx=5)
        tk.Button(frame, text="ファイル1を選択", command=lambda: self.select_file(self.path1)).grid(row=0, column=1)

        self.path2 = tk.StringVar()
        tk.Entry(frame, textvariable=self.path2, width=70).grid(row=1, column=0, padx=5, pady=5)
        tk.Button(frame, text="ファイル2を選択", command=lambda: self.select_file(self.path2)).grid(row=1, column=1)

        tk.Button(root, text="差分を抽出", command=self.run_diff, bg="#e1e1e1", width=20).pack(pady=5)

        # 結果表示エリア (シート名カラムを追加)
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
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            var.set(path)

    def run_diff(self):
        # Treeviewのクリア
        for i in self.tree.get_children():
            self.tree.delete(i)

        p1, p2 = self.path1.get(), self.path2.get()
        if not p1 or not p2:
            messagebox.showwarning("警告", "比較する2つのファイルを選択してください。")
            return

        # 追加チェック1: そもそも全く同じファイルパスを選択している場合
        if os.path.abspath(p1) == os.path.abspath(p2):
            messagebox.showinfo("確認", "全く同じファイルが選択されています。\n異なるファイルを選択してください。")
            return

        parser1 = ExcelParser(p1)
        parser2 = ExcelParser(p2)

        if not parser1.valid or not parser2.valid:
            return

        # 共通のシート名を取得
        sheets1 = set(parser1.sheet_mapping.keys())
        sheets2 = set(parser2.sheet_mapping.keys())
        common_sheets = sheets1 & sheets2

        if not common_sheets:
            error_msg = (
                "比較可能な同名のシートが存在しません。\n\n"
                f"ファイル1のシート: {', '.join(sheets1) if sheets1 else 'なし'}\n"
                f"ファイル2のシート: {', '.join(sheets2) if sheets2 else 'なし'}"
            )
            messagebox.showerror("比較エラー", error_msg)
            return

        diff_count = 0 # 差分の件数をカウントする変数を追加

        # 共通シートごとに差分を抽出
        for sheet_name in sorted(list(common_sheets)):
            data1 = parser1.get_sheet_data(sheet_name)
            data2 = parser2.get_sheet_data(sheet_name)

            all_cells = sorted(list(set(data1.keys()) | set(data2.keys())))

            for addr in all_cells:
                c1 = data1.get(addr, {'val': '', 'fml': ''})
                c2 = data2.get(addr, {'val': '', 'fml': ''})

                # 値か数式のどちらかに差分があれば表示
                if c1['val'] != c2['val'] or c1['fml'] != c2['fml']:
                    self.tree.insert("", tk.END, values=(sheet_name, addr, c1['val'], c2['val'], c1['fml'], c2['fml']))
                    diff_count += 1

        # 追加チェック2: 比較の結果、差分が1件もなかった場合
        if diff_count == 0:
            messagebox.showinfo("比較完了", "差分は見つかりませんでした。\nファイルの内容（値と数式）は完全に一致しています。")
            
if __name__ == "__main__":
    root = tk.Tk()
    app = DiffApp(root)
    root.mainloop()
