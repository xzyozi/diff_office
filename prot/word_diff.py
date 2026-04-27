import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import xml.etree.ElementTree as ET
import os
import difflib

class WordParser:
    """標準ライブラリのみを使用してdocx/docmから段落テキストを抽出するクラス"""
    
    NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    def __init__(self, filepath):
        self.filepath = filepath
        self.paragraphs = []
        self.valid = False
        
        if os.path.exists(filepath):
            try:
                self._load_document()
                self.valid = True
            except Exception as e:
                messagebox.showerror("解析エラー", f"{filepath} の読み込みに失敗しました:\n{e}")

    def _load_document(self):
        """word/document.xml を解析し、段落ごとのテキストをリスト化する"""
        if not zipfile.is_zipfile(self.filepath):
            raise ValueError("有効なZIP(Office)ファイルではありません。")

        with zipfile.ZipFile(self.filepath, 'r') as z:
            if 'word/document.xml' in z.namelist():
                with z.open('word/document.xml') as f:
                    tree = ET.parse(f)
                    # <w:p> (段落) ごとに処理
                    for p in tree.findall('.//w:p', self.NS):
                        # 段落内のすべての <w:t> (テキスト) を抽出して結合
                        texts = [t.text for t in p.findall('.//w:t', self.NS) if t.text]
                        text_content = "".join(texts)
                        self.paragraphs.append(text_content)
            else:
                raise ValueError("文書本体 (word/document.xml) が見つかりません。")

class WinMergeStyleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Document Diff (WinMerge Style)")
        self.root.geometry("1200x700")

        # --- コントロールパネル (上部) ---
        ctrl_frame = tk.Frame(root)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=10)

        # ファイル1
        tk.Label(ctrl_frame, text="左 (Old):").grid(row=0, column=0, sticky="e")
        self.path1 = tk.StringVar()
        tk.Entry(ctrl_frame, textvariable=self.path1, width=60).grid(row=0, column=1, padx=5)
        tk.Button(ctrl_frame, text="参照...", command=lambda: self.select_file(self.path1)).grid(row=0, column=2)

        # ファイル2
        tk.Label(ctrl_frame, text="右 (New):").grid(row=1, column=0, sticky="e", pady=5)
        self.path2 = tk.StringVar()
        tk.Entry(ctrl_frame, textvariable=self.path2, width=60).grid(row=1, column=1, padx=5)
        tk.Button(ctrl_frame, text="参照...", command=lambda: self.select_file(self.path2)).grid(row=1, column=2)

        tk.Button(ctrl_frame, text="比較実行 (Diff)", command=self.run_diff, bg="#e1e1e1", width=20, font=("", 10, "bold")).grid(row=0, column=3, rowspan=2, padx=20)

        # --- メインビュー (左右分割テキストエリア) ---
        main_pane = tk.PanedWindow(root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # スクロールバー (共有)
        self.scrollbar = tk.Scrollbar(root)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 左テキストエリア
        self.text_left = tk.Text(main_pane, wrap=tk.NONE, yscrollcommand=self.scrollbar.set, font=("MS Gothic", 11))
        main_pane.add(self.text_left, stretch="always")

        # 右テキストエリア
        self.text_right = tk.Text(main_pane, wrap=tk.NONE, yscrollcommand=self.scrollbar.set, font=("MS Gothic", 11))
        main_pane.add(self.text_right, stretch="always")

        # スクロールバーの連動設定
        self.scrollbar.config(command=self.sync_yview)
        
        # マウスホイールの連動設定
        self.text_left.bind("<MouseWheel>", self.sync_mousewheel)
        self.text_right.bind("<MouseWheel>", self.sync_mousewheel)

        # --- 色タグの設定 ---
        # 削除・追加・変更の背景色設定
        self.text_left.tag_configure("delete", background="#ffdddd")  # 薄い赤
        self.text_right.tag_configure("insert", background="#ddffdd") # 薄い緑
        self.text_left.tag_configure("replace", background="#ffebcc") # 薄いオレンジ
        self.text_right.tag_configure("replace", background="#ffebcc") 
        self.text_left.tag_configure("empty", background="#f0f0f0")   # グレー(空白補完用)
        self.text_right.tag_configure("empty", background="#f0f0f0")

    def sync_yview(self, *args):
        """スクロールバーの動きを左右のテキストウィジェットに同期"""
        self.text_left.yview(*args)
        self.text_right.yview(*args)

    def sync_mousewheel(self, event):
        """マウスホイールの動きを左右に同期"""
        self.text_left.yview_scroll(int(-1*(event.delta/120)), "units")
        self.text_right.yview_scroll(int(-1*(event.delta/120)), "units")
        return "break"

    def select_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.docm")])
        if path:
            var.set(path)

    def insert_line(self, text_widget, line_num, content, tag=None):
        """テキストウィジェットに行番号付きでテキストを挿入する"""
        text_widget.config(state=tk.NORMAL)
        line_prefix = f"{line_num:4d} | " if line_num else "     | "
        
        # 行番号部分の挿入
        text_widget.insert(tk.END, line_prefix)
        # 本文の挿入
        start_idx = text_widget.index(tk.END)
        text_widget.insert(tk.END, content + "\n")
        end_idx = text_widget.index(tk.END)
        
        # タグ（背景色）の適用
        if tag:
            text_widget.tag_add(tag, f"{start_idx} -1c", f"{end_idx} -1c")
            
        text_widget.config(state=tk.DISABLED)

    def run_diff(self):
        p1, p2 = self.path1.get(), self.path2.get()
        if not p1 or not p2:
            messagebox.showwarning("警告", "比較する2つのファイルを選択してください。")
            return

        if os.path.abspath(p1) == os.path.abspath(p2):
            messagebox.showinfo("確認", "全く同じファイルが選択されています。")
            return

        # Wordファイルの解析
        parser1 = WordParser(p1)
        parser2 = WordParser(p2)

        if not parser1.valid or not parser2.valid:
            return

        paras1 = parser1.paragraphs
        paras2 = parser2.paragraphs

        # テキストエリアのクリア
        self.text_left.config(state=tk.NORMAL)
        self.text_right.config(state=tk.NORMAL)
        self.text_left.delete(1.0, tk.END)
        self.text_right.delete(1.0, tk.END)
        self.text_left.config(state=tk.DISABLED)
        self.text_right.config(state=tk.DISABLED)

        # 差分比較の実行
        matcher = difflib.SequenceMatcher(None, paras1, paras2)
        opcodes = matcher.get_opcodes()

        diff_count = 0
        l_line = 1
        r_line = 1

        # WinMergeのような左右の行位置合わせ処理
        for tag, i1, i2, j1, j2 in opcodes:
            p1_sub = paras1[i1:i2]
            p2_sub = paras2[j1:j2]
            
            # ブロックの最大行数を取得（左右を揃えるため）
            max_len = max(len(p1_sub), len(p2_sub))

            for k in range(max_len):
                val1 = p1_sub[k] if k < len(p1_sub) else ""
                val2 = p2_sub[k] if k < len(p2_sub) else ""

                if tag == 'equal':
                    self.insert_line(self.text_left, l_line, val1)
                    self.insert_line(self.text_right, r_line, val2)
                    l_line += 1
                    r_line += 1
                
                elif tag == 'replace':
                    self.insert_line(self.text_left, l_line if k < len(p1_sub) else None, val1, "replace" if k < len(p1_sub) else "empty")
                    self.insert_line(self.text_right, r_line if k < len(p2_sub) else None, val2, "replace" if k < len(p2_sub) else "empty")
                    if k < len(p1_sub): l_line += 1
                    if k < len(p2_sub): r_line += 1
                    diff_count += 1
                    
                elif tag == 'delete':
                    self.insert_line(self.text_left, l_line, val1, "delete")
                    self.insert_line(self.text_right, None, "", "empty") # 右側は空行で位置合わせ
                    l_line += 1
                    diff_count += 1
                    
                elif tag == 'insert':
                    self.insert_line(self.text_left, None, "", "empty")  # 左側は空行で位置合わせ
                    self.insert_line(self.text_right, r_line, val2, "insert")
                    r_line += 1
                    diff_count += 1

        if diff_count == 0:
            messagebox.showinfo("比較完了", "段落テキストに差分は見つかりませんでした。")

if __name__ == "__main__":
    root = tk.Tk()
    app = WinMergeStyleApp(root)
    root.mainloop()