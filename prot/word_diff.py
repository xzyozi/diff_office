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

        # ファイル選択エリア
        tk.Label(ctrl_frame, text="左 (Old):").grid(row=0, column=0, sticky="e")
        self.path1 = tk.StringVar()
        tk.Entry(ctrl_frame, textvariable=self.path1, width=45).grid(row=0, column=1, padx=5)
        tk.Button(ctrl_frame, text="参照...", command=lambda: self.select_file(self.path1)).grid(row=0, column=2)

        tk.Label(ctrl_frame, text="右 (New):").grid(row=1, column=0, sticky="e", pady=5)
        self.path2 = tk.StringVar()
        tk.Entry(ctrl_frame, textvariable=self.path2, width=45).grid(row=1, column=1, padx=5)
        tk.Button(ctrl_frame, text="参照...", command=lambda: self.select_file(self.path2)).grid(row=1, column=2)

        # 実行ボタン
        tk.Button(ctrl_frame, text="比較実行 (Diff)", command=self.run_diff, bg="#e1e1e1", width=15, font=("", 10, "bold")).grid(row=0, column=3, rowspan=2, padx=15)

        # 同期スクロールのトグルスイッチ
        self.sync_scroll_var = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl_frame, text="スクロール同期", variable=self.sync_scroll_var).grid(row=0, column=4, rowspan=2, padx=5)

        # 【追加】差分ジャンプナビゲーション
        nav_frame = tk.Frame(ctrl_frame)
        nav_frame.grid(row=0, column=5, rowspan=2, padx=15)

        self.btn_prev = tk.Button(nav_frame, text="▲ 前", command=self.prev_diff, state=tk.DISABLED, width=6)
        self.btn_prev.pack(side=tk.LEFT, padx=2)

        self.lbl_diff_count = tk.Label(nav_frame, text="0 / 0", width=10, bg="white", relief=tk.SUNKEN)
        self.lbl_diff_count.pack(side=tk.LEFT, padx=5)

        self.btn_next = tk.Button(nav_frame, text="▼ 次", command=self.next_diff, state=tk.DISABLED, width=6)
        self.btn_next.pack(side=tk.LEFT, padx=2)

        # 差分ジャンプ用の状態管理
        self.diff_positions = []    # [(left_idx, right_idx), ...]
        self.current_diff_idx = -1

        # --- メインビュー (左右分割テキストエリア) ---
        main_pane = tk.PanedWindow(root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self._syncing = False

        # 左ペイン
        left_frame = tk.Frame(main_pane)
        self.text_left = tk.Text(left_frame, wrap=tk.NONE, font=("MS Gothic", 11))
        self.scroll_left = tk.Scrollbar(left_frame, command=self.on_scroll_left)
        self.text_left.config(yscrollcommand=self.set_scroll_left)
        self.text_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_left.pack(side=tk.RIGHT, fill=tk.Y)
        main_pane.add(left_frame, stretch="always")

        # 右ペイン
        right_frame = tk.Frame(main_pane)
        self.text_right = tk.Text(right_frame, wrap=tk.NONE, font=("MS Gothic", 11))
        self.scroll_right = tk.Scrollbar(right_frame, command=self.on_scroll_right)
        self.text_right.config(yscrollcommand=self.set_scroll_right)
        self.text_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_right.pack(side=tk.RIGHT, fill=tk.Y)
        main_pane.add(right_frame, stretch="always")
        
        self.text_left.bind("<MouseWheel>", self.sync_mousewheel)
        self.text_right.bind("<MouseWheel>", self.sync_mousewheel)

        # --- 色タグの設定 ---
        self.text_left.tag_configure("delete", background="#ffdddd")  
        self.text_right.tag_configure("insert", background="#ddffdd") 
        self.text_left.tag_configure("replace", background="#ffebcc") 
        self.text_right.tag_configure("replace", background="#ffebcc") 
        self.text_left.tag_configure("empty", background="#f0f0f0")   
        self.text_right.tag_configure("empty", background="#f0f0f0")

    # ==========================================
    # スクロール・ナビゲーション制御
    # ==========================================
    def on_scroll_left(self, *args):
        self.text_left.yview(*args)
        if self.sync_scroll_var.get(): self.text_right.yview(*args)

    def on_scroll_right(self, *args):
        self.text_right.yview(*args)
        if self.sync_scroll_var.get(): self.text_left.yview(*args)

    def set_scroll_left(self, *args):
        self.scroll_left.set(*args)
        if self.sync_scroll_var.get() and not self._syncing:
            self._syncing = True
            self.text_right.yview_moveto(args[0])
            self._syncing = False

    def set_scroll_right(self, *args):
        self.scroll_right.set(*args)
        if self.sync_scroll_var.get() and not self._syncing:
            self._syncing = True
            self.text_left.yview_moveto(args[0])
            self._syncing = False

    def sync_mousewheel(self, event):
        if self.sync_scroll_var.get():
            self.text_left.yview_scroll(int(-1*(event.delta/120)), "units")
            self.text_right.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"

    def select_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.docm")])
        if path: var.set(path)

    # 【追加】前の差分へジャンプ
    def prev_diff(self):
        if not self.diff_positions: return
        self.current_diff_idx = (self.current_diff_idx - 1) % len(self.diff_positions)
        self.jump_to_current_diff()

    # 【追加】次の差分へジャンプ
    def next_diff(self):
        if not self.diff_positions: return
        self.current_diff_idx = (self.current_diff_idx + 1) % len(self.diff_positions)
        self.jump_to_current_diff()

    # 【追加】指定した差分インデックスへスクロール
    def jump_to_current_diff(self):
        l_idx, r_idx = self.diff_positions[self.current_diff_idx]
        
        # 左右が勝手に連動して表示位置が狂うのを防ぐため、一時的に同期をオフにする
        sync_state = self.sync_scroll_var.get()
        self.sync_scroll_var.set(False)

        # 対象の行が見える位置へジャンプ
        self.text_left.see(l_idx)
        self.text_right.see(r_idx)

        # ラベルの更新
        self.lbl_diff_count.config(text=f"{self.current_diff_idx + 1} / {len(self.diff_positions)}")
        
        self.sync_scroll_var.set(sync_state)

    # ==========================================
    # 差分比較実行メソッド
    # ==========================================
    def run_diff(self):
        p1, p2 = self.path1.get(), self.path2.get()
        if not p1 or not p2:
            messagebox.showwarning("警告", "比較する2つのファイルを選択してください。")
            return
        if os.path.abspath(p1) == os.path.abspath(p2):
            messagebox.showinfo("確認", "全く同じファイルが選択されています。")
            return

        parser1 = WordParser(p1)
        parser2 = WordParser(p2)
        if not parser1.valid or not parser2.valid: return

        paras1 = parser1.paragraphs
        paras2 = parser2.paragraphs

        self.text_left.config(state=tk.NORMAL)
        self.text_right.config(state=tk.NORMAL)
        self.text_left.delete(1.0, tk.END)
        self.text_right.delete(1.0, tk.END)

        # 状態の初期化
        self.diff_positions.clear()
        self.current_diff_idx = -1
        self.btn_prev.config(state=tk.DISABLED)
        self.btn_next.config(state=tk.DISABLED)
        self.lbl_diff_count.config(text="0 / 0")

        matcher = difflib.SequenceMatcher(None, paras1, paras2)
        opcodes = matcher.get_opcodes()

        l_line = 1
        r_line = 1

        for tag, i1, i2, j1, j2 in opcodes:
            p1_sub = paras1[i1:i2]
            p2_sub = paras2[j1:j2]
            max_len = max(len(p1_sub), len(p2_sub))
            
            left_block_text = ""
            right_block_text = ""

            left_start_idx = self.text_left.index(tk.END)
            right_start_idx = self.text_right.index(tk.END)

            # 【追加】変更ブロックの開始位置を記録
            if tag != 'equal':
                self.diff_positions.append((left_start_idx, right_start_idx))

            for k in range(max_len):
                val1 = p1_sub[k] if k < len(p1_sub) else ""
                val2 = p2_sub[k] if k < len(p2_sub) else ""

                prefix1 = f"{l_line:4d} | " if k < len(p1_sub) and tag in ('equal', 'replace', 'delete') else "     | "
                prefix2 = f"{r_line:4d} | " if k < len(p2_sub) and tag in ('equal', 'replace', 'insert') else "     | "

                left_block_text += prefix1 + val1 + "\n"
                right_block_text += prefix2 + val2 + "\n"

                if k < len(p1_sub) and tag in ('equal', 'replace', 'delete'): l_line += 1
                if k < len(p2_sub) and tag in ('equal', 'replace', 'insert'): r_line += 1

            self.text_left.insert(tk.END, left_block_text)
            self.text_right.insert(tk.END, right_block_text)

            left_end_idx = self.text_left.index(f"{tk.END}-1c")
            right_end_idx = self.text_right.index(f"{tk.END}-1c")

            if tag == 'replace':
                self.text_left.tag_add("replace", left_start_idx, left_end_idx)
                self.text_right.tag_add("replace", right_start_idx, right_end_idx)
            elif tag == 'delete':
                self.text_left.tag_add("delete", left_start_idx, left_end_idx)
                self.text_right.tag_add("empty", right_start_idx, right_end_idx)
            elif tag == 'insert':
                self.text_left.tag_add("empty", left_start_idx, left_end_idx)
                self.text_right.tag_add("insert", right_start_idx, right_end_idx)

        self.text_left.config(state=tk.DISABLED)
        self.text_right.config(state=tk.DISABLED)

        # 完了後のジャンプ処理
        if not self.diff_positions:
            messagebox.showinfo("比較完了", "段落テキストに差分は見つかりませんでした。")
        else:
            self.btn_prev.config(state=tk.NORMAL)
            self.btn_next.config(state=tk.NORMAL)
            self.next_diff() # 自動的に最初の差分へジャンプ
            
if __name__ == "__main__":
    root = tk.Tk()
    app = WinMergeStyleApp(root)
    root.mainloop()