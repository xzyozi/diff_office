import difflib
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
import zipfile

try:
    from .word_link_validator import validate_word_links
except ImportError:
    from word_link_validator import validate_word_links


class WordParser:
    """標準ライブラリのみを使用してdocx/docmから段落テキストを抽出するクラス（フィールド・目次無視対応版）"""

    NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

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

    def _extract_paragraph_text(self, paragraph):
        texts = []
        in_field_result = False
        namespace = self.NS["w"]
        field_simple_tag = f"{{{namespace}}}fldSimple"
        field_character_tag = f"{{{namespace}}}fldChar"
        field_character_type = f"{{{namespace}}}fldCharType"
        run_tag = f"{{{namespace}}}r"
        text_tag = f"{{{namespace}}}t"

        def visit(element):
            nonlocal in_field_result

            for child in element:
                if child.tag == field_simple_tag:
                    continue
                if child.tag == run_tag:
                    for run_child in child:
                        if run_child.tag == field_character_tag:
                            field_type = run_child.get(field_character_type)
                            if field_type == "separate":
                                in_field_result = True
                            elif field_type == "end":
                                in_field_result = False
                        elif run_child.tag == text_tag and not in_field_result and run_child.text:
                            texts.append(run_child.text)
                else:
                    visit(child)

        visit(paragraph)
        return "".join(texts).rstrip()

    def _load_document(self):
        if not zipfile.is_zipfile(self.filepath):
            raise ValueError("有効なZIP(Office)ファイルではありません。")

        with zipfile.ZipFile(self.filepath, "r") as z:
            if "word/document.xml" not in z.namelist():
                raise ValueError("文書本体 (word/document.xml) が見つかりません。")

            with z.open("word/document.xml") as f:
                tree = ET.parse(f)
                body = tree.find("./w:body", self.NS)
                if body is None:
                    raise ValueError("文書本体 (w:body) が見つかりません。")

                for paragraph in body.findall("./w:p", self.NS):
                    style_node = paragraph.find(".//w:pStyle", self.NS)
                    if style_node is not None:
                        style_value = style_node.get(f"{{{self.NS['w']}}}val", "").lower()
                        if style_value.startswith("toc") or "目次" in style_value:
                            continue

                    self.paragraphs.append(self._extract_paragraph_text(paragraph))


class WinMergeStyleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Document Diff (WinMerge Style)")
        self.root.geometry("1200x760")

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

        tk.Button(
            ctrl_frame, text="比較実行 (Diff)", command=self.run_diff, bg="#e1e1e1", width=15, font=("", 10, "bold")
        ).grid(row=0, column=3, rowspan=2, padx=15)

        self.sync_scroll_var = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl_frame, text="スクロール同期", variable=self.sync_scroll_var).grid(
            row=0, column=4, rowspan=2, padx=5
        )

        # ナビゲーションエリア
        nav_frame = tk.Frame(ctrl_frame)
        nav_frame.grid(row=0, column=5, rowspan=2, padx=15)

        self.btn_prev = tk.Button(nav_frame, text="▲ 前", command=self.prev_diff, state=tk.DISABLED, width=6)
        self.btn_prev.pack(side=tk.LEFT, padx=2)

        # 番号直接入力用のEntryウィジェット
        self.entry_diff_num = tk.Entry(nav_frame, width=5, justify=tk.CENTER)
        self.entry_diff_num.pack(side=tk.LEFT, padx=2)
        self.entry_diff_num.insert(0, "-")
        self.entry_diff_num.bind("<Return>", self.jump_to_specified_diff)

        # 全件数表示用のラベル
        self.lbl_diff_total = tk.Label(nav_frame, text="/ 0")
        self.lbl_diff_total.pack(side=tk.LEFT, padx=2)

        self.btn_next = tk.Button(nav_frame, text="▼ 次", command=self.next_diff, state=tk.DISABLED, width=6)
        self.btn_next.pack(side=tk.LEFT, padx=2)

        # 差分ジャンプ用の状態管理
        self.diff_positions = []  # [GUI論理行番号, ...]
        self.current_diff_idx = -1

        self.link_validation_issues = []
        self.link_status_frame = tk.Frame(root, bd=1, relief=tk.SOLID, bg="#f8fafc")
        self.link_status_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        self.link_status_text = tk.StringVar(value="リンク整合性: 未確認")
        self.link_status_label = tk.Label(
            self.link_status_frame,
            textvariable=self.link_status_text,
            bg="#f8fafc",
            padx=10,
            pady=7,
        )
        self.link_status_label.pack(side=tk.LEFT)
        tk.Label(self.link_status_frame, text="構造確認のみ", bg="#f8fafc", fg="#64748b").pack(side=tk.LEFT)
        self.link_detail_button = tk.Button(
            self.link_status_frame,
            text="詳細を表示",
            command=self.show_link_validation_details,
            state=tk.DISABLED,
        )
        self.link_detail_button.pack(side=tk.RIGHT, padx=8, pady=4)

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

        # 色タグの設定
        self.text_left.tag_configure("delete", background="#ffdddd")
        self.text_right.tag_configure("insert", background="#ddffdd")
        self.text_left.tag_configure("replace", background="#ffebcc")
        self.text_right.tag_configure("replace", background="#ffebcc")
        self.text_left.tag_configure("empty", background="#f0f0f0")
        self.text_right.tag_configure("empty", background="#f0f0f0")

    def _update_link_validation(self, results):
        self.link_validation_issues = [
            (document_label, issue) for document_label, result in results for issue in result.issues
        ]
        checked_references = sum(result.checked_references for _, result in results)
        unreadable = any(result.status == "UNREADABLE" for _, result in results)

        if not self.link_validation_issues:
            text = f"リンク整合性: 問題なし（{checked_references}件を確認）"
            background, foreground = "#dcfce7", "#166534"
            self.link_detail_button.config(state=tk.DISABLED)
        elif unreadable:
            text = f"リンク整合性: {len(self.link_validation_issues)}件の読み込みエラー"
            background, foreground = "#fee2e2", "#b91c1c"
            self.link_detail_button.config(state=tk.NORMAL)
        else:
            text = f"リンク整合性: {len(self.link_validation_issues)}件の確認事項"
            background, foreground = "#fef3c7", "#92400e"
            self.link_detail_button.config(state=tk.NORMAL)

        self.link_status_text.set(text)
        self.link_status_frame.config(bg=background)
        self.link_status_label.config(bg=background, fg=foreground)
        for child in self.link_status_frame.winfo_children():
            if child is not self.link_status_label and child is not self.link_detail_button:
                child.config(bg=background)

    def show_link_validation_details(self):
        if not self.link_validation_issues:
            return

        detail_window = tk.Toplevel(self.root)
        detail_window.title("リンク整合性の詳細")
        detail_window.geometry("1080x360")
        columns = ("document", "code", "part", "target", "detail")
        tree = ttk.Treeview(detail_window, columns=columns, show="headings")
        headings = ("文書", "結果", "参照元", "リンク先・ID", "詳細")
        widths = (100, 220, 200, 220, 300)
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading)
            tree.column(column, width=width, anchor=tk.W)

        for document_label, issue in self.link_validation_issues:
            target = issue.target or issue.reference_id or "-"
            tree.insert("", tk.END, values=(document_label, issue.code, issue.source_part, target, issue.detail))
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # ==========================================
    # スクロール制御
    # ==========================================
    def on_scroll_left(self, *args):
        self.text_left.yview(*args)
        if self.sync_scroll_var.get():
            self.text_right.yview(*args)

    def on_scroll_right(self, *args):
        self.text_right.yview(*args)
        if self.sync_scroll_var.get():
            self.text_left.yview(*args)

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
            self.text_left.yview_scroll(int(-1 * (event.delta / 120)), "units")
            self.text_right.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

    def select_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx *.docm")])
        if path:
            var.set(path)

    # ==========================================
    # 差分ジャンプナビゲーション制御
    # ==========================================
    def prev_diff(self):
        self.jump_to_diff(self.current_diff_idx - 1)

    def next_diff(self):
        self.jump_to_diff(self.current_diff_idx + 1)

    def jump_to_specified_diff(self, event=None):
        if not self.diff_positions:
            return
        try:
            num = int(self.entry_diff_num.get())
            self.jump_to_diff(num - 1)
        except ValueError:
            self._update_counter_display()

    def jump_to_diff(self, index):
        if not self.diff_positions:
            return

        # ループ処理（最後を超えたら最初に戻る）
        total = len(self.diff_positions)
        self.current_diff_idx = index % total

        line_num = self.diff_positions[self.current_diff_idx]
        pos = f"{line_num}.0"

        sync_state = self.sync_scroll_var.get()
        self.sync_scroll_var.set(False)

        self.text_left.see(pos)
        self.text_right.see(pos)

        self._update_counter_display()
        self.sync_scroll_var.set(sync_state)

    def _update_counter_display(self):
        self.entry_diff_num.delete(0, tk.END)
        self.entry_diff_num.insert(0, str(self.current_diff_idx + 1))

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
        if not parser1.valid or not parser2.valid:
            return

        link_results = [
            ("比較元", validate_word_links(p1)),
            ("比較先", validate_word_links(p2)),
        ]
        self._update_link_validation(link_results)

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
        self.entry_diff_num.delete(0, tk.END)
        self.entry_diff_num.insert(0, "-")
        self.lbl_diff_total.config(text="/ 0")

        matcher = difflib.SequenceMatcher(None, paras1, paras2)
        opcodes = matcher.get_opcodes()

        l_line = 1
        r_line = 1
        ui_line = 1

        all_left_text = []
        all_right_text = []

        tags_left = {"replace": [], "delete": [], "empty": []}
        tags_right = {"replace": [], "insert": [], "empty": []}

        for tag, i1, i2, j1, j2 in opcodes:
            p1_sub = paras1[i1:i2]
            p2_sub = paras2[j1:j2]
            max_len = max(len(p1_sub), len(p2_sub))

            # 変更ブロックの先頭をジャンプ用に記録
            if tag != "equal":
                self.diff_positions.append(ui_line)

            for k in range(max_len):
                v1 = p1_sub[k] if k < len(p1_sub) else None
                v2 = p2_sub[k] if k < len(p2_sub) else None

                v1_str = v1 if v1 is not None else ""
                v2_str = v2 if v2 is not None else ""

                prefix1 = f"{l_line:4d} | " if v1 is not None and tag in ("equal", "replace", "delete") else "     | "
                prefix2 = f"{r_line:4d} | " if v2 is not None and tag in ("equal", "replace", "insert") else "     | "

                all_left_text.append(prefix1 + v1_str)
                all_right_text.append(prefix2 + v2_str)

                # 1行ごとの厳密なタグ判定
                ltag, rtag = None, None
                if tag == "delete":
                    ltag, rtag = "delete", "empty"
                elif tag == "insert":
                    ltag, rtag = "empty", "insert"
                elif tag == "replace":
                    if v1 is not None and v2 is not None:
                        if v1 != v2:
                            ltag, rtag = "replace", "replace"
                    elif v1 is not None and v2 is None:
                        ltag, rtag = "delete", "empty"
                    elif v1 is None and v2 is not None:
                        ltag, rtag = "empty", "insert"

                if ltag:
                    tags_left[ltag].append(ui_line)
                if rtag:
                    tags_right[rtag].append(ui_line)

                if v1 is not None and tag in ("equal", "replace", "delete"):
                    l_line += 1
                if v2 is not None and tag in ("equal", "replace", "insert"):
                    r_line += 1
                ui_line += 1

        # 高速一括挿入
        self.text_left.insert(tk.END, "\n".join(all_left_text) + "\n")
        self.text_right.insert(tk.END, "\n".join(all_right_text) + "\n")

        # 高速タグ適用
        for tname, lines in tags_left.items():
            for line_num in lines:
                self.text_left.tag_add(tname, f"{line_num}.0", f"{line_num}.end")

        for tname, lines in tags_right.items():
            for line_num in lines:
                self.text_right.tag_add(tname, f"{line_num}.0", f"{line_num}.end")

        self.text_left.config(state=tk.DISABLED)
        self.text_right.config(state=tk.DISABLED)

        # 完了後のジャンプとUI更新
        if not self.diff_positions:
            messagebox.showinfo("比較完了", "段落テキストに差分は見つかりませんでした。")
        else:
            self.lbl_diff_total.config(text=f"/ {len(self.diff_positions)}")
            self.btn_prev.config(state=tk.NORMAL)
            self.btn_next.config(state=tk.NORMAL)
            self.jump_to_diff(0)  # 最初の差分へジャンプ


if __name__ == "__main__":
    root = tk.Tk()
    app = WinMergeStyleApp(root)
    root.mainloop()
