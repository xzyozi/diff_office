import os
import html
from datetime import datetime

def generate_html_report(file1_path, file2_path, diff_data, output_dir):
    """
    差分データから直感的なHTMLレポートを生成する
    diff_data format:
    {
        "Sheet1": [
            {"cell": "A1", "v1": "old", "v2": "new", "f1": "sum()", "f2": "sum()"}
        ]
    }
    """
    # 出力先ディレクトリの確保
    os.makedirs(output_dir, exist_ok=True)
    
    # ファイル名と日時の生成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    f1_name = os.path.basename(file1_path)
    f2_name = os.path.basename(file2_path)
    output_filename = f"diff_report_{timestamp}.html"
    output_path = os.path.join(output_dir, output_filename)

    # --- HTML & CSS テンプレート ---
    html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>Excel Diff Report - {timestamp}</title>
    <style>
        :root {{
            --bg-color: #f6f8fa;
            --border-color: #d0d7de;
            --text-main: #24292f;
            --del-bg: #ffebe9;
            --del-text: #cf222e;
            --add-bg: #e6ffec;
            --add-text: #1a7f37;
            --header-bg: #f3f4f6;
        }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; background-color: var(--bg-color); color: var(--text-main); margin: 0; padding: 20px; line-height: 1.5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: #fff; border: 1px solid var(--border-color); border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); overflow: hidden; }}
        
        .header {{ padding: 20px; border-bottom: 1px solid var(--border-color); background-color: var(--header-bg); }}
        .header h1 {{ margin: 0 0 15px 0; font-size: 20px; }}
        .file-info {{ display: flex; justify-content: space-between; font-size: 14px; }}
        .file-box {{ background: #fff; border: 1px solid var(--border-color); padding: 10px 15px; border-radius: 4px; width: 48%; box-sizing: border-box; }}
        .file-box strong {{ display: block; margin-bottom: 4px; color: #57606a; font-size: 12px; text-transform: uppercase; }}
        
        .sheet-section {{ margin: 20px; border: 1px solid var(--border-color); border-radius: 6px; overflow: hidden; }}
        .sheet-title {{ background-color: var(--header-bg); padding: 12px 16px; font-weight: 600; border-bottom: 1px solid var(--border-color); }}
        
        table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
        th, td {{ padding: 8px 12px; border-bottom: 1px solid var(--border-color); vertical-align: top; word-wrap: break-word; }}
        th {{ text-align: left; background-color: #fafbfc; font-size: 13px; color: #57606a; width: 40%; }}
        th.col-cell {{ width: 10%; text-align: center; }}
        
        .cell-id {{ font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace; font-weight: bold; background: #f0f0f0; padding: 4px 8px; border-radius: 4px; display: inline-block; }}
        
        .diff-block {{ margin-bottom: 4px; }}
        .diff-old {{ background-color: var(--del-bg); color: var(--del-text); text-decoration: line-through; padding: 3px 6px; border-radius: 3px; display: block; margin-bottom: 4px; font-family: monospace; font-size: 13px; }}
        .diff-new {{ background-color: var(--add-bg); color: var(--add-text); padding: 3px 6px; border-radius: 3px; display: block; font-family: monospace; font-size: 13px; }}
        .diff-none {{ color: #8c959f; font-style: italic; font-size: 12px; }}

        /* ツールチップ設定 */
        .tooltip {{ position: relative; display: inline-block; border-bottom: 1px dotted #888; cursor: help; }}
        .tooltip .tooltiptext {{ visibility: hidden; width: 250px; background-color: #333; color: #fff; text-align: center; border-radius: 6px; padding: 8px; position: absolute; z-index: 1; bottom: 125%; left: 50%; margin-left: -125px; opacity: 0; transition: opacity 0.2s; font-size: 12px; font-family: sans-serif; font-style: normal; text-decoration: none; }}
        .tooltip:hover .tooltiptext {{ visibility: visible; opacity: 1; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Excel Diff Report</h1>
            <div class="file-info">
                <div class="file-box">
                    <strong>File 1 (Old)</strong>
                    {html.escape(file1_path)}
                </div>
                <div class="file-box">
                    <strong>File 2 (New)</strong>
                    {html.escape(file2_path)}
                </div>
            </div>
        </div>
"""

    # --- データ部の生成 ---
    for sheet_name, diffs in diff_data.items():
        if not diffs:
            continue
            
        html_content += f"""
        <div class="sheet-section">
            <div class="sheet-title">📄 シート: {html.escape(sheet_name)} ({len(diffs)} 件の差分)</div>
            <table>
                <thead>
                    <tr>
                        <th class="col-cell">セル</th>
                        <th>値 (Value)の差分</th>
                        <th>数式 (Formula)の差分</th>
                    </tr>
                </thead>
                <tbody>
"""
        for d in diffs:
            # HTMLエスケープ（タグなどがExcelに入っていた場合の崩れ防止）
            c_id = html.escape(d['cell'])
            v1 = html.escape(d['v1']) if d['v1'] else '<空>'
            v2 = html.escape(d['v2']) if d['v2'] else '<空>'
            f1 = html.escape(d['f1']) if d['f1'] else '<空>'
            f2 = html.escape(d['f2']) if d['f2'] else '<空>'

            # 値の差分HTML生成
            val_html = ""
            if d['v1'] != d['v2']:
                val_html = f"""
                    <div class="diff-old tooltip">{v1}<span class="tooltiptext">File 1 (Old) の値</span></div>
                    <div class="diff-new tooltip">{v2}<span class="tooltiptext">File 2 (New) の値</span></div>
                """
            else:
                val_html = f'<span class="diff-none">変更なし ({v1})</span>'

            # 数式の差分HTML生成
            fml_html = ""
            if d['f1'] != d['f2']:
                fml_html = f"""
                    <div class="diff-old tooltip">{f1}<span class="tooltiptext">File 1 (Old) の数式</span></div>
                    <div class="diff-new tooltip">{f2}<span class="tooltiptext">File 2 (New) の数式</span></div>
                """
            else:
                fml_html = '<span class="diff-none">変更なし</span>' if not d['f1'] else f'<span class="diff-none">変更なし ({f1})</span>'

            html_content += f"""
                    <tr>
                        <td class="col-cell"><span class="cell-id">{c_id}</span></td>
                        <td>{val_html}</td>
                        <td>{fml_html}</td>
                    </tr>
"""
        
        html_content += """
                </tbody>
            </table>
        </div>
"""

    html_content += """
    </div>
</body>
</html>
"""

    # ファイル書き込み
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    return output_path