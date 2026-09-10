---
title: "Word差分処理詳細"
document_type: "implementation_detailed_design"
version: "1.4"
created_at: "2026-09-10"
updated_at: "2026-09-10"
author: "xzyozi"
purpose: "Word差分GUI、段落抽出、差分整列、表示制御の現行動作を記録する"
related_documents:
  - "IMPL-BD-001"
  - "IMPL-DS-001"
---

# Word差分処理詳細

![Word差分GUIの構成イメージ](images/word-diff-gui.svg)

| 項目           | 内容                                                        |
| :------------- | :---------------------------------------------------------- |
| 管理ID         | IMPL-DD-002                                                 |
| 対象モジュール | `prot/word_diff.py`、`prot/word_link_validator.py`          |
| 入力           | `.docx` または `.docm` の2ファイル                          |
| 出力           | 左右比較GUI、差分強調、差分ジャンプ、リンク整合性バーと詳細 |

## 1. 処理契約

旧・新の2ファイルを選択し、本文段落テキストを比較する。未選択または同一の絶対パスでは比較を中止する。無効なZIP、または `word/document.xml` がない入力は解析エラーとする。

有効な2ファイルでは、段落比較の前に各Wordパッケージのリンク整合性を構造検証する。検証はローカルのOOXML ZIP/XML、Relationship、ブックマークだけを読み取り、ネットワーク接続および外部URLへの到達確認は行わない。リンク検証で確認事項または読込エラーが見つかっても、段落抽出に成功していれば差分比較を継続する。

## 2. 処理フロー

```mermaid
sequenceDiagram
    participant U as 利用者
    participant G as WinMergeStyleApp
    participant P as WordParser
    participant L as WordLinkIntegrityChecker
    participant X as Word package ZIP/XML
    U->>G: 2ファイルを選択して比較
    G->>P: 各ファイルの本文段落を抽出
    P->>X: document XMLを読取
    P-->>G: 段落配列
    G->>L: 各ファイルのリンク構造を検証
    L->>X: XML部品・rels・ブックマークを読取
    L-->>G: LinkIntegrityResult
    G-->>U: リンク整合性バーと詳細を表示
    G->>G: SequenceMatcherで差分を算出
    G-->>U: 左右ペインと差分移動を表示
```

## 3. 段落抽出・リンク検証・表示ルール

| 項目                   | 現行動作                                                                             |
| :--------------------- | :----------------------------------------------------------------------------------- |
| 段落抽出               | `w:body` 直下の本文段落だけを対象に、テキスト要素を順番に連結して末尾空白を除去      |
| 目次除外               | 段落スタイルが `toc` で始まる、または「目次」を含む場合は除外                        |
| フィールド除外         | 複合フィールドの表示結果区間と `w:fldSimple` 配下を抽出対象から除外                  |
| 外部ハイパーリンク     | `w:hyperlink` の `r:id`、Relationship種別、`TargetMode="External"`、Target形式を検証 |
| 文書内アンカー         | `w:hyperlink` の `w:anchor` と `w:bookmarkStart` を照合し、欠損・重複を検出          |
| 内部Relationship・画像 | Word XML部品の `r:id`、`r:embed`、`r:link` とrelsを照合し、ZIP内Targetの存在を検証   |
| 検証結果の表示         | 問題なしは緑、確認事項は黄、読込エラーは赤で表示し、問題時は詳細ダイアログを有効化   |
| 比較単位               | 段落配列を `SequenceMatcher` へ渡す                                                  |
| 整列                   | 挿入・削除・置換で不足側に空行を置き、左右の行を同期                                 |
| 表示                   | 削除、追加、置換、空行を色分けし、前後または番号指定で循環移動                       |

## 4. 失敗・制約

書式、表、画像、ヘッダー・フッター、脚注、コメント、変更履歴は段落差分の比較対象外である。ただし、画像を含むWord XML部品の内部Relationshipは、リンク整合性検証の対象とする。リンク検証は構造確認だけであり、外部URLへの接続・到達確認、リンク先コンテンツの正当性確認は行わない。比較結果はGUI表示のみで、HTML、JSON、差分履歴の出力は行わない。

## 改訂履歴

| 版数    | 改訂日     | 変更者 | 変更内容・理由                                 |
| :------ | :--------- | :----- | :--------------------------------------------- |
| Rev.1.0 | 2026-09-10 | xzyozi | Word差分処理の実装済み動作を記録               |
| Rev.1.1 | 2026-09-10 | xzyozi | 本文段落への対象限定と単純フィールド除外を反映 |
| Rev.1.2 | 2026-09-10 | xzyozi | 現行GUIの構成イメージを追加                    |
| Rev.1.3 | 2026-09-10 | xzyozi | リンク整合性バーを含むGUIイメージへ更新        |
| Rev.1.4 | 2026-09-10 | xzyozi | Wordリンク整合性検証の実装仕様を反映           |
