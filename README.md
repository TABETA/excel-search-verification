# Excel 検索・ナレッジ化 検証キット（Claude Cowork 向け）

Claude（Cowork）が Excel ファイルの中身をどこまで検索でき、Excel 群をナレッジとして問い合わせに答えられるかを検証するための、スキル・テストデータ・正解表・検証結果一式。

## 構成

| パス | 内容 |
|---|---|
| `skills/excel-deep-search/` | スキル：フォルダ内の Excel から語句を漏れなく探す（セル・非表示シート・図形・グラフ・コメント・プロパティ・画像内の文字） |
| `skills/excel-knowledge/` | スキル：Excel 群を Markdown ナレッジに変換し、出典つきで問い合わせに回答する |
| `testdata/search-100/` | 検索検証用 Excel 100 ファイル（語句「ペリカン承認」を各所に仕込んだもの） |
| `testdata/knowledge-qa/` | ナレッジ検証用 Excel 19 ファイル（架空の製造業「ミナト精工」の業務文書） |
| `answer-keys/` | 正解表（検索：100 ファイル分、ナレッジ：22 問分） |
| `generators/` | テストデータの生成スクリプト |
| `results/` | 別セッションの Cowork に実行させた結果（原文） |

## スキル

### excel-deep-search
- `scripts/excel_deep_search.py` が XML を直接解析してテキストを拾い、埋め込み画像を一覧画像（contact sheet）にまとめる。
- 画像内の文字は Claude が一覧画像を目視して読む（手順は `SKILL.md`）。
- 完全一致と表記ゆれ（全角/半角、かな、空白・改行・段落による分断）を分けて報告する。暗号化ファイルは「開けない」と報告する。
- 実行例：`python3 excel_deep_search.py <フォルダ> ペリカン承認 --out <出力先> --exclude _検証結果`

### excel-knowledge
- `scripts/excel_to_kb.py build <フォルダ> <ナレッジ先>`：1 Excel = 1 Markdown（表・非表示シート・コメント・図形・矢印のつながり・グラフ・入力規則・プロパティ・画像）と、計算用 CSV を出力する。変更ファイルのみ再処理する。
- `pending` / `set-image`：画像の書き起こしを記入する。書き起こしは画像が変わらない限り再構築後も保持される。
- 回答手順（`SKILL.md`）：索引→grep→該当 Markdown を読む。集計は CSV を計算、新旧（版・改訂履歴・コメント・変更管理）を確認、出典を明記、見つからなければ「記載なし」。

### インストール
各スキルのフォルダ（`SKILL.md` と `scripts/` を含む）を zip にして、Claude の設定のスキル画面からアップロードする。

## 検証結果

### 検索（testdata/search-100、見つけるべき 73 ファイル）

| 条件 | 発見 | 画像内の文字（10） | 誤検出 | 所要時間 |
|---|---|---|---|---|
| 手がかりなしの依頼 | 63 | 0 | 0 | 約 2 分 |
| 図形・画像も含めるよう指示 | 73 | 10 | 0 | 約 3.5 分 |
| excel-deep-search スキル使用 | 73 | 10 | 0 | 約 3 分 |

- テキストとして格納された文字（セル、非表示シート/行/列、白文字、数式結果、コメント、ヘッダー/フッター、入力規則、名前定義、プロパティ、シート名、図形、グループ図形、グラフ、代替テキスト、.xls）は、指示がなくても XML 解析で拾われる。
- 画像内の文字は、指示（またはスキル）がない限り探されない。
- openpyxl のセル走査だけでは、図形・コメント・グラフ・数式結果（`data_only=False` 時）は見つからない。
- スキル使用時は、表記ゆれ（段落分断・ひらがなを含む）を「参考」として分けて報告する。

### ナレッジ回答（testdata/knowledge-qa、22 問）

| 種類 | 正答 |
|---|---|
| 集計（注記の条件を適用） | 3/3 |
| 新旧の判断（改訂履歴・コメント） | 2/2 |
| フロー図（矢印のつながり） | 2/2 |
| 画像内のみの情報 | 2/2 |
| ファイル横断 | 3/3 |
| 回答不能（記載なし） | 2/2 |
| その他（テキストボックス・結合セル・.xls 等） | 8/8 |
| **合計** | **22/22** |

- ナレッジ作成：19 文書、画像 2 枚の書き起こし、作業約 3 分（初回はフォルダアクセス許可待ちが別途発生）。
- 回答：作成済みナレッジのみを根拠に約 1.5 分。全問に出典（ファイル・シート・セル/図形/画像）付き。

## 実行環境の前提
- Cowork のシェルは Windows 上の Linux 仮想マシンで動く。Windows Search の索引・PowerShell・Excel 本体（COM）にはシェルから届かない。
- `.xls` の図形・コメント解析と暗号化テストデータの生成には LibreOffice（soffice）を使う。

## テストデータの再生成

必要なもの：Python 3、`xlsxwriter`、`openpyxl`、`Pillow`、LibreOffice（Python UNO を含む）、Noto Sans/Serif CJK フォント（パスは `generators/images.py` に記載）。

```
cd generators
WORK_DIR=/path/to/work python3 build.py      # 検索用 100 ファイル（xlsx の素）と正解の元データ
WORK_DIR=/path/to/work python3 convert.py    # .xls 変換・暗号化を行い out/final に 100 ファイルを出力
WORK_DIR=/path/to/work python3 verify.py     # XML 生 grep / 段落結合 / openpyxl での検出可否を確認
WORK_DIR=/path/to/work python3 qa_build.py   # ナレッジ検証用 19 ファイルと 22 問の正解
```

乱数シードは固定のため、同じ内容が再生成される（暗号化ファイルのパスワードは `pelican`）。
