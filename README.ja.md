# docx-preflight

**Word を送る前に、リスト設定と共有時の注意点を手元で確認。**

[English](README.md) · [日本語](README.ja.md)

意図しない箇条書きが表示される。原稿に作成者名や古いコメントが残っていないか気になる。**docx-preflight** は、文書を書く人・編集する人・検査を自動化する開発者へ、確認すべき OOXML の構造を一覧にする Python CLI です。正常な箇条書き・番号リストは情報、解決できない設定やプライバシー上の確認事項は警告として表示します。

**Python 3.10以上 · 実行時の追加依存なし · 入力を変更せず検査 · MIT**

原稿の共有前チェックにも、JSON と終了コードを使うローカルの処理にも利用できます。構造を調べるツールで、本文の意味を判断するものではありません。**検査完了は、私的な情報が存在しないという保証ではありません。**

[最短手順](#quick-start) · [出力の読み方](#read-output) · [機能と限界](#checks) · [プライバシー](#privacy) · [診断コード](#diagnostics) · [終了コード](#exit-codes) · [開発と貢献](#contributing)

<a id="quick-start"></a>

## 最短手順

このリポジトリをダウンロードし、そのフォルダでターミナルを開きます。Python が3.10以上であることを確認し、`document.docx` を検査するファイルのパスに置き換えてください。

```sh
python --version
python -m docx_preflight document.docx
```

ソースのフォルダから使う場合、インストールは不要です。空白のあるパスは引用符で囲みます。環境に応じて、例の `python` を `python3` または `py -3` に置き換えてください。

コマンドとしてインストールする場合も、ダウンロードしたソースのフォルダで実行します。

```sh
python -m pip install .
docx-preflight "document.docx"
```

インストール時にはビルド依存の setuptools を取得する場合があります。検査処理には通信機能がありません。PyPI には公開していないため、上記のソースからのインストールを使用してください。

JSON を使う場合や、リストの情報も終了コード1の対象にする場合は次のように実行します。

```sh
python -m docx_preflight document.docx --json
python -m docx_preflight document.docx --fail-on info
python -m docx_preflight --help
```

<a id="read-output"></a>

## 出力の読み方

正常な箇条書きが2段落、値がある作成者プロパティが1項目の合成文書では、次のように表示されます。CLI のメッセージと診断コードは英語です。

```text
DOCX preflight: warnings
Heuristic OOXML checks; not a privacy or rendering guarantee.
info: bullet_list (2)
info: list_paragraph (2)
warning: metadata (1)
```

- 2行の `info` は、同じ2段落を別の観点で示しています。4件の問題という意味ではなく、正常な箇条書きも情報として表示します。
- `metadata (1)` は、値があるプロパティが1項目あるという確認事項です。作成者の値は表示せず、その値が私的な情報かどうかも断定しません。
- `warnings` は、該当する構造を確認する必要がある状態です。この例の既定の終了コードは **1** です。

`count` は各コードに一致した構造の数です。複数のコードが重なる場合があるため、合計を人数や段落数として解釈しないでください。文書全体で集計し、値や場所は表示しません。

<details>
<summary>合成文書でこの例を試す</summary>

次のコードをソースのフォルダに `make_demo.py` として保存し、`python make_demo.py` を実行します。XML 検査用の最小合成 DOCX を作ります。Word の表示レイアウトを検証した文書ではありません。既存の `synthetic-demo.docx` は上書きしません。

```python
from zipfile import ZIP_DEFLATED, ZipFile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/package/2006/content-types"
parts = {
    "[Content_Types].xml": f'<Types xmlns="{C}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
    "_rels/.rels": f'<Relationships xmlns="{R}"><Relationship Id="r1" Type="{O}/officeDocument" Target="word/document.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    "word/document.xml": f'<w:document xmlns:w="{W}"><w:body>' + ''.join('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>' + text + '</w:t></w:r></w:p>' for text in ["First synthetic item", "Second synthetic item"]) + '</w:body></w:document>',
    "word/numbering.xml": f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#x2022;"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>',
    "word/_rels/document.xml.rels": f'<Relationships xmlns="{R}"><Relationship Id="r1" Type="{O}/numbering" Target="numbering.xml"/></Relationships>',
    "docProps/core.xml": '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Synthetic Author</dc:creator></cp:coreProperties>',
}
try:
    with ZipFile("synthetic-demo.docx", "x", ZIP_DEFLATED) as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
except FileExistsError:
    raise SystemExit("The synthetic demo already exists. Choose a new demo filename.")
```

```sh
python -m docx_preflight synthetic-demo.docx
python -m docx_preflight synthetic-demo.docx --json
```

</details>

`--json` の結果は次のデータです。ここでは読みやすく整形しています。

```json
{
  "schema_version": 1,
  "status": "warnings",
  "findings": [
    {"severity": "info", "code": "bullet_list", "count": 2},
    {"severity": "info", "code": "list_paragraph", "count": 2},
    {"severity": "warning", "code": "metadata", "count": 1}
  ],
  "scope": "Heuristic OOXML checks; not a privacy or rendering guarantee."
}
```

<a id="checks"></a>

## 機能と、人が判断すること

| 確認したいこと | ツールが示すもの | 人が確認すること |
| --- | --- | --- |
| 意図しない箇条書きが出る原因の候補 | 有効な番号設定、スタイル継承、空のリスト段落、未解決の定義、既定設定からの継承。 | リストを付ける意図と Word の表示結果。 |
| 原稿に共同編集の情報が残っていないか | 属性、コメント、一部の履歴、共同編集者、外部参照、フィールド、非表示文字の確認事項。 | 実際の文書の内容と、共有してよい情報かどうか。 |
| ローカルの処理に組み込みたい | JSON、固定コードと件数、終了コードのしきい値。 | 検査結果を自分の運用でどう扱うか。 |

文書の既定設定、段落スタイルの継承、段落の直接指定を順に合成します。番号レベルと上書きを解決し、`numId=0` を尊重します。未使用の番号定義だけではリストと判定しません。

箇条書きの修正、メタデータの削除、変更履歴の承認・却下、Word の描画、あらゆる個人情報の発見は行いません。画像メタデータ、OCR、埋め込みバイナリ、本文の意味、視覚的に隠した内容、すべての OOXML 拡張を検査するものではありません。高度な番号スタイル連携やレイアウトは目視確認が必要です。`ok` と `warnings` は検査完了、`error` は検査が不完全な状態です。

<a id="privacy"></a>

## 文書は手元で検査します

検査処理は、文書をアップロードせず、外部リンクを取得せず、Office を起動せず、ZIP をディスクへ展開せず、入力を変更しません。削除・修正コマンドはありません。出力は固定コード・重要度・件数だけで、ファイル名・パス・ZIP 内の名前・作成者の値・コメント本文・リンク先・本文・生 XML を表示しません。入力の拒否や引数エラーでも同様です。ただし、シェルや外部ログがコマンド引数を記録する場合があります。

文書を目視し、必要に応じて別名コピーに [Microsoft のドキュメント検査](https://support.microsoft.com/en-us/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks)も使用してください。そちらにも検出範囲の限界があります。問題報告に私的な文書を添付しないでください。

<a id="diagnostics"></a>

## 診断コード

### 箇条書き・番号リスト

| 診断コード | 重要度 | 意味 |
| --- | --- | --- |
| `list_paragraph` | info | 表示される番号形式を持つ段落。意図したリストか確認してください。 |
| `bullet_list` / `numbered_list` | info | 箇条書き / それ以外の表示番号形式。 |
| `suppressed_numbering` | info | 番号形式が `none`。 |
| `empty_list_paragraph` | warning | 空白以外の文字がない番号付き段落。図だけの項目も該当することがあります。 |
| `default_numbering` | warning | 文書や既定段落スタイルの設定から番号が継承されています。適用範囲を確認してください。 |
| `unresolved_numbering` | warning | 番号 ID・レベルが不正、または定義や番号形式がありません。 |
| `unresolved_style` | warning | スタイルが存在しないか、継承が循環しています。 |
| `unsupported_numbering_link` | warning | 未対応の番号スタイル参照を使用しています。 |

正常な箇条書き・番号リストは情報として扱います。XML から作成者の意図を判断できるわけではありません。Word の描画結果を再現するツールではなく、高度なスタイル連携、スタイル関連付けによる暗黙のレベル選択、互換モード、画像箇条書きやレイアウトの動作は目視確認が必要です。

### プライバシーの確認事項

| 診断コード | 意味 |
| --- | --- |
| `metadata` | `docProps` 内の値がある項目。作成者・最終編集者・日時・アプリ設定・カスタム属性など。無害な値も含みます。 |
| `comments` / `comment_reference` | コメント部品 / コメント参照。一般的な拡張コメント部品も対象。 |
| `tracked_changes` | 挿入・削除・移動・属性変更、および一部の表や番号の変更履歴。 |
| `collaborator_data` / `collaborator_identity` | 一般的な共同編集者部品 / 作成者・イニシャル・ユーザー ID・メールの属性。 |
| `external_relationship` | 外部を参照する関連付け。参照先へアクセスしません。 |
| `field_code` | フィールド命令。パスや外部参照を含む場合があるため、すべて警告します。 |
| `hidden_text` | 有効になっている `vanish` または `webHidden`。 |
| `custom_xml` | カスタム XML 部品。 |
| `uninspected_binary` | 画像や埋め込みなどの非 XML 部品。内部や画像メタデータを検査していません。 |

これらはすべて warning です。警告は確認を促すもので、個人情報の存在を断定するものではありません。

**検出なしでも完全な安全保証にはなりません。** `status=ok` でも、本文・ヘッダー・フッター・任意の拡張 XML・画像・埋め込みなどに私的な情報が残る可能性があります。

<a id="exit-codes"></a>

## 終了コード

| 終了コード | 意味 |
| --- | --- |
| 0 | 指定した重要度以上の検出がない。既定は warning。完全な安全保証ではありません。 |
| 1 | `--fail-on warning` または `--fail-on info` の対象を検出。 |
| 2 | 検査失敗、入力の拒否、引数エラー。検査は不完全であり、問題なしと扱わないでください。 |

`--fail-on none` は警告があっても検査完了時に 0 を返します。検査失敗は常に 2 です。JSON の `status` は重要度しきい値と独立して `ok`、`warnings`、`error` のいずれかです。

<details>
<summary>入力制限とパーサーの対応範囲</summary>

上限を超えた入力や未対応の形式は `status: error`、終了コード2になります。拒否された理由を理解するための制限であり、検査失敗を「問題なし」と扱わないでください。

ZIP は 32 MiB、256 項目、展開後 1 項目 8 MiB、展開後合計 64 MiB、項目ごとの圧縮率 200:1 までです。XML は合計 16 MiB、深さ 100、1 部品あたり 100,000 要素、全部品合計 200,000 要素、1 要素 256 属性、全部品合計 200,000 属性までです。上限超過は検査エラーにします。ZIP の無圧縮と deflate だけに対応し、暗号化、大小文字を無視した重複名、危険な名前、シンボリックリンク、CRC 不一致、壊れた XML、DTD・ENTITY 宣言は拒否します。文字コードを指定する通常の XML 宣言は使用できますが、未知の文字コードは固定診断に変換します。UTF-8・UTF-16 の DTD 拒否をテストしています。内容はメモリ内で読み、展開しません。

`ZipFile` が部品オブジェクトを生成する前に、ZIP32 の中央ディレクトリを 256 KiB 以下に制限し、申告件数が偽装されていても実際のレコードを数えます。サイズ、可変領域の長さ、ローカルヘッダーの位置・名前、ディレクトリ配置を検証します。ZIP64、分割 ZIP、展開必要バージョン 4.5 以上、非標準オフセット、末尾の余分なデータ、曖昧な終了レコードは拒否します。Python の汎用 ZIP リーダーより対応範囲を絞っています。サイズを制限した不変スナップショットを再検証してから読み、検査中の書換えによる大きなディレクトリへの差替えを防ぎます。

ElementTree を構築する前に、名前空間を展開しない Expat 検査で、デコード済み URI を UTF-8 で 2,048 バイト、名前空間宣言を含む要素・属性の修飾名を 512 バイトに制限します。名前空間のスコープ全体をコピーせずに宣言・復元を管理し、ElementTree が展開する名前を事前に見積もります。全部品で保守的な 64 MiB の名前予算を共有し、展開した要素名・属性名または URI を含む名前空間宣言ごとに `4 × (文字数 + 2) + 80` を計上します。同じ名前の繰返しも数えます。短い入力の長い URI が多数の異なる要素名に複製されるメモリ増幅を防ぎます。再宣言・既定名前空間の解除・`xml:` 属性・UTF-16・文字参照を合成テストで確認しています。上限超過は値を表示せず、固定診断にします。

標準 DOCX の content type で、本文が `word/document.xml` にある構成だけに対応します。マクロ文書の本文形式や他の本文位置は拒否します。Transitional と Strict の WordprocessingML 名前空間に対応しますが、OOXML の完全なスキーマ検証は行いません。マルウェア検査や隔離実行環境の代替ではありません。

番号設定は検査 1 回あたり、最初の XML 索引化・段落・スタイル継承の辺を合計 200,000 回まで走査します。未使用・未一致の子要素も数えます。共有する番号定義・レベル・上書き・最初の子要素・文字の有無を一度だけ索引化し、長い参照値や数値の判定も一度だけ処理します。処理量は索引化する XML と実際に辿るスタイルの辺に比例し、段落数と共有定義の大きさの積にはしません。極端に複雑な文書は拒否します。ただし、すべての環境での厳密なメモリ量や実行時間の保証ではありません。

</details>

<a id="contributing"></a>

## 開発と貢献

Python **3.10以上**が必要です。ローカルの検証は Python 3.13 で実行しています。CI 設定は Windows/Linux の Python 3.10・3.13 を対象にしています。ソースの実行とテストに、実行時の追加パッケージは不要です。

```sh
python -m unittest discover -s tests -v
python -m docx_preflight --help
```

テストは一時フォルダで合成 ZIP/XML を生成し、検査前後の入力ハッシュ不変、継承、出力の値抑制、暗号化・破損・サイズ制限、CRC、重複名、XML 宣言、番号解決の線形処理量、中央ディレクトリの偽装件数、ZIP64 の拒否、名前空間・名前の上限、検査中の入力変更を検証します。実文書や個人的なサンプルは同梱していません。CI 設定は Windows/Linux の Python 3.10・3.13 を対象にしています。ZIP32 ヘッダーの解釈は [PKWARE の仕様](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT)を参照し、仕様書自体は同梱していません。

問題報告には、診断コード、期待する動作、Python のバージョン、小さな合成の再現例を記載してください。外部ログの私的なパスは伏せてください。変更を提案する場合は動作を説明し、必要な合成回帰テストを加え、英語・日本語の README を揃えてください。文章の改善や他言語への翻訳も歓迎します。翻訳を追加する際は言語リンクを加えてください。実文書、実際の作成者情報、秘密情報、認証情報を例に含めないでください。

関連する [python-docx](https://github.com/python-openxml/python-docx) は文書作成・編集、[mat2](https://github.com/jvoisin/mat2) は複数形式のメタデータ除去が目的です。本ツールはリスト設定と共有前確認の読み取り専用診断に絞っています。これらからコードやサンプルをコピーしていません。

## ライセンス

[MIT](LICENSE) です。利用条件はライセンスファイルをご確認ください。
