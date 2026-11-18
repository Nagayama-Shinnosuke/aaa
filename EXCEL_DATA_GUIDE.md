# 休業期間の納期データ更新

## Excelファイル

`excel` フォルダーに、休業期間ごとの入力ファイルがあります。

- `obon.xlsx`
- `goldenWeek.xlsx`
- `yearEndNewYear.xlsx`

各ファイルの「商品データ」シートを編集してください。納期や日付は計算式ではなく文字列で記入します。列の追加・削除や見出しの変更はしないでください。年末年始の「締切時間（JSON保持用）」はサイトには表示されませんが、出力JSONに必要なため列を残してください。

## JSONの更新

Node.jsをインストール済みの環境で、プロジェクトのルートから初回のみ依存関係をインストールします。

```sh
npm install
```

Excelを更新した後、以下を実行します。

```sh
npm run holiday:convert
```

次のファイルが更新され、サイトが読み込むJSONになります。

- `products-obon.json`
- `products-goldenWeek.json`
- `products-yearEndNewYear.json`

3つのExcelをすべて検証してからJSONを書き出します。入力エラーが出た場合は内容を修正し、もう一度実行してください。

初期Excelを作り直す必要がある場合は、既存ファイルを退避したうえで `npm run holiday:templates` を実行します。この処理は `products.json` の現在の内容を元にし、既存Excelを上書きしません。
