# 休業期間の納期データ更新

## Excelファイル

`excel` フォルダーに、休業期間ごとの入力ファイルがあります。

- `obon.xlsx`
- `goldenWeek.xlsx`
- `yearEndNewYear.xlsx`

各ファイルの「商品データ」シートを編集してください。Excelには商品名と、そのページの表に掲載する納期・期間・注文日だけを記入します。納期や日付は計算式ではなく文字列で入力してください。見出しの変更や、商品の行順の変更・削除はしないでください。カテゴリーや掲載先などのサイト動作用データは、変換時に商品マスタから自動で補います。

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

既存のExcelを今回の掲載項目だけの形式に更新する場合は、次を実行します。

```sh
npm run holiday:templates:refresh
```

このコマンドは現在の商品名と行順を確認し、既存Excelで入力済みの掲載項目を引き継いで不要な列を取り除きます。行順が商品マスタと異なる場合は、安全のため更新を中止します。
