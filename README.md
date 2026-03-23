# Excel自動整形・重複削除ツール

Excelファイルを読み込み、自動で整形・重複削除を行うツールです。

## 機能
- Excelファイルの自動読み込み
- 重複データの自動検出・削除
- 空白セルの自動補完・削除
- 整形済みファイルの自動保存

## 使用技術
- Python 3.x
- pandas
- openpyxl

## 使い方
1. リポジトリをクローン
2. 必要なライブラリをインストール
```bash
pip install pandas openpyxl
```
3. Excelファイルをdataフォルダに配置
4. スクリプトを実行
```bash
python main.py
```

## 出力例
- `output/formatted.xlsx` - 整形済みExcelファイル
- `output/duplicates.xlsx` - 削除した重複データ一覧
