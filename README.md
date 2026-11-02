# Windows exe 化の検証デモ（kddi_demo）

請求CSV × 端末管理表(Excel) を電話番号で照合し、施設別集計とエラー一覧をExcelで出力する小さなツールです。
GitHub Actions の Windows 環境で exe を作り、その場で実行して動作を確認します。

## ファイル
- `main.py` … 本体（Python）
- `sample_data/請求.csv` … 請求データのサンプル（Shift_JIS、電話番号の表記ゆれあり）
- `sample_data/端末管理表.xlsx` … 端末管理表のサンプル（重複・不足をわざと含む）
- `.github/workflows/build-windows.yml` … exe の作成と動作確認の手順

## 期待される結果
- 施設別集計：施設A 4,300円／施設B 4,100円／本部 3,200円／合計 11,600円
- エラー一覧：3件（管理表にない番号 1件、管理表で番号が重複 2件）

## exe の使い方（Windows）
1. `kddi_demo.exe` と同じ場所に「入力」フォルダを作り、`請求.csv` と `端末管理表.xlsx` を入れる
2. `kddi_demo.exe` をダブルクリック
3. 「出力」フォルダに `集計結果.xlsx` ができる
