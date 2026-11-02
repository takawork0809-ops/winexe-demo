"""
Windows用exe化の検証デモ：請求CSV × 端末管理表(Excel) を電話番号で照合し、
施設別の集計表とエラー一覧をExcelで出力する。

使い方（exe）：
  exeと同じ場所に「入力」フォルダを置き、その中に
    ・請求.csv        （列：電話番号, 金額）
    ・端末管理表.xlsx  （列：電話番号, 施設名）
  を入れて、exeをダブルクリック → 「出力」フォルダに 集計結果.xlsx ができる。
"""
from __future__ import annotations

import sys
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd


def base_dir() -> Path:
    """exeとして動いているときはexeの場所、スクリプトのときはこのファイルの場所。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def normalize_phone(v) -> str:
    """全角→半角、ハイフン・空白・括弧を除去。先頭0が落ちた数値も補正。"""
    if pd.isna(v):
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    if s.endswith(".0"):
        s = s[:-2]
    s = "".join(ch for ch in s if ch.isdigit())
    if len(s) == 10 and not s.startswith("0"):
        s = "0" + s
    return s


def read_csv_any(path: Path) -> pd.DataFrame:
    """KDDIなどのCSVはShift_JIS(cp932)のことが多いので、文字コードを順に試す。"""
    for enc in ("utf-8-sig", "cp932"):
        try:
            return pd.read_csv(path, encoding=enc, dtype=str)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"文字コードを判別できませんでした: {path.name}")


def main() -> int:
    root = base_dir()
    in_dir, out_dir = root / "入力", root / "出力"
    out_dir.mkdir(exist_ok=True)

    bill = read_csv_any(in_dir / "請求.csv")
    dev = pd.read_excel(in_dir / "端末管理表.xlsx", dtype=str)

    bill["キー"] = bill["電話番号"].map(normalize_phone)
    bill["金額"] = pd.to_numeric(bill["金額"].str.replace(",", ""), errors="coerce").fillna(0).astype(int)
    dev["キー"] = dev["電話番号"].map(normalize_phone)

    dup = dev[dev["キー"].duplicated(keep=False) & (dev["キー"] != "")]
    dev_u = dev.drop_duplicates("キー")
    merged = bill.merge(dev_u[["キー", "施設名"]], on="キー", how="left")

    errors = []
    for _, r in merged[merged["施設名"].isna()].iterrows():
        errors.append({"種類": "管理表にない番号", "電話番号": r["電話番号"], "金額": r["金額"],
                       "確認すること": "端末管理表に番号を追加するか、解約済みか確認"})
    for _, r in dup.iterrows():
        errors.append({"種類": "管理表で番号が重複", "電話番号": r["電話番号"], "金額": "",
                       "確認すること": f"施設名「{r['施設名']}」の行を確認"})

    summary = (merged.dropna(subset=["施設名"]).groupby("施設名", as_index=False)["金額"].sum()
               .sort_values("施設名"))
    total_row = pd.DataFrame([{"施設名": "合計", "金額": int(summary["金額"].sum())}])
    summary = pd.concat([summary, total_row], ignore_index=True)

    out = out_dir / "集計結果.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="施設別集計", index=False)
        pd.DataFrame(errors, columns=["種類", "電話番号", "金額", "確認すること"]).to_excel(
            w, sheet_name="エラー一覧", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 22

    check = int(bill["金額"].sum())
    matched = int(summary.iloc[:-1]["金額"].sum())
    print(f"[{datetime.now():%Y-%m-%d %H:%M}] 処理が完了しました")
    print(f"  請求合計: {check:,}円 / 照合できた合計: {matched:,}円 / エラー: {len(errors)}件")
    print(f"  出力: {out}")
    return 0


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
      if stream:
          stream.reconfigure(encoding="utf-8", errors="replace")
    try:
        code = main()
    except Exception:
        print("エラーが発生しました。以下の内容を担当者に送ってください。\n")
        traceback.print_exc()
        code = 1
    if sys.stdin and sys.stdin.isatty():
        input("\nEnterキーを押すと閉じます…")
    sys.exit(code)
