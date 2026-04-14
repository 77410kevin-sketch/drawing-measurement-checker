#!/usr/bin/env python3
"""
圖面識別量測檢表產生器
用法：
  python main.py <圖面檔案>              # 自動輸出 Excel
  python main.py <圖面檔案> -o 輸出.xlsx  # 指定輸出路徑
  python main.py <圖面檔案> --csv        # 輸出 CSV 格式
  python main.py <圖面檔案> --page 2     # 只分析 PDF 第 2 頁
  python main.py <圖面檔案> --all-pages  # 分析 PDF 全部頁（整合）

支援格式：PNG、JPG、JPEG、PDF
"""

import argparse
import os
import sys
import json

# 將專案目錄加入 sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv
from analyzer import analyze_drawing_image, analyze_multiple_images
from exporter import export_to_excel, export_to_csv
from pdf_converter import pdf_to_images, cleanup_temp_images

load_dotenv()


def main():
    parser = argparse.ArgumentParser(
        description="工程圖面識別 → 量測檢表產生器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", help="輸入圖面檔案（PNG/JPG/PDF）")
    parser.add_argument("-o", "--output", help="輸出檔案路徑（預設自動命名）")
    parser.add_argument("--csv", action="store_true", help="輸出 CSV 格式（預設為 Excel）")
    parser.add_argument("--page", type=int, default=None, help="只分析 PDF 的第幾頁（預設第 1 頁）")
    parser.add_argument("--all-pages", action="store_true", help="分析 PDF 全部頁面並整合結果")
    parser.add_argument("--json", action="store_true", help="同時輸出 JSON 原始資料")
    parser.add_argument("--dpi", type=int, default=200, help="PDF 轉圖片解析度（預設 200）")
    parser.add_argument("--api-key", help="Anthropic API Key（可改用環境變數 ANTHROPIC_API_KEY）")

    args = parser.parse_args()
    input_path = args.input

    # 檢查輸入檔案
    if not os.path.exists(input_path):
        print(f"❌ 錯誤：找不到檔案 {input_path}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(input_path)[1].lower()
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")

    if not api_key:
        print("❌ 錯誤：請設定 ANTHROPIC_API_KEY 環境變數或使用 --api-key 參數", file=sys.stderr)
        sys.exit(1)

    # ── 處理輸入 ──────────────────────────────────────────
    temp_images = []

    try:
        if ext == ".pdf":
            print(f"\n📄 輸入：PDF 檔案 → {input_path}")
            all_pages = pdf_to_images(input_path, dpi=args.dpi)
            temp_images = all_pages

            if args.all_pages and len(all_pages) > 1:
                print(f"\n🔍 整合分析全部 {len(all_pages)} 頁...")
                data = analyze_multiple_images(all_pages, api_key)
            else:
                page_num = (args.page or 1) - 1
                if page_num >= len(all_pages):
                    print(f"⚠️  PDF 只有 {len(all_pages)} 頁，改用第 1 頁")
                    page_num = 0
                selected_page = all_pages[page_num]
                print(f"\n🔍 分析第 {page_num + 1} 頁...")
                data = analyze_drawing_image(selected_page, api_key)

        elif ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"):
            print(f"\n🖼️  輸入：圖片 → {input_path}")
            print("\n🔍 分析圖面中...")
            data = analyze_drawing_image(input_path, api_key)

        else:
            print(f"❌ 不支援的格式：{ext}（支援 PNG/JPG/PDF）", file=sys.stderr)
            sys.exit(1)

        # ── 顯示分析結果 ────────────────────────────────────
        part_name = data.get("part_name", "Unknown")
        drawing_no = data.get("drawing_no", "N/A")
        dimensions = data.get("dimensions", [])

        print(f"\n✅ 分析完成！")
        print(f"   零件名稱：{part_name}")
        print(f"   圖    號：{drawing_no}")
        print(f"   尺寸項目：{len(dimensions)} 項")

        if not dimensions:
            print("\n⚠️  未識別到任何尺寸標注，請確認圖面是否清晰。")

        # ── 輸出 JSON（選用） ─────────────────────────────────
        if args.json:
            json_path = args.output.replace(".xlsx", ".json").replace(".csv", ".json") if args.output else f"checklist_raw_{part_name}.json"
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"\n📋 JSON 原始資料：{json_path}")

        # ── 匯出檢表 ───────────────────────────────────────
        if args.csv:
            output_path = export_to_csv(data, args.output)
            print(f"\n📊 CSV 檢表已產生：{output_path}")
        else:
            output_path = export_to_excel(data, args.output)
            print(f"\n📊 Excel 檢表已產生：{output_path}")

        # ── 顯示尺寸預覽 ──────────────────────────────────────
        if dimensions:
            print("\n┌─────────────────────────────────────────────────────────────┐")
            print("│ 尺寸預覽（前 10 項）                                          │")
            print("├──┬───────────────────────────────┬────────┬────────┬────────┤")
            print("│#  │ 量測項目                       │ 標稱值  │ 上公差  │ 下公差  │")
            print("├──┼───────────────────────────────┼────────┼────────┼────────┤")
            for dim in dimensions[:10]:
                no = str(dim.get("item_no", "")).ljust(2)
                name = str(dim.get("name", ""))[:30].ljust(30)
                nom = f"{dim.get('nominal', '')} {dim.get('unit', 'mm')}".ljust(8)
                utol = f"+{dim.get('upper_tol', '')}" if dim.get("upper_tol") is not None else "  -   "
                ltol = str(dim.get("lower_tol", "")) if dim.get("lower_tol") is not None else "  -   "
                print(f"│{no}│ {name}│ {nom}│ {utol:<6} │ {ltol:<6} │")
            if len(dimensions) > 10:
                print(f"│  ╰─ ... 共 {len(dimensions)} 項（請開啟 Excel 檢表查看完整資料）")
            print("└──┴───────────────────────────────┴────────┴────────┴────────┘")

    finally:
        # 清理 PDF 轉換的臨時圖片
        if temp_images:
            cleanup_temp_images(temp_images)


if __name__ == "__main__":
    main()
