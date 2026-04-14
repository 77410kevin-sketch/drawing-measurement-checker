"""
PDF 轉圖片模組 — 使用 PyMuPDF 將 PDF 轉為高解析度圖片
"""

import os
import tempfile
import fitz  # PyMuPDF


def pdf_to_images(pdf_path: str, dpi: int = 200, output_dir: str = None) -> list:
    """
    將 PDF 每頁轉換為圖片

    Args:
        pdf_path: PDF 檔案路徑
        dpi: 輸出解析度（預設 200 dpi，對工程圖面已足夠）
        output_dir: 輸出目錄（None 時使用臨時目錄）

    Returns:
        圖片路徑列表（按頁碼排列）
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="drawing_checker_")

    doc = fitz.open(pdf_path)
    image_paths = []
    matrix = fitz.Matrix(dpi / 72, dpi / 72)  # 72 是 PDF 預設 DPI

    print(f"  → PDF 共 {len(doc)} 頁，正在轉換為圖片（{dpi} dpi）...")

    for page_num in range(len(doc)):
        page = doc[page_num]
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)

        image_path = os.path.join(output_dir, f"page_{page_num + 1:03d}.png")
        pixmap.save(image_path)
        image_paths.append(image_path)
        print(f"    頁 {page_num + 1}: {pixmap.width}×{pixmap.height} px → {image_path}")

    doc.close()
    return image_paths


def cleanup_temp_images(image_paths: list):
    """清理臨時圖片檔案"""
    for path in image_paths:
        try:
            os.remove(path)
        except OSError:
            pass
    # 嘗試刪除目錄（若為空）
    if image_paths:
        try:
            os.rmdir(os.path.dirname(image_paths[0]))
        except OSError:
            pass
