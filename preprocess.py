"""
圖面預處理模組 — 使用 PIL + NumPy 偵測橙色/黃色標記尺寸
"""

import io
import base64
from PIL import Image, ImageDraw

try:
    import numpy as np
    _HAS_NUMPY = True
except ImportError:
    _HAS_NUMPY = False


def find_orange_regions(image_path: str) -> tuple[list, str | None]:
    """
    偵測工程圖面中橙色/黃色文字的位置，並在原圖上畫藍色標記框。

    Returns:
        regions: list of {"idx":int, "x1","y1","x2","y2" (pixel), "cx","cy" (0-100 pct)}
        annotated_b64: 帶有藍色框標記的圖片 base64，若未偵測到橙色則 None
    """
    if not _HAS_NUMPY:
        return [], None

    img = Image.open(image_path).convert('RGB')
    W, H = img.size

    arr = np.array(img, dtype=np.int32)
    R = arr[:, :, 0]
    G = arr[:, :, 1]
    B = arr[:, :, 2]

    # 橙色/金黃色文字偵測條件：
    # R 高（文字主色）、G 中等、B 低 → 橙/金色
    # 兼容略偏黃或深橙
    orange_mask = (
        (R > 155) &
        (G > 75)  & (G < 225) &
        (B < 110) &
        (R - G > 15) &   # R 明顯大於 G（排除白/灰）
        (R - B > 90)     # R 遠大於 B（排除藍/紫）
    )

    total_orange = int(orange_mask.sum())
    print(f"  [preprocess] 偵測到橙色像素：{total_orange}")

    if total_orange < 25:
        return [], None

    # ── Grid-based clustering（20×20px 格）────────────────
    CELL = 18
    ys, xs = np.where(orange_mask)
    gx = (xs // CELL).astype(np.int32)
    gy = (ys // CELL).astype(np.int32)
    grid = set(zip(gx.tolist(), gy.tolist()))

    # BFS 找連通區域
    visited = set()
    components = []

    for key in grid:
        if key in visited:
            continue
        comp = []
        queue = [key]
        while queue:
            k = queue.pop()
            if k in visited or k not in grid:
                continue
            visited.add(k)
            comp.append(k)
            kx, ky = k
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    nk = (kx + dx, ky + dy)
                    if nk not in visited:
                        queue.append(nk)
        if len(comp) >= 2:
            components.append(comp)

    print(f"  [preprocess] 找到 {len(components)} 個橙色文字群組")

    # 轉換為像素 bounding box
    PAD = 6
    raw_regions = []
    for comp in components:
        cx_cells = [c[0] for c in comp]
        cy_cells = [c[1] for c in comp]
        x1 = max(0, min(cx_cells) * CELL - PAD)
        y1 = max(0, min(cy_cells) * CELL - PAD)
        x2 = min(W, (max(cx_cells) + 1) * CELL + PAD)
        y2 = min(H, (max(cy_cells) + 1) * CELL + PAD)
        area = (x2 - x1) * (y2 - y1)
        if area < 60:          # 過小雜訊排除
            continue
        if area > W * H * 0.15:  # 過大（整塊橙色背景）排除
            continue
        raw_regions.append({"x1": x1, "y1": y1, "x2": x2, "y2": y2})

    if not raw_regions:
        return [], None

    # 排序：由上到下、由左到右
    raw_regions.sort(key=lambda r: (round((r["y1"] + r["y2"]) / 2 / 15) * 15,
                                    (r["x1"] + r["x2"]) / 2))

    # 合併過於接近的框（中心距 < 40px）
    merged = []
    used = [False] * len(raw_regions)
    for i, ri in enumerate(raw_regions):
        if used[i]:
            continue
        group = [ri]
        used[i] = True
        cxi = (ri["x1"] + ri["x2"]) / 2
        cyi = (ri["y1"] + ri["y2"]) / 2
        for j, rj in enumerate(raw_regions):
            if used[j]:
                continue
            cxj = (rj["x1"] + rj["x2"]) / 2
            cyj = (rj["y1"] + rj["y2"]) / 2
            if ((cxi - cxj) ** 2 + (cyi - cyj) ** 2) ** 0.5 < 40:
                group.append(rj)
                used[j] = True
        x1m = min(g["x1"] for g in group)
        y1m = min(g["y1"] for g in group)
        x2m = max(g["x2"] for g in group)
        y2m = max(g["y2"] for g in group)
        merged.append({"x1": x1m, "y1": y1m, "x2": x2m, "y2": y2m})

    # 再次排序並加上 index + pct 座標
    merged.sort(key=lambda r: (round((r["y1"] + r["y2"]) / 2 / 15) * 15,
                               (r["x1"] + r["x2"]) / 2))
    regions = []
    for idx, r in enumerate(merged):
        regions.append({
            "idx":  idx + 1,
            "x1":   r["x1"], "y1": r["y1"],
            "x2":   r["x2"], "y2": r["y2"],
            "cx":   round((r["x1"] + r["x2"]) / 2 / W * 100, 1),
            "cy":   round((r["y1"] + r["y2"]) / 2 / H * 100, 1),
        })

    print(f"  [preprocess] 合併後：{len(regions)} 個標記框")
    for r in regions:
        print(f"    #{r['idx']}  box=({r['x1']},{r['y1']})-({r['x2']},{r['y2']})  center=({r['cx']}%,{r['cy']}%)")

    # ── 繪製藍色標記框 ────────────────────────────────────
    annotated = img.copy()
    draw = ImageDraw.Draw(annotated)
    for r in regions:
        # 外框（白色加粗，增強對比）
        draw.rectangle([r["x1"] - 2, r["y1"] - 2,
                        r["x2"] + 2, r["y2"] + 2],
                       outline=(255, 255, 255), width=3)
        # 藍色主框
        draw.rectangle([r["x1"], r["y1"], r["x2"], r["y2"]],
                       outline=(0, 140, 255), width=2)
        # 序號標籤（藍底白字）
        lx, ly = r["x1"], max(0, r["y1"] - 16)
        draw.rectangle([lx, ly, lx + 16, ly + 14],
                       fill=(0, 140, 255))
        draw.text((lx + 2, ly + 1), str(r["idx"]),
                  fill=(255, 255, 255))

    buf = io.BytesIO()
    annotated.save(buf, format='PNG')
    ann_b64 = base64.standard_b64encode(buf.getvalue()).decode()

    return regions, ann_b64
