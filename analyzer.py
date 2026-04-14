"""
圖面分析模組 — 使用 Claude Vision API 提取尺寸標注
"""

import base64
import json
import anthropic

SYSTEM_PROMPT = """你是一位專業的工業製圖量測工程師。
你的工作是從工程圖面中識別尺寸標注，並產出結構化的量測檢表資料。

【規則一：橙色/黃色文字尺寸優先——這是最重要的規則】

工程圖面中尺寸標注通常以「黑色」文字顯示。但有時部分尺寸會以「橙色」、「黃色」或「橘色」文字特別標示，代表這些是需要量測的重點尺寸。

請依以下流程處理：
STEP 1：掃描整張圖面，找出所有「橙色/黃色/橘色」文字的尺寸數值（包含尺寸線、公差文字、標注引線上的數值）。
         這些顏色的文字非常醒目，與周圍黑色文字明顯不同。
STEP 2：
  - 若找到任何橙/黃色文字尺寸 → 只列出這些尺寸，一個都不能漏，設 "has_yellow_marks": true
  - 若所有尺寸文字均為黑色    → 列出全部尺寸，設 "has_yellow_marks": false
STEP 3：確認橙/黃色文字尺寸都已找到，再次核對是否有遺漏。

【規則二：位置座標】
對每個識別到的尺寸標注，估算其在圖面影像中的位置（百分比，整數）：
- x: 水平位置，0 = 最左，100 = 最右
- y: 垂直位置，0 = 最頂，100 = 最底

輸出格式必須是純 JSON，不要有任何額外說明文字，格式如下：
{
  "part_name": "零件名稱（從標題欄讀取，若無則填 Unknown）",
  "drawing_no": "圖號（從標題欄讀取，若無則填 N/A）",
  "has_yellow_marks": false,
  "dimensions": [
    {
      "item_no": 1,
      "name": "尺寸名稱/位置描述（中文或英文均可）",
      "nominal": "標稱值（純數字）",
      "unit": "mm",
      "upper_tol": 0.25,
      "lower_tol": -0.25,
      "note": "所在視圖或備註",
      "x": 25,
      "y": 40
    }
  ]
}
"""

def encode_image(image_path: str) -> tuple[str, str]:
    """將圖片編碼為 base64，回傳 (data, media_type)"""
    ext = image_path.lower().split(".")[-1]
    media_type_map = {
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "png": "image/png",
        "gif": "image/gif",
        "webp": "image/webp",
    }
    media_type = media_type_map.get(ext, "image/png")
    with open(image_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, media_type


def analyze_drawing_image(image_path: str, api_key: str = None) -> dict:
    """
    分析單張圖面圖片，回傳結構化尺寸資料

    Args:
        image_path: 圖片路徑（PNG/JPG）
        api_key: Anthropic API Key（None 時從環境變數讀取）

    Returns:
        dict with keys: part_name, drawing_no, dimensions[]
    """
    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    image_data, media_type = encode_image(image_path)

    print(f"  → 正在分析圖面：{image_path}")

    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": image_data,
                        },
                    },
                    {
                        "type": "text",
                        "text": "請分析此工程圖面。重要：先找出所有橙色/黃色/橘色文字的尺寸標注（這些顏色與黑色文字明顯不同）。若有橙/黃色文字的尺寸，只提取這些；若全部為黑色則提取所有尺寸。同時估算每個尺寸在圖面上的位置（x, y，以0-100百分比表示）。輸出純 JSON。",
                    },
                ],
            }
        ],
    ) as stream:
        final_message = stream.get_final_message()

    # 取得文字回應
    text_response = next(
        (block.text for block in final_message.content if block.type == "text"), ""
    )

    # 嘗試解析 JSON（處理可能有 markdown code block 的情況）
    text_response = text_response.strip()
    if text_response.startswith("```"):
        lines = text_response.split("\n")
        # 去掉首尾的 ``` 行
        text_response = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])

    try:
        result = json.loads(text_response)
    except json.JSONDecodeError:
        # 若解析失敗，嘗試找出 JSON 區塊
        start = text_response.find("{")
        end = text_response.rfind("}") + 1
        if start != -1 and end > start:
            result = json.loads(text_response[start:end])
        else:
            result = {
                "part_name": "Unknown",
                "drawing_no": "N/A",
                "dimensions": [],
                "_raw": text_response,
            }

    return result


def analyze_multiple_images(image_paths: list, api_key: str = None) -> dict:
    """
    分析多張圖面（同一零件的多個視圖），合併結果

    Args:
        image_paths: 圖片路徑列表
        api_key: Anthropic API Key

    Returns:
        合併後的 dict
    """
    if not image_paths:
        return {"part_name": "Unknown", "drawing_no": "N/A", "dimensions": []}

    if len(image_paths) == 1:
        return analyze_drawing_image(image_paths[0], api_key)

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    content = []
    for i, path in enumerate(image_paths):
        image_data, media_type = encode_image(path)
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": image_data},
        })
        content.append({
            "type": "text",
            "text": f"（視圖 {i+1}）",
        })

    content.append({
        "type": "text",
        "text": "以上為同一零件的多個視圖。重要：先找出所有橙色/黃色/橘色文字的尺寸標注。若有橙/黃色文字的尺寸，只提取這些（不能漏）；若全部為黑色則提取所有尺寸。同時估算每個尺寸在整體圖面影像上的位置（x, y，0-100百分比）。輸出完整純 JSON。",
    })

    print(f"  → 分析 {len(image_paths)} 張視圖...")

    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}],
    ) as stream:
        final_message = stream.get_final_message()

    text_response = next(
        (block.text for block in final_message.content if block.type == "text"), ""
    ).strip()

    if text_response.startswith("```"):
        lines = text_response.split("\n")
        text_response = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])

    try:
        result = json.loads(text_response)
    except json.JSONDecodeError:
        start = text_response.find("{")
        end = text_response.rfind("}") + 1
        if start != -1 and end > start:
            result = json.loads(text_response[start:end])
        else:
            result = {"part_name": "Unknown", "drawing_no": "N/A", "dimensions": []}

    return result
