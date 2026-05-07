from __future__ import annotations

import base64
import json
import os
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Protocol


DEFAULT_OCR_PROVIDER = "disabled"
ALIYUN_QWEN_OCR_MODEL = "qwen-vl-ocr"
QWEN25_VL_3B_MODEL = "qwen2.5-vl-3b-instruct"
QWEN25_VL_7B_MODEL = "qwen2.5-vl-7b-instruct"
INTERNVL2_8B_MODEL = "internvl2-8b"
ALIYUN_COMPATIBLE_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
ZHIPU_OCR_MODEL = "glm-ocr"
ZHIPU_LAYOUT_PARSING_URL = "https://open.bigmodel.cn/api/paas/v4/layout_parsing"
SUPPORTED_PROVIDERS = {
    "disabled",
    "paddle_local",
    "aliyun_ocr",
    "zhipu_ocr",
    "baidu_ocr",
    "tencent_ocr",
    "qwen25_vl_3b",
    "qwen25_vl_7b",
    "internvl2_8b",
}

OCR_TEXT_PROMPT = """請對這張圖片做 OCR 文字辨識。

規則：
- 只輸出圖片中實際看得到的文字，不要推測、不要補全。
- 盡量依照畫面從上到下、從左到右排序。
- 公司名稱請完整保留。
- 如果文字不確定，也保留原文，不要改寫。

只輸出 JSON array，不要說明文字或 markdown：
[{"text":"辨識到的文字"}]
"""

OCR_PROMPT_STRICT_V2 = """你是 OCR 引擎，不是摘要器。

任務：
- 只擷取圖片中實際看得到的文字，不要推測、不要補全、不要翻譯。
- 依畫面順序輸出（由上到下、由左到右）。

輸出格式（必須完全遵守）：
- 只能輸出一個合法 JSON Array。
- 每個元素只能是 {"text":"..."}。
- 不得輸出 markdown、註解、前言、結語、額外欄位。
- 不得輸出包裝殘片，如 [, ], {, }（除非它們是圖片裡真的字元，且要放在 text 值中）。

輸出前自檢：
- 你要確認內容是合法 JSON。
- 若無法保證合法 JSON，請輸出 []。

只輸出 JSON：
[{"text":"辨識到的文字"}]
"""

OCR_PROMPT_STRICT_V3 = """你正在執行高嚴格 OCR 結構化輸出。

要求：
1) 僅抄錄圖片可見文字，禁止推測。
2) 每段文字只輸出一次，保持閱讀順序。
3) 公司名稱要完整保留（含括號與大小寫）。

格式要求（違反即視為失敗）：
- 回覆只能是一個 JSON Array。
- Array 內每筆格式固定：{"text":"..."}。
- 不允許任何額外鍵值，不允許程式碼框，不允許自然語言解釋。
- 不允許尾逗號，不允許單引號，不允許不完整 JSON。

若圖片無可讀文字或你無法保證 JSON 正確，輸出 []。

只輸出 JSON：
[{"text":"辨識到的文字"}]
"""

OCR_PROMPT_PROFILES = {
    "default": OCR_TEXT_PROMPT,
    "strict_v2": OCR_PROMPT_STRICT_V2,
    "strict_v3": OCR_PROMPT_STRICT_V3,
}


class OcrProvider(Protocol):
    name: str

    def recognize(self, image_path: Path, prompt_text: str | None = None) -> dict:
        ...


def _company_candidates(items: list[dict]) -> list[dict]:
    suffixes = ("公司", "集团", "集團", "有限公司", "limited", "ltd", "inc", "corp", "corporation", "pte", "bhd")
    return [
        item for item in items
        if any(suffix in item["text"].lower() for suffix in suffixes)
    ]


def _build_result(provider: str, items: list[dict]) -> dict:
    candidates = _company_candidates(items)
    return {
        "provider": provider,
        "engine": provider,
        "text_count": len(items),
        "company_candidate_count": len(candidates),
        "items": items,
        "company_candidates": candidates,
    }


class DisabledOcrProvider:
    name = "disabled"

    def recognize(self, image_path: Path, prompt_text: str | None = None) -> dict:
        raise RuntimeError("OCR provider 尚未啟用。請設定 OCR_PROVIDER=paddle_local 或接入雲端 OCR provider。")


class PlaceholderCloudOcrProvider:
    def __init__(self, name: str):
        self.name = name

    def recognize(self, image_path: Path, prompt_text: str | None = None) -> dict:
        raise RuntimeError(f"{self.name} provider 尚未接入。此版本只預留介面，尚未呼叫雲端 OCR API。")


class PaddleLocalOcrProvider:
    name = "paddle_local"

    def recognize(self, image_path: Path, prompt_text: str | None = None) -> dict:
        try:
            from paddleocr import PaddleOCR
        except Exception as exc:  # pragma: no cover - depends on optional heavy package
            raise RuntimeError(
                "PaddleOCR 尚未安裝。先執行 `python3 -m pip install -r requirements-paddle.txt` 後再測試。"
            ) from exc

        try:
            ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
            raw = ocr.ocr(str(image_path), cls=True)
        except TypeError:
            # PaddleOCR v3 has renamed several parameters.
            ocr = PaddleOCR(
                lang="ch",
                use_doc_orientation_classify=False,
                use_doc_unwarping=False,
                use_textline_orientation=True,
            )
            raw = ocr.ocr(str(image_path))

        items = _flatten_paddle_items(raw)
        return _build_result(self.name, items)


class AliyunQwenOcrProvider:
    name = "aliyun_ocr"
    model_name = ALIYUN_QWEN_OCR_MODEL

    def recognize(self, image_path: Path, prompt_text: str | None = None) -> dict:
        api_key = os.environ.get("DASHSCOPE_API_KEY", "").strip()
        if not api_key:
            raise RuntimeError("尚未設定 DASHSCOPE_API_KEY，無法呼叫阿里百鍊 Qwen-OCR。")

        try:
            from openai import OpenAI
        except ImportError as exc:
            raise RuntimeError("請先安裝 openai 套件：python3 -m pip install openai") from exc

        b64, mime = _encode_image(image_path)
        client = OpenAI(
            api_key=api_key,
            base_url=os.environ.get("DASHSCOPE_BASE_URL", ALIYUN_COMPATIBLE_BASE_URL),
        )
        response = client.chat.completions.create(
            model=os.environ.get("ALIYUN_OCR_MODEL", self.model_name),
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                    {"type": "text", "text": prompt_text or OCR_TEXT_PROMPT},
                ],
            }],
            max_tokens=4096,
        )
        raw = (response.choices[0].message.content or "").strip()
        items = _parse_ocr_text_items(raw)
        result = _build_result(self.name, items)
        result["model"] = os.environ.get("ALIYUN_OCR_MODEL", self.model_name)
        result["raw_text"] = raw
        return result


class Qwen25Vl3bOcrProvider(AliyunQwenOcrProvider):
    name = "qwen25_vl_3b"
    model_name = QWEN25_VL_3B_MODEL


class Qwen25Vl7bOcrProvider(AliyunQwenOcrProvider):
    name = "qwen25_vl_7b"
    model_name = QWEN25_VL_7B_MODEL


class InternVl28bOcrProvider(AliyunQwenOcrProvider):
    name = "internvl2_8b"
    model_name = INTERNVL2_8B_MODEL


class ZhipuOcrProvider:
    name = "zhipu_ocr"

    def recognize(self, image_path: Path, prompt_text: str | None = None) -> dict:
        api_key = (os.environ.get("ZHIPUAI_API_KEY") or os.environ.get("ZHIPU_API_KEY") or "").strip()
        if not api_key:
            raise RuntimeError("尚未設定 ZHIPUAI_API_KEY，無法呼叫智譜 GLM-OCR。")

        b64, mime = _encode_image_for_zhipu(image_path)
        payload = {
            "model": os.environ.get("ZHIPU_OCR_MODEL", ZHIPU_OCR_MODEL),
            "file": f"data:{mime};base64,{b64}",
        }
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            os.environ.get("ZHIPU_OCR_URL", ZHIPU_LAYOUT_PARSING_URL),
            data=body,
            headers={
                "Authorization": api_key if api_key.lower().startswith("bearer ") else f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )

        try:
            with urllib.request.urlopen(request, timeout=90) as response:
                raw_text = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            error_body = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"智譜 GLM-OCR 呼叫失敗（HTTP {exc.code}）：{error_body[:500]}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"智譜 GLM-OCR 連線失敗：{exc}") from exc

        try:
            raw_payload = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"智譜 GLM-OCR 回傳不是 JSON：{raw_text[:500]}") from exc

        items = _flatten_zhipu_items(raw_payload)
        if not items:
            items = _parse_ocr_text_items(json.dumps(raw_payload, ensure_ascii=False))
        result = _build_result(self.name, items)
        result["model"] = payload["model"]
        result["raw_response"] = raw_payload
        return result


def _encode_image(image_path: Path) -> tuple[str, str]:
    suffix = image_path.suffix.lstrip(".").lower()
    mime = "image/jpeg" if suffix in ("jpg", "jpeg") else f"image/{suffix or 'png'}"
    with image_path.open("rb") as f:
        return base64.b64encode(f.read()).decode("utf-8"), mime


def _encode_image_for_zhipu(image_path: Path) -> tuple[str, str]:
    """Compress image before sending to GLM-OCR to avoid slow base64 uploads."""
    try:
        import io
        from PIL import Image
    except ImportError:
        return _encode_image(image_path)

    img = Image.open(image_path)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")

    max_side = int(os.environ.get("ZHIPU_OCR_MAX_SIDE", "1800"))
    width, height = img.size
    if max(width, height) > max_side:
        ratio = max_side / max(width, height)
        img = img.resize((max(1, int(width * ratio)), max(1, int(height * ratio))), Image.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=int(os.environ.get("ZHIPU_OCR_JPEG_QUALITY", "82")), optimize=True)
    return base64.b64encode(buf.getvalue()).decode("utf-8"), "image/jpeg"


def _parse_ocr_text_items(raw: str) -> list[dict]:
    text = raw.strip()
    if text.startswith("```"):
        text = text.removeprefix("```json").removeprefix("```").strip()
    if text.endswith("```"):
        text = text[:-3].strip()

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        parsed = None

    items: list[dict] = []
    if isinstance(parsed, list):
        for item in parsed:
            if isinstance(item, dict):
                value = item.get("text") or item.get("c") or item.get("company")
            else:
                value = item
            value = str(value or "").strip()
            if value:
                items.append({"text": value, "confidence": None, "box": None})
        return items

    for line in raw.replace("\r", "\n").split("\n"):
        value = line.strip().strip("-•·,，")
        if value:
            items.append({"text": value, "confidence": None, "box": None})
    return items


def _flatten_paddle_items(raw: Any) -> list[dict]:
    items: list[dict] = []

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            texts = node.get("rec_texts") or node.get("texts")
            scores = node.get("rec_scores") or node.get("scores") or []
            boxes = node.get("rec_boxes") or node.get("dt_polys") or node.get("boxes") or []
            if isinstance(texts, list):
                for index, text in enumerate(texts):
                    if not text:
                        continue
                    items.append({
                        "text": str(text).strip(),
                        "confidence": float(scores[index]) if index < len(scores) and scores[index] is not None else None,
                        "box": boxes[index].tolist() if index < len(boxes) and hasattr(boxes[index], "tolist") else (boxes[index] if index < len(boxes) else None),
                    })
                return
            for value in node.values():
                walk(value)
            return

        if not isinstance(node, list):
            return

        if len(node) >= 2 and isinstance(node[1], tuple) and len(node[1]) >= 2:
            text, confidence = node[1][0], node[1][1]
            if text:
                items.append({
                    "text": str(text).strip(),
                    "confidence": float(confidence) if confidence is not None else None,
                    "box": node[0],
                })
            return

        for child in node:
            walk(child)

    walk(raw)
    return [item for item in items if item["text"]]


def _flatten_zhipu_items(raw: Any) -> list[dict]:
    texts: list[str] = []

    def push(value: Any) -> None:
        value = str(value or "").strip()
        if value:
            texts.append(value)

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            for key in ("content", "text", "markdown", "md", "md_results", "html", "result"):
                value = node.get(key)
                if isinstance(value, str):
                    push(value)
            for value in node.values():
                if isinstance(value, (dict, list)):
                    walk(value)
            return
        if isinstance(node, list):
            for child in node:
                walk(child)
            return
        if isinstance(node, str):
            push(node)

    walk(raw)
    split_items: list[dict] = []
    seen: set[str] = set()
    for text in texts:
        for line in text.replace("\r", "\n").split("\n"):
            value = line.strip().strip("-•·,，| ")
            if not value or value in seen:
                continue
            seen.add(value)
            split_items.append({"text": value, "confidence": None, "box": None})
    return split_items


def get_ocr_provider(provider_name: str | None = None) -> OcrProvider:
    name = (provider_name or os.environ.get("OCR_PROVIDER") or DEFAULT_OCR_PROVIDER).strip().lower()
    if name not in SUPPORTED_PROVIDERS:
        raise RuntimeError(f"不支援的 OCR provider：{name}")
    if name == "disabled":
        return DisabledOcrProvider()
    if name == "paddle_local":
        return PaddleLocalOcrProvider()
    if name == "aliyun_ocr":
        return AliyunQwenOcrProvider()
    if name == "qwen25_vl_3b":
        return Qwen25Vl3bOcrProvider()
    if name == "qwen25_vl_7b":
        return Qwen25Vl7bOcrProvider()
    if name == "internvl2_8b":
        return InternVl28bOcrProvider()
    if name == "zhipu_ocr":
        return ZhipuOcrProvider()
    return PlaceholderCloudOcrProvider(name)


def resolve_ocr_prompt(profile: str | None) -> tuple[str, str]:
    name = (profile or "").strip().lower() or "default"
    prompt = OCR_PROMPT_PROFILES.get(name)
    if prompt:
        return name, prompt
    return "default", OCR_TEXT_PROMPT


def run_ocr_probe(image_path: Path, provider_name: str | None = None, prompt_profile: str | None = None) -> dict:
    provider = get_ocr_provider(provider_name)
    prompt_name, prompt_text = resolve_ocr_prompt(prompt_profile)
    result = provider.recognize(image_path, prompt_text=prompt_text)
    if isinstance(result, dict):
        result["prompt_profile"] = prompt_name
    return result
