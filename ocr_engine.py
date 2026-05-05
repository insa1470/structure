from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any, Protocol


DEFAULT_OCR_PROVIDER = "disabled"
ALIYUN_QWEN_OCR_MODEL = "qwen-vl-ocr"
ALIYUN_COMPATIBLE_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
SUPPORTED_PROVIDERS = {"disabled", "paddle_local", "aliyun_ocr", "baidu_ocr", "tencent_ocr"}

OCR_TEXT_PROMPT = """請對這張圖片做 OCR 文字辨識。

規則：
- 只輸出圖片中實際看得到的文字，不要推測、不要補全。
- 盡量依照畫面從上到下、從左到右排序。
- 公司名稱請完整保留。
- 如果文字不確定，也保留原文，不要改寫。

只輸出 JSON array，不要說明文字或 markdown：
[{"text":"辨識到的文字"}]
"""


class OcrProvider(Protocol):
    name: str

    def recognize(self, image_path: Path) -> dict:
        ...


def _company_candidates(items: list[dict]) -> list[dict]:
    return [
        item for item in items
        if "公司" in item["text"] or "集团" in item["text"] or "集團" in item["text"]
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

    def recognize(self, image_path: Path) -> dict:
        raise RuntimeError("OCR provider 尚未啟用。請設定 OCR_PROVIDER=paddle_local 或接入雲端 OCR provider。")


class PlaceholderCloudOcrProvider:
    def __init__(self, name: str):
        self.name = name

    def recognize(self, image_path: Path) -> dict:
        raise RuntimeError(f"{self.name} provider 尚未接入。此版本只預留介面，尚未呼叫雲端 OCR API。")


class PaddleLocalOcrProvider:
    name = "paddle_local"

    def recognize(self, image_path: Path) -> dict:
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

    def recognize(self, image_path: Path) -> dict:
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
            model=os.environ.get("ALIYUN_OCR_MODEL", ALIYUN_QWEN_OCR_MODEL),
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                    {"type": "text", "text": OCR_TEXT_PROMPT},
                ],
            }],
            max_tokens=4096,
        )
        raw = (response.choices[0].message.content or "").strip()
        items = _parse_ocr_text_items(raw)
        result = _build_result(self.name, items)
        result["model"] = os.environ.get("ALIYUN_OCR_MODEL", ALIYUN_QWEN_OCR_MODEL)
        result["raw_text"] = raw
        return result


def _encode_image(image_path: Path) -> tuple[str, str]:
    suffix = image_path.suffix.lstrip(".").lower()
    mime = "image/jpeg" if suffix in ("jpg", "jpeg") else f"image/{suffix or 'png'}"
    with image_path.open("rb") as f:
        return base64.b64encode(f.read()).decode("utf-8"), mime


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
    return PlaceholderCloudOcrProvider(name)


def run_ocr_probe(image_path: Path, provider_name: str | None = None) -> dict:
    provider = get_ocr_provider(provider_name)
    return provider.recognize(image_path)
