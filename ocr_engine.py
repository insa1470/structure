from __future__ import annotations

from pathlib import Path
from typing import Any


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


def run_paddle_ocr(image_path: Path) -> dict:
    try:
        from paddleocr import PaddleOCR
    except Exception as exc:  # pragma: no cover - depends on optional heavy package
        raise RuntimeError(
            "PaddleOCR 尚未安裝。先執行 `python3 -m pip install -r requirements-paddle.txt` 後再測試。"
        ) from exc

    try:
        ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
    except TypeError:
        # PaddleOCR v3 has renamed several parameters.
        ocr = PaddleOCR(
            lang="ch",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=True,
        )

    raw = ocr.ocr(str(image_path), cls=True)
    items = _flatten_paddle_items(raw)
    company_candidates = [
        item for item in items
        if "公司" in item["text"] or "集团" in item["text"] or "集團" in item["text"]
    ]
    return {
        "engine": "paddleocr",
        "text_count": len(items),
        "company_candidate_count": len(company_candidates),
        "items": items,
        "company_candidates": company_candidates,
    }
