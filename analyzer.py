"""
analyzer.py — 用 Qwen-VL 解析企查查股權圖，輸出整併後的主表資料。

需要設定環境變數：
    DASHSCOPE_API_KEY=<你的阿里雲百煉 API Key>

安裝依賴：
    pip install dashscope
"""

from __future__ import annotations

import base64
import json
import os
import uuid
from difflib import SequenceMatcher
from pathlib import Path


QWEN_MODEL = "qwen2.5-vl-32b-instruct"

LEVEL_LABELS = {0: "頂層主體", 1: "一級子公司", 2: "二級子公司", 3: "三級子公司"}

PROMPT_CHART1 = """讀取股權結構圖，識別所有方框內的公司名稱和連線上的持股比例。

規則：
- 只輸出圖片中實際可見的文字，禁止猜測
- level：頂層主體=0，一級子公司=1，依此類推
- parent：填上層公司全名，頂層填null
- ratio：填持股比例如"51.2%"，無則填null

只輸出JSON array，不要說明文字或markdown：
[{"c":"公司全名","p":"上層公司全名或null","r":"51.2%或null","l":0}]"""

PROMPT_CHART2 = """讀取集團公司列表，提取每家公司資訊。

規則：
- 只輸出圖片中實際可見的文字，禁止猜測
- 公司名稱完整抄寫，不得縮寫
- 看不清楚填null

只輸出JSON array，不要說明文字或markdown：
[{"c":"公司全名","lr":"法人代表或null","rc":"註冊資本如1000萬元或null","ed":"成立日期如2015-03-01或null","cs":"狀態如存續或null","sl":"如一級子公司或null","ac":"如51.2%或null"}]"""


def _encode_image(image_path: Path) -> tuple[str, str]:
    """壓縮圖片至合適大小，回傳 (base64_data, mime_type)。"""
    import io
    try:
        from PIL import Image
        img = Image.open(image_path)
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGB")
        # Qwen-VL 推薦最大 1120px
        max_px = 1120
        if max(img.width, img.height) > max_px:
            ratio = max_px / max(img.width, img.height)
            img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=88)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode(), "image/jpeg"
    except ImportError:
        # Pillow 未安裝時直接讀原檔
        suffix = image_path.suffix.lstrip(".").lower()
        mime = "image/jpeg" if suffix in ("jpg", "jpeg") else f"image/{suffix}"
        with image_path.open("rb") as f:
            return base64.b64encode(f.read()).decode("utf-8"), mime


def _call_qwen_vl(image_path: Path, prompt: str) -> list[dict]:
    """呼叫 Qwen-VL API（OpenAI 相容格式），回傳解析後的 JSON list。"""
    import re as _re

    api_key = os.environ.get("DASHSCOPE_API_KEY", "")
    if not api_key:
        raise RuntimeError("請設定環境變數 DASHSCOPE_API_KEY")

    try:
        from openai import OpenAI
    except ImportError:
        raise RuntimeError("請安裝 openai：pip install openai")

    b64, mime = _encode_image(image_path)

    client = OpenAI(
        api_key=api_key,
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )

    response = client.chat.completions.create(
        model=QWEN_MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                    {"type": "text", "text": prompt},
                ],
            }
        ],
        max_tokens=16384,
    )

    finish_reason = response.choices[0].finish_reason
    raw = response.choices[0].message.content.strip()

    # 完整記錄到 stderr 方便 Railway log 查看
    import sys
    print(f"[Qwen] finish_reason={finish_reason} raw_len={len(raw)}", file=sys.stderr)
    print(f"[Qwen] raw={raw}", file=sys.stderr)

    if finish_reason == "length":
        raise RuntimeError(f"模型輸出被截斷（token 上限），raw_len={len(raw)}，請裁切圖片後重試")

    # 移除 markdown 包裝
    raw = _re.sub(r"^```(?:json)?\s*", "", raw)
    raw = _re.sub(r"\s*```$", "", raw)
    raw = raw.strip()

    import ast

    # 策略 1：直接 JSON 解析
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # 策略 2：ast.literal_eval（接受單/雙引號混用的 Python dict 格式）
    try:
        result = ast.literal_eval(raw)
        if isinstance(result, list):
            return result
    except (ValueError, SyntaxError):
        pass

    # 策略 3：括號配對找第一個完整 array，再用 json / ast 解析
    start = raw.find("[")
    if start != -1:
        depth, in_str, quote_ch, esc = 0, False, '"', False
        for i, ch in enumerate(raw[start:], start):
            if esc:
                esc = False; continue
            if ch == "\\" and in_str:
                esc = True; continue
            if not in_str and ch in ('"', "'"):
                in_str, quote_ch = True, ch; continue
            if in_str and ch == quote_ch:
                in_str = False; continue
            if in_str:
                continue
            if ch == "[":
                depth += 1
            elif ch == "]":
                depth -= 1
                if depth == 0:
                    candidate = raw[start: i + 1]
                    for parser in (json.loads, ast.literal_eval):
                        try:
                            result = parser(candidate)
                            if isinstance(result, list):
                                return result
                        except Exception:
                            pass
                    break

    raise RuntimeError(f"無法解析模型回傳的 JSON。模型原始回應（前300字）：{raw[:300]}")


def _fuzzy_match(name_a: str, name_b: str) -> float:
    return SequenceMatcher(None, name_a, name_b).ratio()


def _find_best_match(name: str, candidates: list[dict], threshold: float = 0.85) -> tuple[dict | None, float]:
    best, best_score = None, 0.0
    for c in candidates:
        score = _fuzzy_match(name, c.get("company", ""))
        if score > best_score:
            best, best_score = c, score
    return (best, best_score) if best_score >= threshold else (None, best_score)


def analyze_chart1(image_path: Path) -> list[dict]:
    """解析圖一（股權結構圖），回傳節點與關係清單。"""
    raw = _call_qwen_vl(image_path, PROMPT_CHART1)
    # 正規化短 key → 長 key（相容舊格式與新精簡格式）
    result = []
    for item in raw:
        result.append({
            "company": item.get("company") or item.get("c") or "",
            "parent": item.get("parent") or item.get("p") or None,
            "shareholding_ratio": item.get("shareholding_ratio") or item.get("r") or None,
            "level": item.get("level") if item.get("level") is not None else (item.get("l") or 0),
            "uncertain": item.get("uncertain", False),
        })
    return result


def analyze_chart2(image_path: Path) -> list[dict]:
    """解析圖二（集團概覽），回傳公司屬性清單。"""
    raw = _call_qwen_vl(image_path, PROMPT_CHART2)
    # 正規化短 key → 長 key（相容舊格式與新精簡格式）
    result = []
    for item in raw:
        result.append({
            "company": item.get("company") or item.get("c") or "",
            "legal_representative": item.get("legal_representative") or item.get("lr") or None,
            "registered_capital": item.get("registered_capital") or item.get("rc") or None,
            "established_date": item.get("established_date") or item.get("ed") or None,
            "company_status": item.get("company_status") or item.get("cs") or None,
            "subsidiary_level_label": item.get("subsidiary_level_label") or item.get("sl") or None,
            "actual_controller_share": item.get("actual_controller_share") or item.get("ac") or None,
            "uncertain": item.get("uncertain", False),
        })
    return result


def merge_charts(
    chart1_nodes: list[dict],
    chart2_attrs: list[dict],
) -> tuple[list[dict], list[dict], list[dict]]:
    """
    整合圖一節點與圖二屬性，回傳：
        master_rows    — 主表（對應 master_nodes_enriched.csv 格式）
        review_rows    — 待確認清單（對應 reconciliation_report.csv 格式）
        candidate_rows — 圖二獨有候選（對應 chart2_only_candidates.csv 格式）
    """
    master_rows: list[dict] = []
    review_rows: list[dict] = []
    candidate_rows: list[dict] = []

    node_ids: dict[str, str] = {}

    def get_node_id(name: str) -> str:
        if name not in node_ids:
            node_ids[name] = f"N{len(node_ids) + 1:03d}"
        return node_ids[name]

    matched_chart2: set[str] = set()

    for node in chart1_nodes:
        company = node.get("company", "").strip()
        parent_name = node.get("parent") or ""
        level = node.get("level", 0)
        ratio = node.get("shareholding_ratio") or ""
        uncertain = node.get("uncertain", False)

        node_id = get_node_id(company)
        parent_id = get_node_id(parent_name) if parent_name else ""

        chart2_match, match_score = _find_best_match(company, chart2_attrs)

        if chart2_match and not uncertain:
            matched_chart2.add(chart2_match.get("company", ""))
            master_rows.append({
                "node_id": node_id,
                "chart1_name": company,
                "canonical_name": company,
                "chart1_level": level,
                "chart1_parent": parent_id,
                "chart1_parent_name": parent_name,
                "matched_chart2_name": chart2_match.get("company", ""),
                "legal_representative": chart2_match.get("legal_representative") or "",
                "established_date": chart2_match.get("established_date") or "",
                "registered_capital": chart2_match.get("registered_capital") or "",
                "actual_controller_share": chart2_match.get("actual_controller_share") or ratio,
                "subsidiary_level_label": chart2_match.get("subsidiary_level_label") or LEVEL_LABELS.get(level, ""),
                "company_status": chart2_match.get("company_status") or "",
                "match_status": "matched",
                "node_status": "enriched",
                "review_flag": "",
                "review_note": "",
            })

        elif chart2_match and uncertain:
            matched_chart2.add(chart2_match.get("company", ""))
            master_rows.append({
                "node_id": node_id,
                "chart1_name": company,
                "canonical_name": company,
                "chart1_level": level,
                "chart1_parent": parent_id,
                "chart1_parent_name": parent_name,
                "matched_chart2_name": chart2_match.get("company", ""),
                "legal_representative": chart2_match.get("legal_representative") or "",
                "established_date": chart2_match.get("established_date") or "",
                "registered_capital": chart2_match.get("registered_capital") or "",
                "actual_controller_share": chart2_match.get("actual_controller_share") or ratio,
                "subsidiary_level_label": chart2_match.get("subsidiary_level_label") or LEVEL_LABELS.get(level, ""),
                "company_status": chart2_match.get("company_status") or "",
                "match_status": "fuzzy",
                "node_status": "review_match",
                "review_flag": "yes",
                "review_note": f"Qwen-VL 不確定此節點關係，相似度 {match_score:.2f}",
            })
            review_rows.append({
                "issue_type": "review_match",
                "chart1_name": company,
                "chart2_name": chart2_match.get("company", ""),
                "candidate_node_id": node_id,
                "match_score": f"{match_score:.4f}",
                "recommended_action": "confirm_match_or_reject",
                "review_status": "pending",
                "review_note": f"Qwen-VL 標記 uncertain；模糊比對分數 {match_score:.2f}",
            })

        else:
            master_rows.append({
                "node_id": node_id,
                "chart1_name": company,
                "canonical_name": company,
                "chart1_level": level,
                "chart1_parent": parent_id,
                "chart1_parent_name": parent_name,
                "matched_chart2_name": "",
                "legal_representative": "",
                "established_date": "",
                "registered_capital": "",
                "actual_controller_share": ratio,
                "subsidiary_level_label": LEVEL_LABELS.get(level, ""),
                "company_status": "",
                "match_status": "chart1_only",
                "node_status": "chart1_only",
                "review_flag": "yes",
                "review_note": "圖二無對應項目",
            })
            review_rows.append({
                "issue_type": "chart1_only",
                "chart1_name": company,
                "chart2_name": "",
                "candidate_node_id": node_id,
                "match_score": "",
                "recommended_action": "check_if_chart2_missing_or_inactive",
                "review_status": "pending",
                "review_note": "圖二無安全的對應項目",
            })

    for attr in chart2_attrs:
        company = attr.get("company", "").strip()
        if company and company not in matched_chart2:
            node_id = f"C{uuid.uuid4().hex[:6].upper()}"
            candidate_rows.append({
                "node_id": node_id,
                "company": company,
                "legal_representative": attr.get("legal_representative") or "",
                "established_date": attr.get("established_date") or "",
                "registered_capital": attr.get("registered_capital") or "",
                "actual_controller_share": attr.get("actual_controller_share") or "",
                "subsidiary_level_label": attr.get("subsidiary_level_label") or "",
                "company_status": attr.get("company_status") or "",
                "suggested_parent": "",
                "decision": "pending",
                "note": "圖二獨有，圖一無對應節點",
            })

    return master_rows, review_rows, candidate_rows


def run_analysis(chart1_path: Path, chart2_path: Path) -> dict:
    """
    完整分析流程入口。
    回傳與 server.py build_task_payload 相容的 payload dict。
    """
    chart1_nodes = analyze_chart1(chart1_path)
    chart2_attrs = analyze_chart2(chart2_path)
    master_rows, review_rows, candidate_rows = merge_charts(chart1_nodes, chart2_attrs)

    return {
        "master_rows": master_rows,
        "review_rows": review_rows,
        "candidate_rows": candidate_rows,
        "summary": {
            "master_count": len(master_rows),
            "enriched_count": sum(1 for r in master_rows if r.get("node_status") == "enriched"),
            "review_count": len(review_rows),
            "chart1_only_count": sum(1 for r in master_rows if r.get("node_status") == "chart1_only"),
            "candidate_count": len(candidate_rows),
        },
        "graph": {
            "nodes": [
                {"id": r["node_id"], "label": r["canonical_name"], "level": r["chart1_level"]}
                for r in master_rows
            ],
            "edges": [
                {
                    "source": r["chart1_parent"],
                    "target": r["node_id"],
                    "ratio": r["actual_controller_share"],
                }
                for r in master_rows
                if r.get("chart1_parent")
            ],
            "stage2": {
                "status": "reserved",
                "ready_after_review": True,
                "target_output": "equity_structure_chart",
                "note": "第二階段：審核完成後，基於 master_rows + review_decisions + candidate_decisions 生成最終股權架構圖。",
            },
        },
    }
