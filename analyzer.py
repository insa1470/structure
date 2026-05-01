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


QWEN_MODEL = "qwen2.5-vl-7b-instruct"

LEVEL_LABELS = {0: "頂層主體", 1: "一級子公司", 2: "二級子公司", 3: "三級子公司"}

PROMPT_CHART1 = """這是一張企查查股權結構圖。
請識別所有公司節點和持股關係，輸出以下 JSON 格式（只輸出 JSON，不要其他說明文字）：

[
  {
    "company": "子公司名稱",
    "parent": "上層公司名稱（頂層公司則填 null）",
    "shareholding_ratio": "51.2%（若無則填 null）",
    "level": 0,
    "uncertain": false
  }
]

說明：
- level 從 0 開始，頂層主體為 0，一級子公司為 1，依此類推
- 若某個連線關係不確定（線條模糊、方向不明），請將 uncertain 設為 true
- 只輸出 JSON array，不要 markdown 代碼塊"""

PROMPT_CHART2 = """這是一張企查查集團公司概覽圖或表格。
請提取每家公司的基本資訊，輸出以下 JSON 格式（只輸出 JSON，不要其他說明文字）：

[
  {
    "company": "公司名稱",
    "legal_representative": "法人代表（若無填 null）",
    "registered_capital": "1000萬元（若無填 null）",
    "established_date": "2015-03-01（若無填 null）",
    "company_status": "存續（若無填 null）",
    "subsidiary_level_label": "一級子公司（若無填 null）",
    "actual_controller_share": "51.2%（若無填 null）",
    "uncertain": false
  }
]

只輸出 JSON array，不要 markdown 代碼塊"""


def _encode_image(image_path: Path) -> tuple[str, str]:
    """回傳 (base64_data, mime_type)。"""
    suffix = image_path.suffix.lstrip(".").lower()
    mime = "image/jpeg" if suffix in ("jpg", "jpeg") else f"image/{suffix}"
    with image_path.open("rb") as f:
        return base64.b64encode(f.read()).decode("utf-8"), mime


def _call_qwen_vl(image_path: Path, prompt: str) -> list[dict]:
    """呼叫 Qwen-VL API，回傳解析後的 JSON list。"""
    try:
        import dashscope
        from dashscope import MultiModalConversation
    except ImportError:
        raise RuntimeError("請先安裝 dashscope：pip install dashscope")

    api_key = os.environ.get("DASHSCOPE_API_KEY", "")
    if not api_key:
        raise RuntimeError("請設定環境變數 DASHSCOPE_API_KEY")

    dashscope.api_key = api_key
    b64, mime = _encode_image(image_path)

    messages = [
        {
            "role": "user",
            "content": [
                {"image": f"data:{mime};base64,{b64}"},
                {"text": prompt},
            ],
        }
    ]

    response = MultiModalConversation.call(model=QWEN_MODEL, messages=messages)

    if response.status_code != 200:
        raise RuntimeError(f"Qwen-VL API 錯誤：{response.code} {response.message}")

    raw = response.output.choices[0].message.content[0]["text"].strip()
    raw = raw.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    return json.loads(raw)


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
    return _call_qwen_vl(image_path, PROMPT_CHART1)


def analyze_chart2(image_path: Path) -> list[dict]:
    """解析圖二（集團概覽），回傳公司屬性清單。"""
    return _call_qwen_vl(image_path, PROMPT_CHART2)


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
