"""
analyzer.py — 用 Qwen-VL 解析企查查股權圖，輸出整併後的主表資料。

需要設定環境變數：
    DASHSCOPE_API_KEY=<你的阿里雲百煉 API Key>
"""

from __future__ import annotations

import base64
import json
import os
import uuid
from difflib import SequenceMatcher
from pathlib import Path


QWEN_MODEL = "qwen2.5-vl-72b-instruct"

LEVEL_LABELS = {0: "頂層主體", 1: "一級子公司", 2: "二級子公司", 3: "三級子公司"}
ACTIVE_COMPANY_MARKERS = ("存续", "存續", "在业", "在業", "开业", "開業", "仍注册", "仍註冊")
INACTIVE_COMPANY_MARKERS = ("注销", "註銷", "吊销", "吊銷", "清算", "停业", "停業", "歇业", "歇業")

PROMPT_CHART1 = """這是一張企業股權結構圖。方框代表公司，連線代表股權關係。圖可能左右展開，層級以連線方向為準，不以版面位置上下為準。

請由上而下逐層辨識：

第一步：找出頂層公司（Level 0）。通常是圖中最上方或中央的大框，沒有任何上層連線，只有一家。"p" = null。

第二步：找出所有直接連線到 Level 0 公司的方框（Level 1）。每家 "p" = Level 0 公司名。

第三步：找出所有直接連線到某個 Level 1 公司的方框（Level 2）。每家 "p" = 各自對應的那個 Level 1 公司名，不能全部填同一個。

第四步：依此類推，辨識 Level 3、Level 4……直到沒有更下層為止。

注意：
- 每家公司的 "p" 只填直接相連的上層公司，不跳層
- 同一條連線下可能有多家子公司，全部列出
- 務必辨識圖中【全部】方框，包括細長小框和豎排文字，不得遺漏
- 只輸出圖片中實際可見的文字，禁止猜測

只輸出JSON array，不要說明文字或markdown：
[{"c":"公司全名","p":"直接上層公司全名或null","r":"51.2%或null","l":0}]

輸出前請自我檢查：每個字串值是否都有開始和結束的雙引號？每個數字值（l欄位）後面是否沒有多餘的引號？"""

PROMPT_CHART2 = """讀取集團公司列表，提取每家公司資訊。

規則：
- 只輸出圖片中實際可見的文字，禁止猜測
- 公司名稱完整抄寫，不得縮寫
- 看不清楚填null

只輸出JSON array，不要說明文字或markdown：
[{"c":"公司全名","lr":"法人代表或null","rc":"註冊資本如1000萬元或null","ed":"成立日期如2015-03-01或null","cs":"狀態如存續或null","sl":"如一級子公司或null","ac":"如51.2%或null"}]

輸出前請自我檢查：每個日期字串（ed欄位）是否有完整的開始和結束雙引號？格式應為 "ed":"2023-01-01" 而不是 "ed":"2023-01-01,"""


def _encode_image(image_path: Path) -> tuple[str, str]:
    """
    智慧縮放圖片後回傳 (base64, mime_type)：
    - 直式圖（高 > 2×寬）：以寬度為基準，限寬 900px，高度不壓縮
      （適用手機截圖公司列表，避免寬度縮到幾百px讓文字無法辨識）
    - 橫式/方形圖：最長邊限 1568px
    """
    import io
    try:
        from PIL import Image
        img = Image.open(image_path)
        if img.mode in ("RGBA", "LA", "P"):
            img = img.convert("RGB")
        w, h = img.width, img.height
        if h > w * 2:
            # 直式圖：限寬 900px，保持高度（讓模型讀完整列表）
            max_w = 900
            if w > max_w:
                ratio = max_w / w
                img = img.resize((max_w, int(h * ratio)), Image.LANCZOS)
        else:
            # 橫式/方形圖：限最長邊 1568px
            max_px = 1568
            if max(w, h) > max_px:
                ratio = max_px / max(w, h)
                img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode(), "image/jpeg"
    except ImportError:
        suffix = image_path.suffix.lstrip(".").lower()
        mime = "image/jpeg" if suffix in ("jpg", "jpeg") else f"image/{suffix}"
        with image_path.open("rb") as f:
            return base64.b64encode(f.read()).decode("utf-8"), mime


MAX_RETRIES = 1  # JSON 解析失敗時最多重試次數（失敗就讓用戶重傳，不要讓他等太久）

def _call_qwen_vl(image_path: Path, prompt: str) -> list[dict]:
    """呼叫 Qwen-VL，回傳解析後的 list。JSON 解析失敗時自動重試最多 MAX_RETRIES 次。"""
    last_err: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return _call_qwen_vl_once(image_path, prompt)
        except RuntimeError as e:
            last_err = e
            import sys
            print(f"[Qwen] attempt {attempt}/{MAX_RETRIES} failed: {e}", file=sys.stderr)
            if attempt < MAX_RETRIES:
                import time
                time.sleep(2)  # 短暫等待後重試
    raise last_err  # type: ignore


def _call_qwen_vl_once(image_path: Path, prompt: str) -> list[dict]:
    """單次呼叫 Qwen-VL。"""
    import re as _re
    import ast
    import sys

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
        messages=[{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                {"type": "text", "text": prompt},
            ],
        }],
        max_tokens=4096,
    )

    finish_reason = response.choices[0].finish_reason
    raw = response.choices[0].message.content.strip()

    print(f"[Qwen] finish_reason={finish_reason} raw_len={len(raw)}", file=sys.stderr)
    print(f"[Qwen] raw={raw}", file=sys.stderr)

    if finish_reason == "length":
        raise RuntimeError(f"模型輸出被截斷（公司數量過多），raw_len={len(raw)}，請裁切圖片後重試")

    # 去除 markdown 包裝
    raw = _re.sub(r"^```(?:json)?\s*", "", raw)
    raw = _re.sub(r"\s*```$", "", raw)
    raw = raw.strip()

    # ── 前置清理：修復常見格式錯誤 ────────────────────────────
    def _repair_json(text: str) -> str:
        # 1. 字串值缺少結尾引號，下一個 key 緊跟其後
        #    例如 "ed":"2003-11-16,"cs" → "ed":"2003-11-16","cs"
        #    模式：:"非引號內容,"  → :"非引號內容","
        text = _re.sub(r':"([^",\n{}\[\]]+),(")', r':"\1",\2', text)
        # 2. 數字/null/true/false 後接雜散引號，例如 "l":2" → "l":2
        text = _re.sub(r'(\b(?:\d+(?:\.\d+)?|null|true|false))"(\s*[,}\]])', r'\1\2', text)
        # 3. 字串值結尾多一個引號，例如 "abc"" → "abc"
        text = _re.sub(r'""(\s*[,}\]])', r'"\1', text)
        # 4. 尾隨逗號（陣列/物件最後一個元素後的逗號）
        text = _re.sub(r',(\s*[}\]])', r'\1', text)
        # 5. 中文全形逗號換成半形
        text = text.replace("，", ",")
        # 6. 截斷的 JSON：結尾若沒有 ] 就補上（只在有 [ 的情況）
        stripped = text.rstrip()
        if stripped.startswith("[") and not stripped.endswith("]"):
            last_close = stripped.rfind("}")
            if last_close != -1:
                text = stripped[:last_close + 1] + "]"
        return text

    raw = _repair_json(raw)

    # 策略 1：標準 JSON
    try:
        result = json.loads(raw)
        if isinstance(result, list):
            return result
    except json.JSONDecodeError:
        pass

    # 策略 2：ast.literal_eval（容許單引號）
    try:
        result = ast.literal_eval(raw)
        if isinstance(result, list):
            return result
    except (ValueError, SyntaxError):
        pass

    # 策略 3：括號配對提取第一個完整 array
    candidate = None
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

    # 策略 4：補上遺漏的引號（模型有時對字串值省略引號）再重試
    import re as _re2

    def _repair_unquoted(text: str) -> str:
        fixed = []
        for line in text.split("\n"):
            m = _re2.match(r'^(\s*"[^"]+"\s*:\s*)(.+?)(\s*,?\s*)$', line)
            if m:
                prefix, value = m.group(1), m.group(2).rstrip("，,").strip()
                # 已是合法 JSON 值則不動
                if (value.startswith('"') or value.startswith("[") or value.startswith("{")
                        or value in ("null", "true", "false")):
                    fixed.append(line)
                    continue
                # 純數字則不動
                try:
                    float(value); fixed.append(line); continue
                except ValueError:
                    pass
                # 補引號，保留行尾逗號
                trailing = "," if (line.rstrip().endswith(",") or "，" in line) else ""
                fixed.append(f'{prefix}"{value}"{trailing}')
            else:
                fixed.append(line)
        return "\n".join(fixed)

    target = candidate if candidate is not None else raw
    repaired = _repair_unquoted(target)
    if repaired != target:
        for parser in (json.loads, ast.literal_eval):
            try:
                result = parser(repaired)
                if isinstance(result, list):
                    return result
            except Exception:
                pass

    raise RuntimeError(f"無法解析模型回傳的 JSON。原始回應（前500字）：{raw[:500]}")


def _fuzzy_match(name_a: str, name_b: str) -> float:
    return SequenceMatcher(None, name_a, name_b).ratio()


def _normalize_text(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def _is_active_company_status(status: str | None) -> bool:
    normalized = _normalize_text(status)
    if not normalized:
        return False
    if any(marker in normalized for marker in INACTIVE_COMPANY_MARKERS):
        return False
    return any(marker in normalized for marker in ACTIVE_COMPANY_MARKERS)


def _parse_level(value: object, default: int = 0) -> int:
    try:
        return max(int(value), 0)
    except (TypeError, ValueError):
        return default


def _sanitize_chart1_nodes(chart1_nodes: list[dict]) -> list[dict]:
    cleaned: list[dict] = []
    seen_companies: set[str] = set()

    for node in chart1_nodes:
        company = _normalize_text(node.get("company"))
        if not company:
            continue
        if company in seen_companies:
            continue
        seen_companies.add(company)

        parent = _normalize_text(node.get("parent")) or None
        if parent == company:
            parent = None

        cleaned.append({
            "company": company,
            "parent": parent,
            "shareholding_ratio": _normalize_text(node.get("shareholding_ratio")) or None,
            "level": _parse_level(node.get("level"), 0),
            "uncertain": bool(node.get("uncertain", False)),
        })

    companies = {node["company"] for node in cleaned}
    parent_levels = {node["company"]: node["level"] for node in cleaned}
    roots = [node for node in cleaned if not node.get("parent")]

    for node in cleaned:
        parent = node.get("parent")
        if parent and parent not in companies:
            node["parent"] = None
            node["uncertain"] = True
            continue
        if parent:
            expected_level = parent_levels.get(parent, 0) + 1
            if node["level"] <= parent_levels.get(parent, -1):
                node["level"] = expected_level
                node["uncertain"] = True
        else:
            if node["level"] != 0:
                node["level"] = 0
                node["uncertain"] = True

    if len(roots) > 1:
        for node in roots[1:]:
            node["uncertain"] = True

    return cleaned


def _find_best_match(name: str, candidates: list[dict], threshold: float = 0.85) -> tuple[dict | None, float]:
    best, best_score = None, 0.0
    for c in candidates:
        score = _fuzzy_match(name, c.get("company", ""))
        if score > best_score:
            best, best_score = c, score
    return (best, best_score) if best_score >= threshold else (None, best_score)


def analyze_chart1(image_path: Path) -> list[dict]:
    """解析圖一（股權結構圖），回傳節點清單。"""
    raw = _call_qwen_vl(image_path, PROMPT_CHART1)
    result = []
    for item in raw:
        result.append({
            "company": item.get("company") or item.get("c") or "",
            "parent": item.get("parent") or item.get("p") or None,
            "shareholding_ratio": item.get("shareholding_ratio") or item.get("r") or None,
            "level": item.get("level") if item.get("level") is not None else (item.get("l") or 0),
            "uncertain": item.get("uncertain", False),
        })
    return _sanitize_chart1_nodes(result)


def analyze_chart2(image_path: Path) -> list[dict]:
    """解析圖二（集團概覽），回傳公司屬性清單。"""
    raw = _call_qwen_vl(image_path, PROMPT_CHART2)
    result = []
    for item in raw:
        row = {
            "company": _normalize_text(item.get("company") or item.get("c") or ""),
            "legal_representative": _normalize_text(item.get("legal_representative") or item.get("lr") or None) or None,
            "registered_capital": _normalize_text(item.get("registered_capital") or item.get("rc") or None) or None,
            "established_date": _normalize_text(item.get("established_date") or item.get("ed") or None) or None,
            "company_status": _normalize_text(item.get("company_status") or item.get("cs") or None) or None,
            "subsidiary_level_label": _normalize_text(item.get("subsidiary_level_label") or item.get("sl") or None) or None,
            "actual_controller_share": _normalize_text(item.get("actual_controller_share") or item.get("ac") or None) or None,
            "uncertain": item.get("uncertain", False),
        }
        if row["company"] and _is_active_company_status(row["company_status"]):
            result.append(row)
    return result


def merge_charts(
    chart1_nodes: list[dict],
    chart2_attrs: list[dict],
) -> tuple[list[dict], list[dict], list[dict]]:
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
                "review_note": f"模糊比對，相似度 {match_score:.2f}",
            })
            review_rows.append({
                "issue_type": "review_match",
                "chart1_name": company,
                "chart2_name": chart2_match.get("company", ""),
                "candidate_node_id": node_id,
                "match_score": f"{match_score:.4f}",
                "recommended_action": "confirm_match_or_reject",
                "review_status": "pending",
                "review_note": f"模糊比對分數 {match_score:.2f}",
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


def _make_graph(master_rows: list[dict]) -> dict:
    return {
        "nodes": [{"id": r["node_id"], "label": r["canonical_name"], "level": r["chart1_level"]} for r in master_rows],
        "edges": [
            {"source": r["chart1_parent"], "target": r["node_id"], "ratio": r["actual_controller_share"]}
            for r in master_rows if r.get("chart1_parent")
        ],
    }


def _make_summary(master_rows: list[dict], review_rows: list[dict], candidate_rows: list[dict]) -> dict:
    return {
        "master_count": len(master_rows),
        "enriched_count": sum(1 for r in master_rows if r.get("node_status") == "enriched"),
        "review_count": len(review_rows),
        "chart1_only_count": sum(1 for r in master_rows if r.get("node_status") == "chart1_only"),
        "candidate_count": len(candidate_rows),
    }


def run_chart1_stage(chart1_path: Path) -> dict:
    """第一階段：只分析圖一，回傳骨架主表（全部為 chart1_only 狀態）。"""
    chart1_nodes = analyze_chart1(chart1_path)

    master_rows: list[dict] = []
    node_ids: dict[str, str] = {}

    def get_node_id(name: str) -> str:
        if name not in node_ids:
            node_ids[name] = f"N{len(node_ids) + 1:03d}"
        return node_ids[name]

    for node in chart1_nodes:
        company = node.get("company", "").strip()
        parent_name = node.get("parent") or ""
        level = node.get("level", 0)
        ratio = node.get("shareholding_ratio") or ""
        node_id = get_node_id(company)
        parent_id = get_node_id(parent_name) if parent_name else ""
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
            "review_flag": "",
            "review_note": "",
        })

    return {
        "master_rows": master_rows,
        "review_rows": [],
        "candidate_rows": [],
        "summary": _make_summary(master_rows, [], []),
        "graph": _make_graph(master_rows),
    }


def enrich_with_chart2(existing_master_rows: list[dict], chart2_path: Path) -> dict:
    """第二階段：用圖二補充現有主表，保留用戶已調整的結構（層級、父層、名稱）。

    只更新補充欄位：legal_representative, registered_capital, established_date,
    actual_controller_share, company_status, subsidiary_level_label。
    不動：node_id, canonical_name, chart1_level, chart1_parent, chart1_parent_name。
    """
    chart2_attrs = analyze_chart2(chart2_path)

    master_rows = [dict(r) for r in existing_master_rows]  # 深拷貝，不改原始資料
    review_rows: list[dict] = []
    candidate_rows: list[dict] = []
    matched_chart2: set[str] = set()

    for row in master_rows:
        company = row.get("canonical_name") or row.get("chart1_name", "")
        chart2_match, match_score = _find_best_match(company, chart2_attrs)

        if chart2_match and match_score >= 0.85:
            matched_chart2.add(chart2_match.get("company", ""))
            # 只更新補充欄位
            row["matched_chart2_name"] = chart2_match.get("company", "")
            row["legal_representative"] = chart2_match.get("legal_representative") or row.get("legal_representative", "")
            row["established_date"] = chart2_match.get("established_date") or row.get("established_date", "")
            row["registered_capital"] = chart2_match.get("registered_capital") or row.get("registered_capital", "")
            row["actual_controller_share"] = chart2_match.get("actual_controller_share") or row.get("actual_controller_share", "")
            row["subsidiary_level_label"] = chart2_match.get("subsidiary_level_label") or row.get("subsidiary_level_label", "")
            row["company_status"] = chart2_match.get("company_status") or row.get("company_status", "")
            row["match_status"] = "matched"
            row["node_status"] = "enriched"
            row["review_flag"] = ""
            row["review_note"] = ""
            matched_chart2.add(chart2_match.get("company", ""))
        elif chart2_match and match_score >= 0.6:
            row["matched_chart2_name"] = chart2_match.get("company", "")
            row["match_status"] = "review_match"
            row["node_status"] = "review_match"
            row["review_flag"] = "yes"
            row["review_note"] = f"模糊比對，相似度 {match_score:.2f}"
            review_rows.append({
                "issue_type": "review_match",
                "chart1_name": company,
                "chart2_name": chart2_match.get("company", ""),
                "candidate_node_id": row["node_id"],
                "match_score": f"{match_score:.4f}",
                "recommended_action": "confirm_match_or_reject",
                "review_status": "pending",
                "review_note": f"模糊比對分數 {match_score:.2f}",
            })
        else:
            row["node_status"] = "chart1_only"
            row["review_flag"] = "yes"
            review_rows.append({
                "issue_type": "chart1_only",
                "chart1_name": company,
                "chart2_name": "",
                "candidate_node_id": row["node_id"],
                "match_score": "",
                "recommended_action": "check_if_chart2_missing_or_inactive",
                "review_status": "pending",
                "review_note": "圖二無安全的對應項目",
            })

    for attr in chart2_attrs:
        company = attr.get("company", "").strip()
        if company and company not in matched_chart2:
            candidate_rows.append({
                "node_id": f"C{uuid.uuid4().hex[:6].upper()}",
                "company": company,
                "chart2_name": company,
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

    return {
        "master_rows": master_rows,
        "review_rows": review_rows,
        "candidate_rows": candidate_rows,
        "summary": _make_summary(master_rows, review_rows, candidate_rows),
        "graph": _make_graph(master_rows),
    }


def run_analysis(chart1_path: Path, chart2_path: Path) -> dict:
    """完整分析流程（舊介面相容用，server.py 直接呼叫兩段式）。"""
    stage1 = run_chart1_stage(chart1_path)
    return enrich_with_chart2(stage1["master_rows"], chart2_path)
