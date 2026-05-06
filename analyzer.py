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
from typing import Callable


QWEN_MODEL = "qwen2.5-vl-72b-instruct"

LEVEL_LABELS = {0: "頂層主體", 1: "一級子公司", 2: "二級子公司", 3: "三級子公司"}
ACTIVE_COMPANY_MARKERS = ("存续", "存續", "在业", "在業", "开业", "開業", "仍注册", "仍註冊")
INACTIVE_COMPANY_MARKERS = ("注销", "註銷", "吊销", "吊銷", "清算", "停业", "停業", "歇业", "歇業")

# ── 圖二分塊參數 ─────────────────────────────────────────────────────
# 手機截圖每張公司卡片約 90–110px（縮放後），1500px ≈ 15 家，辨識品質最佳
CHUNK_HEIGHT_PX    = 1500   # 每塊高度（縮放後 px）
CHUNK_OVERLAP_PX   = 180    # 相鄰塊重疊區，防止邊界公司被裁掉
CHUNK_THRESHOLD_PX = 3600   # 高於此值才切塊（max_tokens=8192 可覆蓋約 150 家，3600px ≈ 36 家/1000px）

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

PROMPT_CHART1_COMPANY_LIST = """這是一張企業股權結構圖。方框代表公司，連線代表股權關係。

這一次只做一件事：辨識圖中所有公司方框內的公司名稱。

規則：
- 只輸出圖片中實際看得到的公司名稱，禁止推測、禁止補全。
- 暫時不要判斷層級、父子關係或持股比例。
- 不要輸出「存續」「法代」「資本」「成立」「100%」等非公司名稱文字。
- 公司名稱必須完整抄寫；看不清楚就保留你能看清楚的原文並標記 uncertain=true。
- 不要重複輸出同一家公司。

只輸出 JSON array，不要說明文字或 markdown：
[{"c":"公司全名","uncertain":false}]
"""

PROMPT_CHART2_STAGE1 = """這是一張企業查詢 App 的子公司列表長截圖。畫面由多張公司卡片自上而下排列，每張卡片代表一家公司。

你的任務是第一階段：只抽公司清單，不抽細節。

請嚴格依照畫面從上到下，一家公司一家公司處理，不能跳列，不能合併相鄰兩家公司。

規則：
- 忽略頂部導覽列、搜尋列、篩選列、底部 logo、浮水印、廣告、App 提示文字、icon、頁面按鈕。
- 只處理真正的公司列表卡片。
- 每張公司卡片只抽這 1 個欄位：
  - c: 公司全名
- 公司名稱必須完整抄寫，不可縮寫。
- 看不清楚就填 null，不要編造。
- 忽略右側「二級子公司 / 三級子公司 / 四級子公司」標籤。
- 忽略「存續 / 在業 / 註銷」狀態標籤。
- 忽略「實控人總持股」「企業族群中心」「高新技術企業」等其他資訊。
- 不要漏掉中間任何一張公司卡片。
- 不要輸出重複公司。

只輸出 JSON array，不要說明文字或 markdown：
[{"c":"公司全名"}]
"""

PROMPT_CHART2_STAGE2 = """這是一張企業查詢 App 的子公司列表長截圖。畫面由多張公司卡片自上而下排列，每張卡片代表一家公司。

你的任務是第二階段：只抽取每張公司卡片中的 4 個欄位。

規則：
- 忽略頂部導覽列、搜尋列、篩選列、底部 logo、浮水印、廣告、App 提示文字、icon、頁面按鈕。
- 只處理真正的公司列表卡片。
- 每張公司卡片只抽以下 4 個欄位：
  - c: 公司全名
  - lr: 法定代表人 / 法人代表
  - rc: 註冊資本
  - ed: 成立時間
- 忽略右側「二級子公司 / 三級子公司 / 四級子公司」標籤。
- 忽略「存續 / 在業 / 註銷」狀態標籤。
- 忽略「實控人總持股」「企業族群中心」「高新技術企業」等其他資訊。
- 看不清楚就填 null，不要編造。
- 不要把欄位標題「法定代表人」「註冊資本」「成立日期」當成值。
- 不要把上一家公司的資訊接到下一家公司。
- 同一家公司只輸出一次。
- 即使只有公司名清楚、其他欄位不清楚，也要保留該公司，其他欄位填 null。

只輸出 JSON array，不要說明文字或 markdown：
[{"c":"公司全名","lr":"法人代表或null","rc":"資本額如5000萬元或null","ed":"成立時間如2014-12-03或null"}]
"""


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


def _log_info(message: str) -> None:
    print(message, flush=True)


def _log_warn(message: str) -> None:
    print(message, flush=True)

def _call_qwen_vl(image_path: Path, prompt: str) -> list[dict]:
    """呼叫 Qwen-VL，回傳解析後的 list。JSON 解析失敗時自動重試最多 MAX_RETRIES 次。"""
    last_err: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return _call_qwen_vl_once(image_path, prompt)
        except RuntimeError as e:
            last_err = e
            _log_warn(f"[Qwen] attempt {attempt}/{MAX_RETRIES} failed: {e}")
            if attempt < MAX_RETRIES:
                import time
                time.sleep(2)  # 短暫等待後重試
    raise last_err  # type: ignore


def _call_qwen_vl_once(image_path: Path, prompt: str) -> list[dict]:
    """單次呼叫 Qwen-VL。"""
    import re as _re
    import ast

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
        max_tokens=8192,
    )

    finish_reason = response.choices[0].finish_reason
    raw = response.choices[0].message.content.strip()

    _log_info(f"[Qwen] finish_reason={finish_reason} raw_len={len(raw)}")
    if os.environ.get("DEBUG_QWEN_RAW", "").strip().lower() in {"1", "true", "yes"}:
        preview = raw[:1200] + ("..." if len(raw) > 1200 else "")
        _log_info(f"[Qwen] raw_preview={preview}")

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
        # 1.5 字串值在物件或陣列結尾前少了右引號
        #     例如 "ed":"2014-12-03} → "ed":"2014-12-03"}
        text = _re.sub(r'(:")([^"\n]*?)(?=\s*[}\]])', r'\1\2"', text)
        # 2. 數字/null/true/false 後接雜散引號，例如 "l":2" → "l":2
        text = _re.sub(r'(:\s*)(\d+(?:\.\d+)?|null|true|false)"(\s*[,}\]])', r'\1\2\3', text)
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

    # 策略 5：逐物件搶救。整包 JSON 壞掉時，只要 {...} 還在，就逐筆修復並解析。
    fragments = _re.findall(r"\{[^{}]+\}", target, flags=_re.S)
    salvaged: list[dict] = []
    if fragments:
        for fragment in fragments:
            fragment = _repair_json(fragment)
            fragment = _repair_unquoted(fragment)
            for parser in (json.loads, ast.literal_eval):
                try:
                    item = parser(fragment)
                    if isinstance(item, dict):
                        salvaged.append(item)
                        break
                except Exception:
                    continue
        if salvaged:
            return salvaged

    raise RuntimeError(f"無法解析模型回傳的 JSON。原始回應（前500字）：{raw[:500]}")


# ── 繁→簡 對照表（公司名稱常用字，無需外部套件）────────────────────
_TRAD_TO_SIMP: dict[str, str] = {
    # 通用高頻字
    "協": "协", "創": "创", "數": "数", "據": "据", "興": "兴",
    "廣": "广", "電": "电", "網": "网", "業": "业", "開": "开",
    "發": "发", "實": "实", "際": "际", "來": "来", "設": "设",
    "計": "计", "資": "资", "產": "产", "義": "义", "務": "务",
    "術": "术", "專": "专", "題": "题", "問": "问", "應": "应",
    "關": "关", "聯": "联", "國": "国", "運": "运", "動": "动",
    "輸": "输", "儲": "储", "億": "亿", "萬": "万", "銀": "银",
    "貿": "贸", "農": "农", "強": "强", "醫": "医", "藥": "药",
    "機": "机", "圖": "图", "導": "导", "總": "总", "長": "长",
    "學": "学", "華": "华", "麗": "丽", "豐": "丰", "達": "达",
    "遠": "远", "進": "进", "門": "门", "區": "区", "時": "时",
    "間": "间", "層": "层", "樓": "楼", "陽": "阳", "龍": "龙",
    "鳳": "凤", "鋼": "钢", "鐵": "铁", "礦": "矿", "廠": "厂",
    "縣": "县", "鎮": "镇", "鄉": "乡", "歷": "历", "橋": "桥",
    "財": "财", "貨": "货", "幣": "币", "鋁": "铝", "銅": "铜",
    "視": "视", "頻": "频", "聲": "声", "記": "记", "習": "习",
    "勞": "劳", "雲": "云", "線": "线", "點": "点", "場": "场",
    "變": "变", "種": "种", "類": "类", "體": "体", "輕": "轻",
    "項": "项", "組": "组", "統": "统", "維": "维", "織": "织",
    "經": "经", "濟": "济", "稱": "称", "稅": "税", "積": "积",
    "優": "优", "驗": "验", "觀": "观", "購": "购", "銷": "销",
    "費": "费", "質": "质", "標": "标", "續": "续", "聽": "听",
    "讀": "读", "語": "语", "課": "课", "識": "识", "試": "试",
    "護": "护", "報": "报", "讓": "让", "認": "认", "調": "调",
    "談": "谈", "請": "请", "說": "说", "論": "论", "評": "评",
    "辦": "办", "處": "处", "節": "节", "環": "环", "樹": "树",
    "藝": "艺", "傳": "传", "訊": "讯", "號": "号", "複": "复",
    "製": "制", "幫": "帮", "從": "从", "給": "给", "針": "针",
    "鍵": "键", "鏈": "链", "鎖": "锁", "簡": "简", "歡": "欢",
    "樂": "乐", "愛": "爱", "親": "亲", "寬": "宽", "嚴": "严",
    "氣": "气", "熱": "热", "匯": "汇", "帶": "带", "錢": "钱",
    "縱": "纵", "聰": "聪", "穩": "稳", "償": "偿", "債": "债",
    "貸": "贷", "臺": "台", "灣": "湾", "歐": "欧", "亞": "亚",
    "盧": "卢", "羅": "罗", "馬": "马", "廈": "厦", "鏡": "镜",
    "緒": "绪", "範": "范", "補": "补", "備": "备", "縮": "缩",
    "繁": "繁", "鑰": "钥", "樣": "样", "紹": "绍", "絡": "络",
    "結": "结", "約": "约", "純": "纯", "細": "细", "紅": "红",
    "綠": "绿", "藍": "蓝", "黃": "黄", "黑": "黑", "銳": "锐",
    "錄": "录", "鏡": "镜", "際": "际", "環": "环", "競": "竞",
    "爭": "争", "興": "兴", "縣": "县", "鎮": "镇",
}
_TRAD_TRANS = str.maketrans(_TRAD_TO_SIMP)

# 公司類型後綴（比對前可選擇性去除以提高核心名稱相似度）
_CORP_SUFFIXES = (
    "股份有限公司", "有限責任公司", "有限合夥", "合夥企業",
    "有限公司", "股份公司", "集團", "控股",
)

_NON_COMPANY_TEXTS = (
    "存续", "存續", "在业", "在業", "开业", "開業", "注销", "註銷", "吊销", "吊銷",
    "法代", "法人", "法人代表", "法定代表人", "资本", "資本", "注册资本", "註冊資本",
    "成立", "成立日期", "成立时间", "成立時間", "实控人", "實控人", "总持股", "總持股",
    "二级子公司", "二級子公司", "三级子公司", "三級子公司", "四级子公司", "四級子公司",
    "查看", "详情", "詳情", "企业", "企業", "公司列表",
)

_COMPANY_HINTS = (
    "公司", "集团", "集團", "企业", "企業", "厂", "廠", "中心", "合伙", "合夥",
    "事务所", "事務所", "合作社", "银行", "銀行", "院", "所",
)


def _normalize_for_match(name: str) -> str:
    """比對前正規化：繁→簡 + 去空白。不去後綴（避免誤配）。"""
    return name.translate(_TRAD_TRANS).strip()


def _fuzzy_match(name_a: str, name_b: str) -> float:
    """模糊比對，處理簡繁差異與縮寫 vs 全名兩種情況。"""
    a = _normalize_for_match(name_a)
    b = _normalize_for_match(name_b)
    base = SequenceMatcher(None, a, b).ratio()
    # 縮寫包含於全名（如「协创数据」包含於「协创数据技术股份有限公司」）
    if len(a) >= 2 and len(b) >= 2:
        short, long = (a, b) if len(a) <= len(b) else (b, a)
        if short in long:
            return max(base, 0.88)
    return base


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


def _looks_like_company_name(name: str) -> bool:
    value = _normalize_text(name)
    if not value or len(value) < 4:
        return False
    if value in _CORP_SUFFIXES:
        return False
    if value.replace(".", "", 1).replace("%", "").isdigit():
        return False
    if any(token == value for token in _NON_COMPANY_TEXTS):
        return False
    if any(token in value for token in _NON_COMPANY_TEXTS) and not any(hint in value for hint in _COMPANY_HINTS):
        return False
    return any(hint in value for hint in _COMPANY_HINTS) or len(value) >= 8


def _evaluate_chart1_quality(nodes: list[dict]) -> dict:
    count = len(nodes)
    roots = [node for node in nodes if not node.get("parent")]
    max_level = max((_parse_level(node.get("level"), 0) for node in nodes), default=0)
    uncertain_count = sum(1 for node in nodes if node.get("uncertain"))
    suspicious_names = [node.get("company", "") for node in nodes if not _looks_like_company_name(node.get("company", ""))]
    parent_count = sum(1 for node in nodes if node.get("parent"))
    score = 100
    notes: list[str] = []

    if count == 0:
        return {
            "score": 0,
            "level": "low",
            "needs_rescue": True,
            "notes": ["圖一沒有辨識到公司"],
            "company_count": 0,
            "root_count": 0,
            "max_level": 0,
            "rescue_used": False,
        }

    if count < 2:
        score -= 45
        notes.append("公司數過少")
    elif count < 4:
        score -= 15
        notes.append("公司數偏少")

    if count >= 5 and max_level == 0:
        score -= 35
        notes.append("所有公司都在同一層")
    elif count >= 8 and max_level <= 1:
        score -= 18
        notes.append("層級偏扁")

    if len(roots) > 1:
        root_ratio = len(roots) / max(count, 1)
        if root_ratio > 0.45:
            score -= 30
            notes.append("多數公司沒有上層")
        elif len(roots) > 2:
            score -= 15
            notes.append("頂層公司偏多")

    if count >= 4 and parent_count == 0:
        score -= 28
        notes.append("缺少父子關係")

    suspicious_ratio = len(suspicious_names) / max(count, 1)
    if suspicious_ratio > 0.35:
        score -= 25
        notes.append("疑似非公司名稱偏多")
    elif suspicious_names:
        score -= 8
        notes.append("部分名稱需確認")

    uncertain_ratio = uncertain_count / max(count, 1)
    if uncertain_ratio > 0.4:
        score -= 12
        notes.append("不確定節點偏多")

    score = max(0, min(100, score))
    level = "high" if score >= 75 else "mid" if score >= 55 else "low"
    return {
        "score": score,
        "level": level,
        "needs_rescue": score < 60,
        "notes": notes or ["圖一結構初步合理"],
        "company_count": count,
        "root_count": len(roots),
        "max_level": max_level,
        "suspicious_name_count": len(suspicious_names),
        "rescue_used": False,
    }


def _parse_chart1_nodes(raw: list[dict]) -> list[dict]:
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


def _analyze_chart1_company_names(image_path: Path) -> list[str]:
    raw = _call_qwen_vl(image_path, PROMPT_CHART1_COMPANY_LIST)
    names: list[str] = []
    seen: set[str] = set()
    for item in raw:
        name = _normalize_text(item.get("company") or item.get("c") or item.get("text") or "")
        if not name or name in seen or not _looks_like_company_name(name):
            continue
        seen.add(name)
        names.append(name)
    return names


def _build_chart1_relation_prompt(company_names: list[str]) -> str:
    company_json = json.dumps(company_names, ensure_ascii=False)
    return f"""這是一張企業股權結構圖。你已先辨識出以下公司名稱清單，請以圖片中的連線為準，重建直接上下層關係與持股比例。

公司名稱清單：
{company_json}

規則：
- 優先使用清單中的公司名稱，不要任意新增不存在的公司。
- 每家公司只填直接上層，不要跳層。
- 如果看不清楚父層，p 填 null，並把 uncertain 設為 true。
- 如果持股比例看不清楚，r 填 null。
- 層級 l 從 0 開始，頂層為 0。
- 只輸出 JSON array，不要說明文字或 markdown。

格式：
[{{"c":"公司全名","p":"直接上層公司全名或null","r":"51.2%或null","l":0,"uncertain":false}}]
"""


def _append_missing_chart1_names(nodes: list[dict], names: list[str]) -> list[dict]:
    existing = {node["company"] for node in nodes}
    root = next((node["company"] for node in nodes if not node.get("parent")), "")
    merged = [dict(node) for node in nodes]
    for name in names:
        if name in existing:
            continue
        merged.append({
            "company": name,
            "parent": root or None,
            "shareholding_ratio": None,
            "level": 1 if root else 0,
            "uncertain": True,
        })
        existing.add(name)
    return _sanitize_chart1_nodes(merged)


def _find_best_match(name: str, candidates: list[dict], threshold: float = 0.85) -> tuple[dict | None, float]:
    best, best_score = None, 0.0
    for c in candidates:
        score = _fuzzy_match(name, c.get("company", ""))
        if score > best_score:
            best, best_score = c, score
    return (best, best_score) if best_score >= threshold else (None, best_score)


def _dedupe_companies(rows: list[dict]) -> list[dict]:
    deduped: list[dict] = []
    seen: set[str] = set()
    for row in rows:
        company = _normalize_text(row.get("company"))
        if not company or company in seen:
            continue
        seen.add(company)
        row["company"] = company
        deduped.append(row)
    return deduped


def analyze_chart1(image_path: Path) -> list[dict]:
    """解析圖一（股權結構圖），回傳節點清單。"""
    nodes, _quality = analyze_chart1_with_quality(image_path)
    return nodes


def analyze_chart1_with_quality(image_path: Path) -> tuple[list[dict], dict]:
    """先跑原本完整辨識；若結果品質偏低，再啟動公司名清單 + 層級重建補救。"""
    raw = _call_qwen_vl(image_path, PROMPT_CHART1)
    primary_nodes = _parse_chart1_nodes(raw)
    primary_quality = _evaluate_chart1_quality(primary_nodes)
    _log_info(
        f"[Chart1] primary quality={primary_quality['score']} "
        f"count={primary_quality['company_count']} max_level={primary_quality['max_level']}"
    )

    if not primary_quality["needs_rescue"]:
        return primary_nodes, primary_quality

    try:
        company_names = _analyze_chart1_company_names(image_path)
        _log_info(f"[Chart1] rescue company list count={len(company_names)}")
    except Exception as exc:
        primary_quality["notes"] = list(primary_quality.get("notes", [])) + [f"補救公司清單失敗：{exc}"]
        return primary_nodes, primary_quality

    rescued_nodes: list[dict] = []
    rescued_quality: dict | None = None
    if company_names:
        try:
            rescue_raw = _call_qwen_vl(image_path, _build_chart1_relation_prompt(company_names))
            rescued_nodes = _parse_chart1_nodes(rescue_raw)
            rescued_quality = _evaluate_chart1_quality(rescued_nodes)
            _log_info(
                f"[Chart1] rescue quality={rescued_quality['score']} "
                f"count={rescued_quality['company_count']} max_level={rescued_quality['max_level']}"
            )
        except Exception as exc:
            primary_quality["notes"] = list(primary_quality.get("notes", [])) + [f"補救層級重建失敗：{exc}"]

    merged_nodes = _append_missing_chart1_names(primary_nodes, company_names) if company_names else primary_nodes
    merged_quality = _evaluate_chart1_quality(merged_nodes)

    selected_nodes, selected_quality, selected_mode = primary_nodes, primary_quality, "primary"
    if rescued_nodes and rescued_quality:
        if (
            rescued_quality.get("score", 0) >= primary_quality.get("score", 0)
            or rescued_quality.get("company_count", 0) > primary_quality.get("company_count", 0)
        ):
            selected_nodes, selected_quality, selected_mode = rescued_nodes, rescued_quality, "rescue"

    # 公司名稱是後續人工修正的底稿；若補救清單抓到更多公司，即使層級仍不完美，也保留為待確認節點。
    if selected_mode == "primary" and len(merged_nodes) > len(primary_nodes):
        selected_nodes, selected_quality, selected_mode = merged_nodes, merged_quality, "primary_plus_missing_names"

    selected_quality = dict(selected_quality)
    selected_quality["rescue_used"] = selected_mode != "primary"
    selected_quality["rescue_mode"] = selected_mode
    selected_quality["rescued_company_count"] = len(company_names)
    if selected_mode != "primary":
        selected_quality["notes"] = list(selected_quality.get("notes", [])) + ["已啟用圖一補救辨識"]
    return selected_nodes, selected_quality


def _split_image_into_chunks(image_path: Path) -> list[Path]:
    """將長截圖垂直切割成多個重疊小塊，回傳暫存檔路徑清單。
    若圖片不夠長或 PIL 未安裝，回傳原路徑的單元素清單。
    """
    import io
    import tempfile

    try:
        from PIL import Image
    except ImportError:
        return [image_path]

    img = Image.open(image_path)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    w, h = img.size

    # 先縮到 900px 寬（與 _encode_image 直式圖邏輯一致）
    if w > 900:
        ratio = 900 / w
        img = img.resize((900, int(h * ratio)), Image.LANCZOS)
        w, h = img.size

    if h <= CHUNK_THRESHOLD_PX:
        return [image_path]

    tmp_dir = Path(tempfile.mkdtemp(prefix="codex_c2_"))
    chunks: list[Path] = []
    step = CHUNK_HEIGHT_PX - CHUNK_OVERLAP_PX
    y = 0
    idx = 0
    while True:
        y_end = min(y + CHUNK_HEIGHT_PX, h)
        chunk_img = img.crop((0, y, w, y_end))
        chunk_path = tmp_dir / f"chunk_{idx:02d}.jpg"
        buf = io.BytesIO()
        chunk_img.save(buf, format="JPEG", quality=85)
        chunk_path.write_bytes(buf.getvalue())
        chunks.append(chunk_path)
        if y_end >= h:
            break
        y += step
        idx += 1

    _log_info(
        f"[Chart2] 圖高 {h}px，切成 {len(chunks)} 塊"
        f"（每塊 {CHUNK_HEIGHT_PX}px，重疊 {CHUNK_OVERLAP_PX}px）"
    )
    return chunks


def _dedup_merge_chart2(rows: list[dict], threshold: float = 0.85) -> list[dict]:
    """合併多塊辨識結果：相似度 >= threshold 視為同一家公司。
    重複時保留較完整的欄位（非空優先）。
    """
    merged: list[dict] = []
    for row in rows:
        company = row.get("company", "").strip()
        if not company:
            continue
        existing, _ = _find_best_match(company, merged, threshold)
        if existing:
            # 用本塊的非空值填充舊紀錄的空欄位
            for field in ("legal_representative", "registered_capital", "established_date"):
                if not existing.get(field) and row.get(field):
                    existing[field] = row[field]
        else:
            merged.append(dict(row))
    return merged


def _parse_chart2_detail_rows(raw_stage2: list[dict]) -> list[dict]:
    return _dedupe_companies([
        {
            "company": _normalize_text(item.get("company") or item.get("c") or ""),
            "legal_representative": _normalize_text(item.get("legal_representative") or item.get("lr") or None) or None,
            "registered_capital": _normalize_text(item.get("registered_capital") or item.get("rc") or None) or None,
            "established_date": _normalize_text(item.get("established_date") or item.get("ed") or None) or None,
        }
        for item in raw_stage2
    ])


def _chart2_output_rows(rows: list[dict], uncertain: bool) -> list[dict]:
    return [
        {
            "company": row["company"],
            "legal_representative": row.get("legal_representative"),
            "registered_capital": row.get("registered_capital"),
            "established_date": row.get("established_date"),
            "company_status": None,
            "subsidiary_level_label": None,
            "actual_controller_share": None,
            "uncertain": uncertain,
        }
        for row in rows
        if row.get("company")
    ]


def _analyze_chart2_single(image_path: Path) -> list[dict]:
    """對單一圖片（或單塊）解析圖二；正常情況只呼叫一次模型。"""
    try:
        raw_stage2 = _call_qwen_vl(image_path, PROMPT_CHART2_STAGE2)
        detail_rows = _parse_chart2_detail_rows(raw_stage2)
        if detail_rows:
            return _chart2_output_rows(detail_rows, uncertain=False)
    except RuntimeError:
        detail_rows = []

    raw_stage1 = _call_qwen_vl(image_path, PROMPT_CHART2_STAGE1)
    stage1_rows = _dedupe_companies([
        {"company": _normalize_text(item.get("company") or item.get("c") or "")}
        for item in raw_stage1
    ])
    return _chart2_output_rows(stage1_rows, uncertain=True)


def analyze_chart2(
    image_path: Path,
    progress_callback: Callable[[dict], None] | None = None,
    should_continue: Callable[[], bool] | None = None,
) -> list[dict]:
    """解析圖二（集團概覽），回傳公司屬性清單。超長截圖自動切塊處理。"""
    import shutil
    chunk_paths = _split_image_into_chunks(image_path)
    is_chunked = not (len(chunk_paths) == 1 and chunk_paths[0] == image_path)
    tmp_dir = chunk_paths[0].parent if is_chunked else None

    try:
        all_rows: list[dict] = []
        failed_chunks: list[dict] = []
        if progress_callback:
            progress_callback({
                "status": "running",
                "current_chunk": 0,
                "total_chunks": len(chunk_paths),
                "rows_so_far": 0,
                "failed_chunks": [],
            })
        for i, chunk_path in enumerate(chunk_paths):
            if should_continue and not should_continue():
                if progress_callback:
                    progress_callback({
                        "status": "cancelled",
                        "current_chunk": i,
                        "total_chunks": len(chunk_paths),
                        "rows_so_far": len(all_rows),
                        "failed_chunks": failed_chunks,
                    })
                raise InterruptedError("task_cancelled")
            try:
                rows = _analyze_chart2_single(chunk_path)
                _log_info(f"[Chart2] 塊 {i + 1}/{len(chunk_paths)} 辨識到 {len(rows)} 家")
                all_rows.extend(rows)
                if progress_callback:
                    progress_callback({
                        "status": "running",
                        "current_chunk": i + 1,
                        "total_chunks": len(chunk_paths),
                        "last_chunk_rows": len(rows),
                        "rows_so_far": len(all_rows),
                        "failed_chunks": failed_chunks,
                    })
            except RuntimeError as exc:
                _log_warn(f"[Chart2] 塊 {i + 1}/{len(chunk_paths)} 失敗，跳過：{exc}")
                failed_chunks.append({"chunk": i + 1, "message": str(exc)[:300]})
                if progress_callback:
                    progress_callback({
                        "status": "running",
                        "current_chunk": i + 1,
                        "total_chunks": len(chunk_paths),
                        "last_chunk_rows": 0,
                        "rows_so_far": len(all_rows),
                        "failed_chunks": failed_chunks,
                    })

        if not all_rows:
            raise RuntimeError("圖二所有分塊均辨識失敗，請確認圖片品質後重試")

        if is_chunked:
            deduped = _dedup_merge_chart2(all_rows)
            _log_info(f"[Chart2] 合併後共 {len(deduped)} 家（原始 {len(all_rows)} 筆）")
            if progress_callback:
                progress_callback({
                    "status": "partial_done" if failed_chunks else "done",
                    "current_chunk": len(chunk_paths),
                    "total_chunks": len(chunk_paths),
                    "rows_so_far": len(all_rows),
                    "deduped_count": len(deduped),
                    "failed_chunks": failed_chunks,
                })
            return deduped

        if progress_callback:
            progress_callback({
                "status": "partial_done" if failed_chunks else "done",
                "current_chunk": len(chunk_paths),
                "total_chunks": len(chunk_paths),
                "rows_so_far": len(all_rows),
                "deduped_count": len(all_rows),
                "failed_chunks": failed_chunks,
            })
        return all_rows
    finally:
        if tmp_dir and tmp_dir.exists():
            shutil.rmtree(tmp_dir, ignore_errors=True)


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
    chart1_nodes, chart1_quality = analyze_chart1_with_quality(chart1_path)

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
        "summary": {**_make_summary(master_rows, [], []), "chart1_quality": chart1_quality},
        "graph": _make_graph(master_rows),
    }


def enrich_with_chart2_precomputed(existing_master_rows: list[dict], chart2_attrs: list[dict]) -> dict:
    """用已辨識完成的圖二資料補充主表（跳過 OCR，直接 merge）。

    只更新補充欄位：legal_representative, registered_capital, established_date,
    actual_controller_share, company_status, subsidiary_level_label。
    不動：node_id, canonical_name, chart1_level, chart1_parent, chart1_parent_name。
    """

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


def enrich_with_chart2(existing_master_rows: list[dict], chart2_path: Path) -> dict:
    """第二階段：用圖二補充現有主表（含 OCR + merge）。"""
    chart2_attrs = analyze_chart2(chart2_path)
    return enrich_with_chart2_precomputed(existing_master_rows, chart2_attrs)


def run_analysis(chart1_path: Path, chart2_path: Path) -> dict:
    """完整分析流程（舊介面相容用，server.py 直接呼叫兩段式）。"""
    stage1 = run_chart1_stage(chart1_path)
    return enrich_with_chart2(stage1["master_rows"], chart2_path)
