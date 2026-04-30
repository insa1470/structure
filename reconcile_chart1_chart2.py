from __future__ import annotations

import csv
import re
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path


BASE = Path(__file__).resolve().parent


def read_csv(path: Path):
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, rows, fieldnames):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


PAREN_MAP = str.maketrans({
    "（": "(",
    "）": ")",
    "【": "[",
    "】": "]",
    "　": " ",
})


def normalize_name(name: str) -> str:
    text = (name or "").translate(PAREN_MAP)
    text = re.sub(r"\s+", "", text)
    return text.lower()


def level_from_label(label: str) -> int | None:
    mapping = {
        "本级": 0,
        "一级子公司": 1,
        "二级子公司": 2,
        "三级子公司": 3,
        "四级子公司": 4,
        "五级子公司": 5,
    }
    return mapping.get((label or "").strip())


def is_english_name(name: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9 .,&()'/-]+", name or ""))


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_name(a), normalize_name(b)).ratio()


def main():
    nodes = read_csv(BASE / "qcc_nodes.csv")
    edges = read_csv(BASE / "qcc_edges.csv")
    chart2 = read_csv(BASE / "chart2_extract_active.csv")

    parent_by_target = {edge["target_id"]: edge["source_id"] for edge in edges}
    node_by_id = {row["id"]: row for row in nodes}

    chart1_rows = []
    raw_name_map = defaultdict(list)
    norm_name_map = defaultdict(list)
    for node in nodes:
        chart1 = {
            "node_id": node["id"],
            "chart1_name": node["name"],
            "chart1_level": node["level"],
            "chart1_parent": parent_by_target.get(node["id"], ""),
            "chart1_confidence": node["confidence"],
            "chart1_notes": node["notes"],
        }
        chart1_rows.append(chart1)
        raw_name_map[node["name"]].append(chart1)
        norm_name_map[normalize_name(node["name"])].append(chart1)

    chart2_matches = []
    matched_by_node = defaultdict(list)
    review_by_node = defaultdict(list)

    for row in chart2:
        level2 = level_from_label(row["subsidiary_level_label"])
        exact_candidates = raw_name_map.get(row["chart2_name"], [])
        norm_candidates = norm_name_map.get(normalize_name(row["chart2_name"]), [])
        status = ""
        match_method = ""
        candidate_node = None
        score = ""
        note = row.get("notes", "")

        if len(exact_candidates) == 1:
            candidate_node = exact_candidates[0]
            status = "matched"
            match_method = "exact"
            score = "1.0000"
        elif len(exact_candidates) > 1:
            status = "review_match"
            match_method = "exact_multi"
            score = "1.0000"
            note = f"{note}; exact match duplicated".strip("; ")
        elif len(norm_candidates) == 1:
            candidate_node = norm_candidates[0]
            status = "matched"
            match_method = "normalized"
            score = "0.9950"
        elif len(norm_candidates) > 1:
            status = "review_match"
            match_method = "normalized_multi"
            score = "0.9950"
            note = f"{note}; normalized match duplicated".strip("; ")
        else:
            scored = sorted(
                (
                    {
                        "node_id": item["node_id"],
                        "chart1_name": item["chart1_name"],
                        "score": similarity(row["chart2_name"], item["chart1_name"]),
                    }
                    for item in chart1_rows
                ),
                key=lambda x: x["score"],
                reverse=True,
            )
            best = scored[0]
            second = scored[1] if len(scored) > 1 else {"score": 0}
            if best["score"] >= 0.72:
                status = "review_match"
                match_method = "fuzzy"
                score = f"{best['score']:.4f}"
                candidate_node = next(item for item in chart1_rows if item["node_id"] == best["node_id"])
                note = f"{note}; fuzzy candidate {best['chart1_name']}; gap {best['score'] - second['score']:.4f}".strip("; ")
            else:
                status = "chart2_only"
                match_method = "unmatched"
                score = f"{best['score']:.4f}"
                note = f"{note}; no safe candidate".strip("; ")

        if candidate_node and status == "matched":
            chart1_level = int(candidate_node["chart1_level"])
            if is_english_name(row["chart2_name"]) and row["chart2_name"] != candidate_node["chart1_name"]:
                status = "review_match"
                note = f"{note}; overseas english name requires manual review".strip("; ")
            elif level2 is not None and chart1_level != level2:
                status = "review_match"
                note = f"{note}; level conflict chart1={chart1_level} chart2={level2}".strip("; ")

        result = {
            **row,
            "candidate_node_id": candidate_node["node_id"] if candidate_node else "",
            "candidate_chart1_name": candidate_node["chart1_name"] if candidate_node else "",
            "chart2_level_num": level2 if level2 is not None else "",
            "match_method": match_method,
            "match_score": score,
            "match_status": status,
            "review_note": note.strip("; "),
        }
        chart2_matches.append(result)
        if candidate_node:
            if status == "matched":
                matched_by_node[candidate_node["node_id"]].append(result)
            else:
                review_by_node[candidate_node["node_id"]].append(result)

    master_rows = []
    reconciliation_rows = []
    chart2_only_rows = []

    for node in chart1_rows:
        matched = matched_by_node.get(node["node_id"], [])
        review = review_by_node.get(node["node_id"], [])
        row = {
            "node_id": node["node_id"],
            "chart1_name": node["chart1_name"],
            "canonical_name": node["chart1_name"],
            "chart1_level": node["chart1_level"],
            "chart1_parent": node["chart1_parent"],
            "chart1_parent_name": node_by_id.get(node["chart1_parent"], {}).get("name", ""),
            "matched_chart2_name": "",
            "legal_representative": "",
            "established_date": "",
            "registered_capital": "",
            "actual_controller_share": "",
            "subsidiary_level_label": "",
            "company_status": "",
            "match_status": "",
            "node_status": "",
            "review_flag": "",
            "review_note": node["chart1_notes"],
        }
        if len(matched) == 1:
            m = matched[0]
            row.update({
                "matched_chart2_name": m["chart2_name"],
                "legal_representative": m["legal_representative"],
                "established_date": m["established_date"],
                "registered_capital": m["registered_capital"],
                "actual_controller_share": m["actual_controller_share"],
                "subsidiary_level_label": m["subsidiary_level_label"],
                "company_status": m["company_status"],
                "match_status": m["match_status"],
                "node_status": "enriched",
                "review_flag": "",
                "review_note": m["review_note"] or node["chart1_notes"],
            })
        elif review:
            top = sorted(review, key=lambda x: float(x["match_score"] or 0), reverse=True)[0]
            row.update({
                "matched_chart2_name": top["chart2_name"],
                "match_status": top["match_status"],
                "node_status": "review_match",
                "review_flag": "yes",
                "review_note": top["review_note"] or node["chart1_notes"],
            })
            reconciliation_rows.append({
                "issue_type": "review_match",
                "chart1_name": node["chart1_name"],
                "chart2_name": top["chart2_name"],
                "candidate_node_id": node["node_id"],
                "match_score": top["match_score"],
                "recommended_action": "confirm_match_or_reject",
                "review_status": "pending",
                "review_note": top["review_note"],
            })
        else:
            row.update({
                "match_status": "chart1_only",
                "node_status": "chart1_only",
                "review_flag": "yes",
                "review_note": node["chart1_notes"] or "chart2 had no safe active-company match",
            })
            reconciliation_rows.append({
                "issue_type": "chart1_only",
                "chart1_name": node["chart1_name"],
                "chart2_name": "",
                "candidate_node_id": node["node_id"],
                "match_score": "",
                "recommended_action": "check_if_chart2_missing_or_inactive",
                "review_status": "pending",
                "review_note": row["review_note"],
            })
        master_rows.append(row)

    for item in chart2_matches:
        if item["match_status"] == "chart2_only":
            chart2_only_rows.append({
                "chart2_name": item["chart2_name"],
                "legal_representative": item["legal_representative"],
                "established_date": item["established_date"],
                "registered_capital": item["registered_capital"],
                "actual_controller_share": item["actual_controller_share"],
                "subsidiary_level_label": item["subsidiary_level_label"],
                "company_status": item["company_status"],
                "reason_not_merged": item["review_note"],
            })
            reconciliation_rows.append({
                "issue_type": "chart2_only",
                "chart1_name": "",
                "chart2_name": item["chart2_name"],
                "candidate_node_id": "",
                "match_score": item["match_score"],
                "recommended_action": "check_if_chart1_missing_node",
                "review_status": "pending",
                "review_note": item["review_note"],
            })
        elif item["match_status"] == "review_match" and not item["candidate_node_id"]:
            reconciliation_rows.append({
                "issue_type": "review_match",
                "chart1_name": "",
                "chart2_name": item["chart2_name"],
                "candidate_node_id": "",
                "match_score": item["match_score"],
                "recommended_action": "manual_name_review",
                "review_status": "pending",
                "review_note": item["review_note"],
            })

    out_dir = BASE / "reconciliation_outputs"
    write_csv(
        out_dir / "master_nodes_enriched.csv",
        master_rows,
        [
            "node_id", "chart1_name", "canonical_name", "chart1_level", "chart1_parent",
            "chart1_parent_name", "matched_chart2_name", "legal_representative",
            "established_date", "registered_capital", "actual_controller_share",
            "subsidiary_level_label", "company_status", "match_status", "node_status",
            "review_flag", "review_note",
        ],
    )
    write_csv(
        out_dir / "reconciliation_report.csv",
        reconciliation_rows,
        [
            "issue_type", "chart1_name", "chart2_name", "candidate_node_id",
            "match_score", "recommended_action", "review_status", "review_note",
        ],
    )
    write_csv(
        out_dir / "chart2_only_candidates.csv",
        chart2_only_rows,
        [
            "chart2_name", "legal_representative", "established_date", "registered_capital",
            "actual_controller_share", "subsidiary_level_label", "company_status",
            "reason_not_merged",
        ],
    )
    write_csv(
        out_dir / "chart2_match_results.csv",
        chart2_matches,
        [
            "chart2_name", "legal_representative", "established_date", "registered_capital",
            "actual_controller_share", "subsidiary_level_label", "company_status", "notes",
            "candidate_node_id", "candidate_chart1_name", "chart2_level_num",
            "match_method", "match_score", "match_status", "review_note",
        ],
    )


if __name__ == "__main__":
    main()
