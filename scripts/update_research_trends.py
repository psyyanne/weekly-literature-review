"""Update cumulative trend evidence and keep weak evidence labeled as clues."""
from __future__ import annotations

import argparse
import json
import re
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from common import append_change, atomic_save, project_root, upsert_dict, worksheet_records


def split_ids(value: str) -> list[str]:
    return [x.strip() for x in re.split(r"[；;,，\s]+", str(value or "")) if x.strip()]


def update(root: Path, bundle_path: Path) -> dict[str, int]:
    bundle = json.loads(bundle_path.read_text(encoding="utf-8"))
    wb = load_workbook(root / "library/master_literature_review.xlsx")
    lit_rows = worksheet_records(wb["文献表"])
    by_doi = {str(r.get("DOI") or "").lower(): r for r in lit_rows}
    old_to_new = {}
    for item in bundle.get("literature_records") or []:
        row = by_doi.get(str(item.get("DOI") or "").lower())
        if row and item.get("文献ID"):
            old_to_new[str(item["文献ID"])] = str(row["文献编号"])
    ws = wb["研究趋势"]
    counts = {"added": 0, "updated": 0}
    for index, trend in enumerate(bundle.get("trends") or [], 1):
        support = [old_to_new.get(x, x) for x in split_ids(trend.get("支持文献"))]
        weeks = split_ids(trend.get("支持周次"))
        deep_support = split_ids(trend.get("重点精读支持"))
        evidence = "强" if len(set(support)) >= 3 and len(set(deep_support)) >= 2 else "初步"
        state = "研究趋势" if evidence == "强" and len(set(weeks)) >= 2 else ("本周线索" if len(set(support)) <= 2 else "初步迹象")
        trend_id = re.sub(r"^TREND-", "", str(trend.get("趋势ID") or f"{date.today().isoformat()}-{index:02d}"))
        trend_id = f"TREND-{trend_id.replace('2026W31', '2026-W31')}"
        row = {
            "趋势编号": trend_id,
            "趋势名称": trend.get("趋势名称") or "待命名趋势",
            "所属心理学方向": trend.get("所属心理学方向") or "人工智能与心理学",
            "支持文献编号": "；".join(dict.fromkeys(support)),
            "主要理论": trend.get("理论变化") or trend.get("主要理论") or "待继续核验",
            "主要变量": trend.get("变量变化") or trend.get("主要变量") or "待继续核验",
            "常用研究方法": trend.get("方法变化") or trend.get("常用研究方法") or "待继续核验",
            "主要结论": trend.get("一致结果") or trend.get("主要结论") or "待继续核验",
            "不一致结果": trend.get("不一致结果") or "",
            "近期变化": trend.get("近期变化") or trend.get("样本变化") or "待继续观察",
            "研究缺口": trend.get("研究缺口") or "待继续核验",
            "可以继续发展的方向": trend.get("可以继续发展的方向") or "待继续核验",
            "证据强度": evidence,
            "趋势状态": state,
            "最近更新日期": date.today().isoformat(),
            "备注": "仅在累计证据达到门槛后升级为稳定趋势；当前状态按支持文献和跨周证据自动判断。",
        }
        action, _ = upsert_dict(ws, "趋势编号", row)
        counts[action] += 1
        append_change(wb["修改日志"], trend_id, action, "研究趋势", "", row["趋势名称"], str(bundle_path), "update_research_trends.py")
    atomic_save(wb, root / "library/master_literature_review.xlsx", make_backup=True)
    return counts


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("bundle_json")
    parser.add_argument("--project-root")
    args = parser.parse_args()
    print(json.dumps(update(project_root(args.project_root), Path(args.bundle_json)), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
