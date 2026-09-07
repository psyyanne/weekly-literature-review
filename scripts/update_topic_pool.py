"""Synchronize three to five source-traceable candidate topics per deep reading."""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from build_deep_read_docx import topic_candidates
from common import TOPIC_COLUMNS, append_change, atomic_save, project_root, upsert_dict, worksheet_records


def update(root: Path, bundle_path: Path) -> dict[str, int]:
    bundle = json.loads(bundle_path.read_text(encoding="utf-8"))
    wb = load_workbook(root / "library/master_literature_review.xlsx")
    lit_rows = worksheet_records(wb["文献表"])
    by_doi = {str(r.get("DOI") or "").lower(): r for r in lit_rows}
    ws = wb["选题池"]
    counts = {"added": 0, "updated": 0}
    for deep in bundle.get("deep_readings") or []:
        lit = by_doi.get(str(deep.get("DOI") or "").lower())
        if not lit:
            continue
        source_id = str(lit["文献编号"])
        direction = str(deep.get("适合拓展的心理学方向") or deep.get("所属心理学方向") or deep.get("研究背景与领域位置") or "人工智能与心理学")
        for index, topic in enumerate(topic_candidates(deep), 1):
            topic_id = f"{source_id}-T{index:02d}"
            row = {
                "选题编号": topic_id,
                "来源文献编号": source_id,
                "备选研究题目": topic["title"],
                "所属心理学方向": direction,
                "研究缺口": topic["gap"],
                "理论基础": topic["theory"],
                "核心变量": topic["vars"],
                "研究对象": topic["sample"],
                "方法建议": topic["method"],
                "相比原文的推进": topic["advance"],
                "创新点": topic["novelty"],
                "可行性": topic["feasibility"],
                "需要继续核验的问题": topic["verify"],
                "首次提出日期": date.today().isoformat(),
                "最近更新日期": date.today().isoformat(),
                "选题状态": "新提出",
                "备注": "由全文精读证据生成；需与其他文献关联后再进入稳定研究计划。",
            }
            state, _ = upsert_dict(ws, "选题编号", row)
            counts[state] += 1
            append_change(wb["修改日志"], topic_id, state, "选题池", "", row["备选研究题目"], str(bundle_path), "update_topic_pool.py")
    atomic_save(wb, root / "library/master_literature_review.xlsx", make_backup=True)
    return counts


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("bundle_json")
    parser.add_argument("--project-root")
    args = parser.parse_args()
    result = update(project_root(args.project_root), Path(args.bundle_json))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
