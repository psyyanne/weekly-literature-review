"""Validate a candidate JSON pool without inventing metadata or downloading papers."""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

from common import normalize_doi, project_root

DOI_RE = re.compile(r"^10\.\d{4,9}/\S+$", re.I)


def classify(record: dict[str, Any]) -> str:
    journal = str(record.get("期刊") or "").lower()
    kind = str(record.get("文章类型") or "").lower()
    if "frontiers" in journal:
        return "F"
    if any(x in kind for x in ("preprint", "预印本", "conference", "会议")) or not record.get("DOI"):
        return "O"
    if any(x in journal for x in ("digital medicine", "social sciences", "medical", "medicine", "health", "jmir", "jama network open", "education technology", "computer", "information", "management", "human-computer")):
        return "X"
    return "C"


def validate(records: list[dict[str, Any]]) -> dict[str, Any]:
    errors: list[str] = []
    seen: set[str] = set()
    output = []
    for index, record in enumerate(records, 1):
        label = f"第{index}条"
        aliases = {"英文题名": ("英文题名", "英文标题"), "第一作者": ("第一作者", "作者"), "期刊": ("期刊", "来源或期刊"), "正式发表日期": ("正式发表日期", "Online first日期"), "文章类型": ("文章类型", "材料类型")}
        for field, names in aliases.items():
            if not any(str(record.get(name) or "").strip() for name in names):
                errors.append(f"{label}缺少{field}")
        doi = normalize_doi(record.get("DOI"))
        if doi and not DOI_RE.match(doi):
            errors.append(f"{label} DOI格式错误")
        if doi in seen and doi:
            errors.append(f"{label} DOI重复: {doi}")
        if doi:
            seen.add(doi)
        if record.get("原文摘要") in (None, ""):
            record["原文摘要"] = ""
        if record.get("中文摘要") in (None, ""):
            record["中文摘要"] = ""
        record["文献类别"] = classify(record)
        output.append(record)
    return {"valid": not errors, "errors": errors, "records": output, "counts": {c: sum(r["文献类别"] == c for r in output) for c in "CFXO"}}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("candidates_json")
    parser.add_argument("--output")
    parser.add_argument("--project-root")
    args = parser.parse_args()
    path = Path(args.candidates_json)
    payload = json.loads(path.read_text(encoding="utf-8"))
    records = payload.get("literature_records", payload.get("records", payload if isinstance(payload, list) else []))
    result = validate(records)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if result["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
