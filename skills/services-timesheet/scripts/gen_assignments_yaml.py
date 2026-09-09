#!/usr/bin/env python3
"""
gen_assignments_yaml.py
Generate assignments.yaml from SAP ISP Staffing export Excel.

Usage:
  python3 gen_assignments_yaml.py staffing-export.xlsx \
    --output <skill-disk-path>/assets/assignments.yaml

Source: https://isp.hec.net.sap/sap/bc/webdynpro/sap/zui_staffinglist?&sap-language=EN#
Navigate to the Export Excel icon on the far right of the screen, export the .xlsx file,
then run this script.

Expected columns (ISP staffing export format):
  Hide | Cost Object Type | Cost Object | Description Cost Object | Customer ID |
  Customer Name | Start date | End Date | Ty. | Task lvl | Resp. CCtr |
  Rec. Act. Days | Est. Act. Days | Rem. Act. Days | Rec. Act. Days % |
  Responsible Person | UNote | Customer PO # | Description

Note: WBS element rows are skipped. Only Order and Sales document item rows are imported.
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)

TYPE_MAP = {
    "Order":               "Order",
    "Sales document item": "SalesDoc",
}

SUPPORTED_TYPES = set(TYPE_MAP.keys())


def fmt_date(v):
    """Return YYYY-MM-DD string from a datetime object or string."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    try:
        return datetime.strptime(str(v), "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
    except ValueError:
        return str(v)


def infer_keywords(customer, description, project):
    kw = []
    if customer and customer.strip():
        c = customer.strip()
        kw.append(c)
        first = c.split()[0] if c.split() else ""
        if len(first) > 3 and first not in kw:
            kw.append(first)
    for m in re.findall(r"[A-Z][A-Za-z]{2,}", description or ""):
        if m not in kw and m.lower() not in ("and", "the", "for", "con", "sal"):
            kw.append(m)
    for seg in re.split(r"[:/\s]+", project or ""):
        seg = seg.strip()
        if 2 <= len(seg) <= 10 and seg.isupper() and seg not in kw:
            kw.append(seg)
    return kw[:6]


def parse_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return [], 0

    header = [str(h).strip() if h else "" for h in rows[0]]
    idx = {h: i for i, h in enumerate(header)}

    def cell(row, col):
        i = idx.get(col)
        return row[i] if i is not None and i < len(row) else None

    out = []
    skipped_wbs = 0
    for row in rows[1:]:
        raw_type = str(cell(row, "Cost Object Type") or "").strip()
        cost_obj = str(cell(row, "Cost Object") or "").strip()
        if not raw_type or not cost_obj:
            continue  # skip blank / totals rows
        if raw_type == "WBS element":
            skipped_wbs += 1
            continue  # WBS not supported
        if raw_type not in SUPPORTED_TYPES:
            continue

        customer    = str(cell(row, "Customer Name") or "").strip()
        description = str(cell(row, "Description Cost Object") or "").strip()
        project     = str(cell(row, "Description") or "").strip()
        responsible = str(cell(row, "Responsible Person") or "").strip()
        start_d     = fmt_date(cell(row, "Start date"))
        end_d       = fmt_date(cell(row, "End Date"))

        # Infer customer from description if blank (common for Orders)
        if not customer:
            m = re.search(r"-\s+([A-Z][a-zA-Z\s&]+?)(?:\s+-|$)", description)
            if m:
                customer = m.group(1).strip()

        entry = {
            "cost_object": cost_obj,
            "type":        TYPE_MAP[raw_type],
            "description": description,
            "customer":    customer,
            "project":     project,
            "start_date":  start_d,
            "end_date":    end_d,
            "responsible": responsible,
            "keywords":    infer_keywords(customer, description, project),
        }
        out.append(entry)
    return out, skipped_wbs


def to_yaml(assignments, source_file):
    lines = [
        "# SAP CATS Assignment List",
        f"# Generated from ISP Staffing export: {source_file}",
        f"# Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "#",
        "# To update: export from the ISP staffing portal (Export Excel icon, far right),",
        "# then tell Joule: 'import my assignments from <filename>.xlsx'",
        "#",
        "# TIP: add extra keywords (abbreviations, email domains) after import for better auto-matching.",
        "",
        "assignments:",
    ]

    for a in assignments:
        lines.append(f"  - cost_object: \"{a['cost_object']}\"")
        lines.append(f"    type: \"{a['type']}\"")
        lines.append(f"    description: \"{a['description']}\"")
        if a["customer"]:
            lines.append(f"    customer: \"{a['customer']}\"")
        if a["project"]:
            lines.append(f"    project: \"{a['project']}\"")
        if a["start_date"]:
            lines.append(f"    start_date: \"{a['start_date']}\"")
        if a["end_date"]:
            lines.append(f"    end_date: \"{a['end_date']}\"")
        if a["responsible"]:
            lines.append(f"    responsible: \"{a['responsible']}\"")
        if a["keywords"]:
            lines.append("    keywords:")
            for kw in a["keywords"]:
                lines.append(f"      - \"{kw}\"")
        lines.append("")

    return "\n".join(lines)


def main():
    p = argparse.ArgumentParser(description="Generate assignments.yaml from ISP staffing export Excel")
    p.add_argument("excel",    help="Path to ISP staffing export .xlsx")
    p.add_argument("--output", default="assignments.yaml")
    args = p.parse_args()

    path = Path(args.excel)
    if not path.exists():
        print(f"ERROR: {path} not found")
        sys.exit(1)

    assignments, skipped_wbs = parse_excel(path)
    print(f"Parsed {len(assignments)} assignments from {path.name}")
    if skipped_wbs:
        print(f"  Skipped {skipped_wbs} WBS element row(s) (not supported)")
    for a in assignments:
        print(f"  {a['type']:8} {a['cost_object']:25} {a.get('customer', '')}")

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(to_yaml(assignments, path.name), encoding="utf-8")
    print(f"\nWritten to: {out}")
    print("TIP: Open the YAML and add extra keywords (abbreviations, email domains) for better matching.")


if __name__ == "__main__":
    main()
