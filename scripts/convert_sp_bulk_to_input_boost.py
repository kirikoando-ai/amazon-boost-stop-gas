#!/usr/bin/env python3
"""Convert Amazon SP bulk XLSX (sheet2) into input_boost CSV.

Usage:
  python3 scripts/convert_sp_bulk_to_input_boost.py \
    --input /path/to/bulk.xlsx \
    --output /path/to/input_boost.csv
"""

import argparse
import csv
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
SHEET_PATH = "xl/worksheets/sheet2.xml"  # スポンサープロダクト広告キャンペーン

OUT_HEADERS = [
    "boost_campaign_id",
    "boost_campaign_name",
    "boost_ad_group_id",
    "boost_ad_group_name",
    "product_key",
    "product_label",
    "target_type",
    "target_text",
    "match_type",
    "impressions",
    "clicks",
    "orders",
    "sales",
    "spend",
    "cpc",
    "target_key",
    "existing_exact_in_normal",
]

H = {
    "product": "プロダクト",
    "entity": "エンティティ",
    "campaign_id": "キャンペーンID",
    "ad_group_id": "広告グループID",
    "campaign_name": "キャンペーン名",
    "campaign_name_info": "キャンペーン名（情報提供のみ）",
    "ad_group_name": "広告グループ名",
    "ad_group_name_info": "広告グループ名（情報提供のみ）",
    "target_kw": "キーワードテキスト",
    "target_pt": "商品ターゲティング式",
    "target_pt_resolved": "解決済みの商品ターゲティング式（情報提供のみ）",
    "match_type": "マッチタイプ",
    "impressions": "インプレッション数",
    "clicks": "クリック数",
    "orders": "注文",
    "sales": "売上",
    "spend": "支出",
    "cpc": "CPC",
    "asin": "ASIN（情報提供のみ）",
    "sku": "SKU",
}


def col_letters(cell_ref: str) -> str:
    m = re.match(r"([A-Z]+)", cell_ref)
    return m.group(1) if m else ""


def get_cell_text(cell: ET.Element) -> str:
    t = cell.attrib.get("t")
    if t == "inlineStr":
        node = cell.find("x:is/x:t", NS)
        return node.text if node is not None and node.text is not None else ""
    v = cell.find("x:v", NS)
    return v.text if v is not None and v.text is not None else ""


def convert(input_xlsx: Path, output_csv: Path) -> int:
    header_by_col = {}
    rows_written = 0

    with zipfile.ZipFile(input_xlsx, "r") as zf:
        with zf.open(SHEET_PATH) as sheet_xml, output_csv.open("w", newline="", encoding="utf-8-sig") as wf:
            writer = csv.DictWriter(wf, fieldnames=OUT_HEADERS)
            writer.writeheader()

            for _, elem in ET.iterparse(sheet_xml, events=("end",)):
                if elem.tag != "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row":
                    continue

                rnum = int(elem.attrib.get("r", "0") or 0)
                row_map = {}
                for c in elem.findall("x:c", NS):
                    ref = c.attrib.get("r", "")
                    row_map[col_letters(ref)] = get_cell_text(c)

                if rnum == 1:
                    for col, val in row_map.items():
                        header_by_col[col] = val
                    elem.clear()
                    continue

                d = {}
                for col, val in row_map.items():
                    h = header_by_col.get(col, "")
                    if h:
                        d[h] = val

                if d.get(H["product"], "").strip() != "スポンサープロダクト広告":
                    elem.clear()
                    continue

                ent = d.get(H["entity"], "").strip()
                if ent == "キーワード":
                    target_type = "keyword"
                    target_text = d.get(H["target_kw"], "").strip()
                elif ent == "商品ターゲティング":
                    target_type = "product_targeting"
                    target_text = (
                        d.get(H["target_pt"], "").strip()
                        or d.get(H["target_pt_resolved"], "").strip()
                    )
                else:
                    elem.clear()
                    continue

                if not target_text:
                    elem.clear()
                    continue

                campaign_name = (
                    d.get(H["campaign_name"], "").strip()
                    or d.get(H["campaign_name_info"], "").strip()
                )
                ad_group_name = (
                    d.get(H["ad_group_name"], "").strip()
                    or d.get(H["ad_group_name_info"], "").strip()
                )
                product_key = (
                    d.get(H["asin"], "").strip()
                    or d.get(H["sku"], "").strip()
                    or ad_group_name
                    or d.get(H["campaign_id"], "").strip()
                )
                product_label = d.get(H["sku"], "").strip()

                writer.writerow(
                    {
                        "boost_campaign_id": d.get(H["campaign_id"], "").strip(),
                        "boost_campaign_name": campaign_name,
                        "boost_ad_group_id": d.get(H["ad_group_id"], "").strip(),
                        "boost_ad_group_name": ad_group_name,
                        "product_key": product_key,
                        "product_label": product_label,
                        "target_type": target_type,
                        "target_text": target_text,
                        "match_type": d.get(H["match_type"], "").strip(),
                        "impressions": d.get(H["impressions"], "").strip(),
                        "clicks": d.get(H["clicks"], "").strip(),
                        "orders": d.get(H["orders"], "").strip(),
                        "sales": d.get(H["sales"], "").strip(),
                        "spend": d.get(H["spend"], "").strip(),
                        "cpc": d.get(H["cpc"], "").strip(),
                        "target_key": "",
                        "existing_exact_in_normal": "",
                    }
                )
                rows_written += 1
                elem.clear()

    return rows_written


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to Amazon bulk XLSX")
    parser.add_argument("--output", required=True, help="Path to output CSV for input_boost")
    args = parser.parse_args()

    input_xlsx = Path(args.input)
    output_csv = Path(args.output)

    if not input_xlsx.exists():
        raise FileNotFoundError(f"Input file not found: {input_xlsx}")

    output_csv.parent.mkdir(parents=True, exist_ok=True)
    rows = convert(input_xlsx, output_csv)
    print(f"WROTE_ROWS={rows}")
    print(f"OUTPUT={output_csv}")


if __name__ == "__main__":
    main()
