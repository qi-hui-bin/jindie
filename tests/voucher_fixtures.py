import argparse
import json
from pathlib import Path

from jindie.kingdee_export import HEADERS, IMPORT_SHEET

from .fixtures import boc_rows, make_xls
from openpyxl import Workbook


def write_example(directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    rows = boc_rows()
    for i, original in enumerate(rows[9:], 1):
        row = list(original)
        row[10] = "收货款" if i <= 37 else "银行手续费" if i <= 43 else "付货款"
        if 38 <= i <= 43:
            row[2] = f"202611{12 + (i - 38) // 2}"
        if i >= 48:
            row[10] = "同账户划转"
        rows[i + 8] = tuple(row)
    with (directory / "boc-voucher-synthetic.xls").open("xb") as file:
        file.write(make_xls(rows))
    config = {
        "ledger": {
            "book_id": "demo",
            "version": 1,
            "bank_subjects": {"boc": "100203"},
            "subjects": [
                {"code": "100203", "name": "合成银行存款"},
                {"code": "1122", "name": "合成应收账款", "required_aux": ["客户"]},
                {"code": "2202", "name": "合成应付账款", "required_aux": ["供应商"]},
                {"code": "560303", "name": "合成手续费"},
            ],
            "auxiliary": [{"category": "客户", "code": "001"}, {"category": "供应商", "code": "029"}],
            "legacy_assignments": {
                "2202_029": {"subject": "2202", "auxiliary": [{"category": "供应商", "code": "029"}]},
            },
        },
        "rules": [
            {
                "id": name, "version": 1, "book_id": "demo", "account_id": "boc", "direction": direction,
                "conditions": [{"field": "purpose", "operation": "equals", "values": [purpose]}],
                "counterpart": assignment, "summary": purpose, "merge": merge, "enabled": True,
            }
            for name, direction, purpose, assignment, merge in (
                ("receive", "income", "收货款", {"subject": "1122", "auxiliary": [{"category": "客户", "code": "001"}]}, "none"),
                ("pay", "expense", "付货款", {"subject": "2202_029"}, "none"),
                ("fee", "expense", "银行手续费", {"subject": "560303"}, "bank_detail"),
            )
        ],
        "web": {
            "label": "合成演示账套",
            "accounts": [{"id": "boc", "number": "0000000000009130", "label": "合成中行"}],
            "template": "synthetic-template.xlsx",
        },
    }
    for name, data in (
        ("voucher-config.json", config),
        ("confirmed-skips.json", {"000048": "合成样本已入账 A", "000049": "合成样本已入账 B"}),
    ):
        with (directory / name).open("x", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)
    workbook = Workbook()
    workbook.active.title = IMPORT_SHEET
    workbook.active.append(HEADERS)
    workbook.create_sheet("合成备查").append(["用于测试的备查页"])
    with (directory / "synthetic-template.xlsx").open("xb") as file:
        workbook.save(file)
    workbook.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="生成合成中行流水、凭证规则和空白模板。")
    parser.add_argument("directory", type=Path)
    write_example(parser.parse_args().directory)
