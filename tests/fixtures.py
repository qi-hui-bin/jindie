import argparse
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import xlwt
from openpyxl import Workbook

BOC_ACCOUNT = "0000000000009130"
CCB_ACCOUNT = "0000000000000306"
BOC_HEADERS = (
    "交易类型", "交易金额", "交易日期", "交易时间", "起息日期", "付款人名称", "付款人账号",
    "收款人名称", "收款人账号", "摘要", "用途", "交易附言", "备注", "交易流水号", "交易后余额", "交易货币",
)
CCB_HEADERS = (
    "账号", "借方发生额(支取)", "贷方发生额（收入）", "交易时间", "记账日期", "对方户名",
    "对方账号", "摘要", "备注", "账户明细编号-交易流水号", "余额", "币种",
)


def make_xls(rows, title="Sheet1") -> bytes:
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(title)
    date_style = xlwt.easyxf(num_format_str="yyyy-mm-dd hh:mm:ss")
    for y, row in enumerate(rows):
        for x, value in enumerate(row):
            if value is not None:
                if isinstance(value, date):
                    sheet.write(y, x, value, date_style)
                else:
                    sheet.write(y, x, value)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def make_xlsx(rows, title="数据", *, second_sheet=False, epoch=None) -> bytes:
    workbook = Workbook()
    if epoch:
        workbook.epoch = epoch
    workbook.active.title = title
    for row in rows:
        workbook.active.append(row)
    if second_sheet:
        second = workbook.create_sheet("另一账户")
        for row in rows:
            second.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def boc_rows() -> list[tuple]:
    result = [("流水查询",), ("查询账号", BOC_ACCOUNT), *([()] * 6), BOC_HEADERS]
    balance = 0
    for index in range(49):
        income = index < 37
        day = min(12 + index // 7, 18)
        minor = 150_000 + index * 100 if income else (900, 90, 330, 250, 200, 400)[(index - 37) % 6]
        balance += minor if income else -minor
        counterpart = f"000000000000{index:04d}"
        result.append((
            "来账" if income else "往账", minor / 100 if income else -minor / 100, f"202611{day:02d}",
            f"09:{index:02d}:01", f"2026-11-{day:02d}",
            f"合成对方{index}" if income else "合成本方", counterpart if income else BOC_ACCOUNT,
            "合成本方" if income else f"合成对方{index}", BOC_ACCOUNT if income else counterpart,
            "收款" if income else "付款", "测试用途", None, "测试备注",
            f"{index + 1:06d}", balance / 100, "CNY",
        ))
    return result


def ccb_row(*, account=CCB_ACCOUNT, expense=None, income=10, tx="2026-11-02 8:30:00",
            posting="2026/11/02", reference="000001", remark="货款", balance=0) -> tuple:
    return (account, expense, income, tx, posting, "合成对方", "0000123", "电子转账", remark,
            reference, balance, "人民币元")


def ccb_rows() -> list[tuple]:
    result = [CCB_HEADERS]
    for index in range(91):
        day = datetime(2026, 11, 2, 8, 30) + timedelta(days=index * 18 // 90, minutes=index)
        income = index < 32
        result.append(ccb_row(expense=None if income else 100 + index, income=100 + index if income else 0,
                              tx=day, posting=day.date(), reference=f"{index + 1:06d}"))
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="生成仅包含虚构数据的银行流水测试样本。")
    parser.add_argument("directory", type=Path)
    target = parser.parse_args().directory
    target.mkdir(parents=True, exist_ok=True)
    for name, data in (
        ("boc-synthetic.xls", make_xls(boc_rows())),
        ("ccb-synthetic.xls", make_xls(ccb_rows(), "合成建行0306")),
        ("ccb-synthetic.xlsx", make_xlsx(ccb_rows(), "合成建行0306")),
    ):
        # Refuse to overwrite a pre-existing document.
        with (target / name).open("xb") as output:
            output.write(data)


if __name__ == "__main__":
    main()
