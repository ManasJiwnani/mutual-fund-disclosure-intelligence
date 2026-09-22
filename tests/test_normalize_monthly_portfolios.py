from datetime import date

from openpyxl import Workbook

from scripts.normalize_monthly_portfolios import normalize_workbook


def test_normalizes_provenance_and_units(tmp_path):
    path = tmp_path / "INF123456789_Test Fund.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Portfolio as on 31-Aug-2026"])
    sheet.append(["Name of Instrument", "ISIN Code", "Industry", "Quantity", "Market Value (Rs. in Lacs)", "% to NAV"])
    sheet.append(["Equity", None, None, None, None, None])
    sheet.append(["ACME LTD", "INE123456789", "Banks", "10", "2.5", "0.01"])
    workbook.save(path)

    records = normalize_workbook(path)

    assert len(records) == 1
    record = records[0]
    assert record["fund_isin"] == "INF123456789"
    assert record["portfolio_date"] == date(2026, 8, 31).isoformat()
    assert record["section"] == "Equity"
    assert record["market_value_inr"] == 250000
    assert record["pct_of_nav"] == 1
    assert record["source_sheet"] == "Sheet"
    assert record["source_row"] == 4