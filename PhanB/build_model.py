"""
Phần B - Model-Driven DSS (Excel): dựng workbook 4 bài toán từ số liệu Phần A.

Chạy (từ thư mục PhanB, trong venv dựng từ requirements.txt):
    python3 build_model.py
Đầu vào : ../PhanA/tableau_data/*.csv  (đầu ra của etl_phanA.py)
Đầu ra  : ./MaNhom_Model.xlsx          (đổi MaNhom thành mã nhóm khi nộp)

Workbook gồm: TongQuan, DuLieu, BT1_Solver, BT2_WhatIf, BT3_GoalSeek, BT4_Solver.
Mọi ô tính toán là công thức Excel thật; số liệu gốc ở DuLieu là giá trị tĩnh
tổng hợp từ CSV. Cấu hình Solver được ghi sẵn dưới dạng tên ẩn của sheet - mở
Data > Solver sẽ thấy sẵn tham số (cần kiểm tra lại một lần trong Excel).
"""
import os

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import DataTableFormula
from openpyxl.worksheet.scenario import InputCells, Scenario, ScenarioList
from scipy.optimize import linprog

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "PhanA", "tableau_data")
OUT = os.path.join(HERE, os.environ.get("MODEL_OUT", "MaNhom_Model.xlsx"))

# --------------------------------------------------------------------------
# Giả định (mỗi giá trị là một quyết định có lý do - xem PhanB_BaoCao.md)
# --------------------------------------------------------------------------
PENALTY_PER_LATE = 0.0      # USD / đơn trễ - mặc định 0: chỉ dùng lợi nhuận từ dữ liệu; nhập > 0 để thử
LATE_CAP = 0.45             # ngưỡng tỷ lệ trễ chung - mục tiêu quản trị, tự đặt
SHIFT_BAND = 0.15           # biên dịch chuyển tỷ trọng mỗi phương thức (±)
FIXED_COST_RATIO = 0.07     # chi phí cố định / doanh thu thuần ngành - tự đặt
DEMAND_BAND = 0.20          # biên nhu cầu mỗi thị trường (±)
MODE_ORDER = ["First Class", "Same Day", "Second Class", "Standard Class"]
DISCOUNT_BINS = [-0.001, 0.0, 0.05, 0.10, 0.15, 0.20, 0.25]
DISCOUNT_LABELS = ["0%", "(0%,5%]", "(5%,10%]", "(10%,15%]", "(15%,20%]", "(20%,25%]"]

# --------------------------------------------------------------------------
# Kiểu ô theo workbook mẫu của giảng viên: tiêu đề vàng đậm, tiêu đề cột xanh nhạt,
# ô biến quyết định / biến mục tiêu nền vàng, ô giả định nền xanh nhạt.
# --------------------------------------------------------------------------
YELLOW = PatternFill("solid", fgColor="FFFF00")
BLUE = PatternFill("solid", fgColor="E6F3FF")
GREY = PatternFill("solid", fgColor="F2F2F2")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
USD = '#,##0.00'
INT = '#,##0'
PCT = '0.00%'


FACTS = {}  # số liệu dùng trong các đoạn mô tả; tính từ CSV để không lệch khi ETL đổi


def vn(x, dec=1):
    """Số kiểu Việt Nam: dấu phẩy thập phân, dấu chấm hàng nghìn."""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def title(ws, text):
    ws["A1"] = text
    ws["A1"].font = BOLD
    ws["A1"].fill = YELLOW


def head(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = BOLD
        c.fill = BLUE
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")


def put(ws, addr, value, fmt=None, fill=None, bold=False):
    c = ws[addr]
    c.value = value
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if bold:
        c.font = BOLD
    c.border = BOX
    return c


def textbox(ws, rng, text):
    ws.merge_cells(rng)
    first = rng.split(":")[0]
    ws[first] = text
    ws[first].alignment = WRAP


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def steps(ws, row, lines):
    ws.cell(row=row, column=1, value="Hướng dẫn").font = BOLD
    for i, t in enumerate(lines):
        ws.cell(row=row + 1 + i, column=1, value=t)


def solver_names(ws, opt, adj, kind, rows):
    """Ghi tên ẩn solver_* để hộp thoại Solver mở ra đã điền sẵn.
    kind: 'max' | 'min'; rows: [(lhs, rel, rhs)] với rel 1:<=, 2:=, 3:>=.
    Engine 2 = Simplex LP (mô hình đều tuyến tính)."""
    d = {
        "opt": opt, "adj": adj, "typ": 1 if kind == "max" else 2, "val": 0,
        "eng": 2, "lin": 1, "neg": 1, "num": len(rows),
        "cvg": 0.0001, "drv": 1, "itr": 2147483647, "mip": 2147483647,
        "mni": 30, "mrt": 0.075, "msl": 2, "nod": 2147483647, "pre": 0.000001,
        "rbv": 1, "rlx": 1, "rsd": 0, "scl": 2, "sho": 2, "ssz": 100,
        "tim": 2147483647, "tol": 0.01, "ver": 2,
    }
    for i, (lhs, rel, rhs) in enumerate(rows, start=1):
        d[f"lhs{i}"], d[f"rel{i}"], d[f"rhs{i}"] = lhs, rel, rhs
    for k, v in d.items():
        name = f"solver_{k}"
        ws.defined_names[name] = DefinedName(name, attr_text=str(v), hidden=True)


# --------------------------------------------------------------------------
# 1. Tổng hợp số liệu từ CSV Phần A
# --------------------------------------------------------------------------
def load():
    fs = pd.read_csv(os.path.join(SRC, "fact_sales.csv"))
    fo = pd.read_csv(os.path.join(SRC, "fact_order_delivery.csv"))
    dl = pd.read_csv(os.path.join(SRC, "dim_delivery.csv"))
    lc = pd.read_csv(os.path.join(SRC, "dim_location.csv"))
    pr = pd.read_csv(os.path.join(SRC, "dim_product.csv"))

    o = fo.merge(dl, on="delivery_key").merge(lc, on="location_key")
    mode = (o.groupby("shipping_mode")
            .agg(n=("order_id", "count"), profit=("order_profit", "mean"),
                 late=("late_delivery_flag", "mean"), delay=("shipping_delay", "mean"))
            .reindex(MODE_ORDER).reset_index())
    market = (o.groupby("market")
              .agg(n=("order_id", "count"), sales=("order_sales", "mean"),
                   profit=("order_profit", "mean"), late=("late_delivery_flag", "mean"))
              .sort_values("n", ascending=False).reset_index())

    # làm tròn như số ghi vào Excel để nghiệm đối chiếu trùng với Solver
    mode = mode.round({"profit": 4, "late": 6, "delay": 4})
    market = market.round({"sales": 4, "profit": 4, "late": 6})

    f = fs.merge(pr, on="product_key")
    dept = (f.groupby("department_name")
            .agg(lines=("order_item_id", "count"), units=("quantity", "sum"),
                 gross=("sales", "sum"), disc=("discount", "sum"),
                 net=("net_sales", "sum"), profit=("profit", "sum"))
            .sort_values("net", ascending=False).reset_index())

    f["band"] = pd.cut(f["discount_rate"], DISCOUNT_BINS, labels=DISCOUNT_LABELS)
    band = (f.groupby("band", observed=True)
            .agg(rate=("discount_rate", "mean"), lines=("order_item_id", "count"),
                 profit=("profit", "mean"), gross=("sales", "mean")).reset_index())
    overall = dict(lines=len(f), gross=f["sales"].mean(), d0=f["discount_rate"].mean(),
                   profit=f["profit"].mean())
    m = mode.set_index("shipping_mode")
    others = market[market["market"] != "Pacific Asia"]
    pac = market[market["market"] == "Pacific Asia"].iloc[0]
    d = dept.set_index("department_name")
    FACTS.update(
        orders=int(mode["n"].sum()), lines=len(f),
        late=(mode["n"] * mode["late"]).sum() / mode["n"].sum() * 100,
        late_fc=m.loc["First Class", "late"] * 100, late_sc=m.loc["Second Class", "late"] * 100,
        d0=overall["d0"] * 100, corr=round(float(f["discount_rate"].corr(f["quantity"])), 3) + 0.0,
        pac_profit=pac["profit"], pac_sales=pac["sales"],
        oth_profit_min=others["profit"].min(), oth_profit_max=others["profit"].max(),
        oth_sales_min=others["sales"].min(), oth_sales_max=others["sales"].max(),
        margin_book=d.loc["Book Shop", "profit"] / d.loc["Book Shop", "net"] * 100,
        margin_fan=d.loc["Fan Shop", "profit"] / d.loc["Fan Shop", "net"] * 100,
        margin_all=dept["profit"].sum() / dept["net"].sum() * 100,
    )
    return mode, market, dept, band, overall


# --------------------------------------------------------------------------
# 2. Sheet DuLieu
# --------------------------------------------------------------------------
def sheet_data(wb, mode, market, dept, band, overall):
    ws = wb.create_sheet("DuLieu")
    title(ws, "Dữ liệu tham chiếu cho Phần B (tổng hợp từ CSV Phần A)")
    ws["A2"] = (f"Nguồn: PhanA/tableau_data/*.csv, tính trên {vn(FACTS['orders'], 0)} đơn / "
                f"{vn(FACTS['lines'], 0)} dòng, đã loại đơn CANCELED và SUSPECTED_FRAUD - khớp số liệu "
                "Phần A. Giá trị tĩnh; chạy lại build_model.py nếu ETL đổi.")

    # A. Phương thức vận chuyển: hàng 5-8, tổng hàng 9
    ws["A3"] = "A. Phương thức vận chuyển (mức đơn hàng)"
    ws["A3"].font = BOLD
    head(ws, 4, ["Phương thức", "Số đơn", "Tỷ trọng", "Lợi nhuận/đơn (USD)", "Tỷ lệ trễ", "Độ trễ TB (ngày)"])
    for i, r in mode.iterrows():
        row = 5 + i
        put(ws, f"A{row}", r["shipping_mode"])
        put(ws, f"B{row}", int(r["n"]), INT)
        put(ws, f"C{row}", f"=B{row}/$B$9", PCT)
        put(ws, f"D{row}", round(float(r["profit"]), 4), USD)
        put(ws, f"E{row}", round(float(r["late"]), 6), PCT)
        put(ws, f"F{row}", round(float(r["delay"]), 4), '0.000')
    put(ws, "A9", "Tổng / TB có trọng số", bold=True)
    put(ws, "B9", "=SUM(B5:B8)", INT, bold=True)
    put(ws, "C9", "=SUM(C5:C8)", PCT, bold=True)
    put(ws, "D9", "=SUMPRODUCT(B5:B8,D5:D8)/B9", USD, bold=True)
    put(ws, "E9", "=SUMPRODUCT(B5:B8,E5:E8)/B9", PCT, bold=True)

    # B. Thị trường: hàng 13-17, tổng hàng 18
    ws["A11"] = "B. Thị trường (mức đơn hàng)"
    ws["A11"].font = BOLD
    head(ws, 12, ["Thị trường", "Số đơn", "Tỷ trọng", "Doanh thu/đơn (USD)", "Lợi nhuận/đơn (USD)", "Tỷ lệ trễ"])
    for i, r in market.iterrows():
        row = 13 + i
        put(ws, f"A{row}", r["market"])
        put(ws, f"B{row}", int(r["n"]), INT)
        put(ws, f"C{row}", f"=B{row}/$B$18", PCT)
        put(ws, f"D{row}", round(float(r["sales"]), 4), USD)
        put(ws, f"E{row}", round(float(r["profit"]), 4), USD)
        put(ws, f"F{row}", round(float(r["late"]), 6), PCT)
    put(ws, "A18", "Tổng / TB có trọng số", bold=True)
    put(ws, "B18", "=SUM(B13:B17)", INT, bold=True)
    put(ws, "C18", "=SUM(C13:C17)", PCT, bold=True)
    put(ws, "D18", "=SUMPRODUCT(B13:B17,D13:D17)/B18", USD, bold=True)
    put(ws, "E18", "=SUMPRODUCT(B13:B17,E13:E17)/B18", USD, bold=True)
    put(ws, "F18", "=SUMPRODUCT(B13:B17,F13:F17)/B18", PCT, bold=True)

    # C. Ngành hàng: hàng 22-32, tổng hàng 33
    ws["A20"] = "C. Ngành hàng (mức dòng đơn)"
    ws["A20"].font = BOLD
    head(ws, 21, ["Ngành hàng", "Số dòng", "Số lượng (đơn vị)", "Doanh thu gộp", "Chiết khấu",
                  "Doanh thu thuần", "Lợi nhuận", "Giá thuần/đơn vị (p)", "Lợi nhuận/đơn vị",
                  "Biến phí/đơn vị (v)", "Biên lợi nhuận"])
    for i, r in dept.iterrows():
        row = 22 + i
        put(ws, f"A{row}", r["department_name"])
        put(ws, f"B{row}", int(r["lines"]), INT)
        put(ws, f"C{row}", int(r["units"]), INT)
        put(ws, f"D{row}", round(float(r["gross"]), 2), INT)
        put(ws, f"E{row}", round(float(r["disc"]), 2), INT)
        put(ws, f"F{row}", round(float(r["net"]), 2), INT)
        put(ws, f"G{row}", round(float(r["profit"]), 2), INT)
        put(ws, f"H{row}", f"=F{row}/C{row}", USD)
        put(ws, f"I{row}", f"=G{row}/C{row}", USD)
        put(ws, f"J{row}", f"=H{row}-I{row}", USD)
        put(ws, f"K{row}", f"=G{row}/F{row}", PCT)
    put(ws, "A33", "Tổng", bold=True)
    for col in "BCDEFG":
        put(ws, f"{col}33", f"=SUM({col}22:{col}32)", INT, bold=True)
    put(ws, "K33", "=G33/F33", PCT, bold=True)

    # D. Dải chiết khấu: hàng 37-42
    ws["A35"] = "D. Lợi nhuận theo dải chiết khấu (mức dòng đơn)"
    ws["A35"].font = BOLD
    head(ws, 36, ["Dải chiết khấu", "Chiết khấu TB của dải", "Số dòng", "Lợi nhuận/dòng (USD)", "Doanh thu gộp/dòng (USD)"])
    for i, r in band.iterrows():
        row = 37 + i
        put(ws, f"A{row}", str(r["band"]))
        put(ws, f"B{row}", round(float(r["rate"]), 6), PCT)
        put(ws, f"C{row}", int(r["lines"]), INT)
        put(ws, f"D{row}", round(float(r["profit"]), 4), USD)
        put(ws, f"E{row}", round(float(r["gross"]), 4), USD)
    ws["A44"] = "Tham số chung (toàn bộ dòng đơn)"
    ws["A44"].font = BOLD
    put(ws, "A45", "Tổng số dòng đơn (N0)")
    put(ws, "B45", int(overall["lines"]), INT)
    put(ws, "A46", "Doanh thu gộp TB/dòng (USD)")
    put(ws, "B46", round(overall["gross"], 4), USD)
    put(ws, "A47", "Chiết khấu TB hiện tại (d0)")
    put(ws, "B47", round(overall["d0"], 6), PCT)
    put(ws, "A48", "Lợi nhuận TB/dòng (USD)")
    put(ws, "B48", round(overall["profit"], 4), USD)
    widths(ws, {"A": 30, "B": 16, "C": 16, "D": 18, "E": 18, "F": 18, "G": 16, "H": 18, "I": 16, "J": 16, "K": 14})
    return ws


# --------------------------------------------------------------------------
# 3. BT1 - cơ cấu phương thức vận chuyển (Solver, LP)
# --------------------------------------------------------------------------
def solve_bt1(mode, penalty=PENALTY_PER_LATE):
    share0 = (mode["n"] / mode["n"].sum()).to_numpy()
    net = (mode["profit"] - penalty * mode["late"]).to_numpy()
    late = mode["late"].to_numpy()
    lo = np.maximum(0, share0 - SHIFT_BAND)
    hi = np.minimum(1, share0 + SHIFT_BAND)
    res = linprog(-net, A_ub=[late], b_ub=[LATE_CAP], A_eq=[np.ones(4)], b_eq=[1],
                  bounds=list(zip(lo, hi)), method="highs")
    assert res.success, f"BT1 vô nghiệm: {res.message}"
    n_total = int(mode["n"].sum())
    return dict(share0=share0, share=res.x, obj=float(n_total * -res.fun),
                obj0=float(n_total * share0 @ net), late=float(late @ res.x),
                late0=float(late @ share0), n_total=n_total)


def sheet_bt1(wb, mode, ref, ref20):
    ws = wb.create_sheet("BT1_Solver")
    title(ws, "Bài toán 1: Cơ cấu phương thức vận chuyển - tối đa lợi nhuận, kiểm soát tỷ lệ trễ")
    textbox(ws, "H1:N12",
            f"Bối cảnh: tỷ lệ giao trễ hiện {vn(FACTS['late'])}% (First Class {vn(FACTS['late_fc'], 0)}%, "
            f"Second Class {vn(FACTS['late_sc'], 0)}%). Ban điều hành "
            "đặt mục tiêu đưa tỷ lệ trễ chung về ≤ 45%. Câu hỏi: nên dịch chuyển tỷ trọng đơn giữa 4 "
            "phương thức thế nào để lợi nhuận cao nhất mà vẫn đạt ngưỡng trễ?\n\nTham số do người ra "
            "quyết định đặt (ô nền xanh): ngưỡng trễ; biên dịch chuyển ±15 điểm % vì khách chọn phương "
            "thức theo nhu cầu, không thể ép hoàn toàn; chi phí phạt mỗi đơn trễ (mặc định 0, nhập số "
            "dương để thử độ nhạy). Lợi nhuận/đơn và tỷ lệ trễ của mỗi phương thức giữ nguyên khi tỷ "
            "trọng đổi.")

    head(ws, 3, ["Phương thức"] + MODE_ORDER + ["Tổng"])
    labels = ["Tỷ trọng đơn (BIẾN QUYẾT ĐỊNH)", "Tỷ trọng hiện tại", "Tỷ trọng tối thiểu",
              "Tỷ trọng tối đa", "Lợi nhuận/đơn (USD)", "Tỷ lệ trễ", "Phạt/đơn (USD)",
              "Lợi nhuận ròng/đơn (USD)"]
    for i, t in enumerate(labels):
        put(ws, f"A{4 + i}", t, bold=(i == 0))
    for j, col in enumerate("BCDE"):
        dr = 5 + j  # hàng tương ứng trong DuLieu
        put(ws, f"{col}4", float(ref["share0"][j]), PCT, YELLOW, True)
        put(ws, f"{col}5", f"=DuLieu!C{dr}", PCT)
        put(ws, f"{col}6", f"=MAX(0,{col}5-$B$17)", PCT)
        put(ws, f"{col}7", f"=MIN(1,{col}5+$B$17)", PCT)
        put(ws, f"{col}8", f"=DuLieu!D{dr}", USD)
        put(ws, f"{col}9", f"=DuLieu!E{dr}", PCT)
        put(ws, f"{col}10", f"=$B$15*{col}9", USD)
        put(ws, f"{col}11", f"={col}8-{col}10", USD)
    put(ws, "F4", "=SUM(B4:E4)", PCT, bold=True)

    ws["A13"] = "Tham số do người ra quyết định đặt"
    ws["A13"].font = BOLD
    put(ws, "A14", "Tổng số đơn (N)")
    put(ws, "B14", "=DuLieu!B9", INT)
    put(ws, "A15", "Chi phí phạt mỗi đơn trễ (USD) (tham số đặt)")
    put(ws, "B15", PENALTY_PER_LATE, USD, BLUE)
    put(ws, "A16", "Ngưỡng tỷ lệ trễ tối đa (tham số đặt)")
    put(ws, "B16", LATE_CAP, PCT, BLUE)
    put(ws, "A17", "Biên dịch chuyển tỷ trọng ± (tham số đặt)")
    put(ws, "B17", SHIFT_BAND, PCT, BLUE)

    ws["A19"] = "Kết quả"
    ws["A19"].font = BOLD
    put(ws, "A20", "Tổng lợi nhuận ròng (BIẾN MỤC TIÊU)", bold=True)
    put(ws, "B20", "=$B$14*SUMPRODUCT(B4:E4,B11:E11)", USD, YELLOW, True)
    put(ws, "A21", "Tỷ lệ trễ chung")
    put(ws, "B21", "=SUMPRODUCT(B4:E4,B9:E9)", PCT)
    put(ws, "A22", "Tổng tỷ trọng")
    put(ws, "B22", "=SUM(B4:E4)", PCT)
    put(ws, "A23", "Lợi nhuận ròng hiện tại")
    put(ws, "B23", "=$B$14*SUMPRODUCT(B5:E5,B11:E11)", USD)
    put(ws, "A24", "Chênh lệch so với hiện tại (âm = chi phí để đạt ngưỡng trễ)")
    put(ws, "B24", "=B20-B23", USD)
    put(ws, "A25", "Tỷ lệ trễ hiện tại")
    put(ws, "B25", "=SUMPRODUCT(B5:E5,B9:E9)", PCT)

    ws["A27"] = "Ràng buộc (nhập vào hộp thoại Solver)"
    ws["A27"].font = BOLD
    head(ws, 28, ["Ràng buộc", "Ô", "Dấu", "Giá trị giới hạn"])
    cons = [("Tỷ trọng từng phương thức ≥ tối thiểu", "B4:E4", ">=", "B6:E6"),
            ("Tỷ trọng từng phương thức ≤ tối đa", "B4:E4", "<=", "B7:E7"),
            ("Tổng tỷ trọng = 1", "B22", "=", "1"),
            ("Tỷ lệ trễ chung ≤ ngưỡng", "B21", "<=", "B16")]
    for i, r in enumerate(cons):
        for j, v in enumerate(r):
            put(ws, f"{'ABCD'[j]}{29 + i}", v)

    ws["A34"] = "Nghiệm tối ưu đối chiếu (tính độc lập bằng quy hoạch tuyến tính)"
    ws["A34"].font = BOLD
    head(ws, 35, ["Chỉ tiêu"] + MODE_ORDER)
    put(ws, "A36", "Tỷ trọng tối ưu")
    for j, col in enumerate("BCDE"):
        put(ws, f"{col}36", round(float(ref["share"][j]), 6), PCT)
    put(ws, "A37", "Lợi nhuận ròng tối ưu (USD)")
    put(ws, "B37", round(ref["obj"], 2), USD)
    put(ws, "A38", "Tỷ lệ trễ tại phương án tối ưu")
    put(ws, "B38", round(ref["late"], 6), PCT)

    ws["A40"] = "Nghiệm đối chiếu khi chi phí phạt mỗi đơn trễ (B15) = 20 USD"
    ws["A40"].font = BOLD
    head(ws, 41, ["Chỉ tiêu"] + MODE_ORDER)
    put(ws, "A42", "Tỷ trọng tối ưu")
    for j, col in enumerate("BCDE"):
        put(ws, f"{col}42", round(float(ref20["share"][j]), 6), PCT)
    put(ws, "A43", "Lợi nhuận ròng tối ưu (USD)")
    put(ws, "B43", round(ref20["obj"], 2), USD)
    put(ws, "A44", "Tỷ lệ trễ tại phương án tối ưu")
    put(ws, "B44", round(ref20["late"], 6), PCT)

    steps(ws, 46, [
        "B1: Biến quyết định = B4:E4 (khởi đầu bằng tỷ trọng hiện tại); biến mục tiêu = B20; ràng buộc ở bảng A28:D32",
        "B2: Công thức liên kết: B20 = N × SUMPRODUCT(tỷ trọng, lợi nhuận ròng/đơn); B21 = SUMPRODUCT(tỷ trọng, tỷ lệ trễ)",
        "B3: Công thức ràng buộc: B22 = SUM(B4:E4); B21 ≤ B16; B4:E4 nằm trong [B6:E6, B7:E7]",
        "B4: Bảng tính đã lập sẵn ở trên",
        "B5: Data > Solver: Set Objective B20 (Max); By Changing B4:E4; thêm 4 ràng buộc; chọn Simplex LP; tick 'Make Unconstrained Variables Non-Negative'; Solve",
    ])
    widths(ws, {"A": 46, "B": 18, "C": 16, "D": 16, "E": 18, "F": 12})
    solver_names(ws, "BT1_Solver!$B$20", "BT1_Solver!$B$4:$E$4", "max", [
        ("BT1_Solver!$B$4:$E$4", 3, "BT1_Solver!$B$6:$E$6"),
        ("BT1_Solver!$B$4:$E$4", 1, "BT1_Solver!$B$7:$E$7"),
        ("BT1_Solver!$B$22", 2, "1"),
        ("BT1_Solver!$B$21", 1, "BT1_Solver!$B$16"),
    ])
    return ws


# --------------------------------------------------------------------------
# 4. BT2 - chính sách chiết khấu (Data Table, Scenario, Goal Seek)
# --------------------------------------------------------------------------
def fit_discount(band):
    slope, intercept = np.polyfit(band["rate"], band["profit"], 1)
    return float(intercept), float(slope)


def sheet_bt2(wb, band, overall):
    a, b = fit_discount(band)
    n0, s0, d0 = overall["lines"], overall["gross"], overall["d0"]
    ln_ht = n0 * (a + b * d0)

    def profit(d, eta):
        return n0 * (1 + eta * (d - d0) * 100) * (a + b * d)

    ws = wb.create_sheet("BT2_WhatIf")
    title(ws, "Bài toán 2: Chính sách chiết khấu - lợi nhuận thay đổi thế nào khi đổi mức giảm giá")
    textbox(ws, "H1:N14",
            f"Bối cảnh: chiết khấu trung bình hiện ≈{vn(FACTS['d0'])}%. Dữ liệu cho thấy chiết khấu KHÔNG làm tăng số "
            f"lượng mua (tương quan {vn(FACTS['corr'], 3)}) nhưng làm giảm lợi nhuận/dòng. Câu hỏi: nếu đổi mức chiết "
            "khấu, lợi nhuận đổi bao nhiêu, và cầu phải nhạy tới mức nào thì tăng giảm giá mới đáng?\n\n"
            "Mô hình: lợi nhuận/dòng = a + b × d (hồi quy từ 6 dải chiết khấu, hàm INTERCEPT/SLOPE); "
            "số dòng đơn mới = N0 × (1 + η × (d − d0) × 100). η là GIẢ ĐỊNH (dữ liệu cho η ≈ 0, "
            "không ước lượng được độ co giãn thật).\n\nCông cụ: Data Table (1 và 2 biến), Scenario "
            "Manager, Goal Seek.")

    rows = [
        (3, "Số dòng đơn hiện tại (N0)", "=DuLieu!B45", INT, None, "N0"),
        (4, "Doanh thu gộp TB/dòng (USD)", "=DuLieu!B46", USD, None, "S0"),
        (5, "Chiết khấu hiện tại (d0)", "=DuLieu!B47", PCT, None, "d0"),
        (6, "Hệ số a: lợi nhuận/dòng tại d = 0 (USD)", "=INTERCEPT(DuLieu!D37:D42,DuLieu!B37:B42)", USD, None, "HeSoA"),
        (7, "Hệ số b: USD đổi trên mỗi +100% chiết khấu", "=SLOPE(DuLieu!D37:D42,DuLieu!B37:B42)", USD, None, "HeSoB"),
        (8, "Mức chiết khấu mới d (BIẾN QUYẾT ĐỊNH)", 0.15, PCT, YELLOW, "d_moi"),
        (9, "η: % sản lượng tăng mỗi +1 điểm % chiết khấu (tham số đặt)", 0.0, PCT, BLUE, "eta"),
        (11, "Hệ số sản lượng", "=1+eta*(d_moi-d0)*100", '0.0000', None, None),
        (12, "Số dòng đơn mới", "=N0*B11", INT, None, None),
        (13, "Lợi nhuận/dòng mới (USD)", "=HeSoA+HeSoB*d_moi", USD, None, None),
        (14, "Doanh thu thuần mới (USD)", "=B12*S0*(1-d_moi)", INT, None, None),
        (15, "Tổng lợi nhuận mới (BIẾN KẾT QUẢ)", "=B12*B13", INT, YELLOW, "LoiNhuan"),
        (16, "Lợi nhuận hiện tại (mô hình tại d0, η = 0)", "=N0*(HeSoA+HeSoB*d0)", INT, None, "LN_HT"),
        (17, "Chênh lệch lợi nhuận so với hiện tại", "=LoiNhuan-LN_HT", INT, None, "ChenhLech"),
    ]
    for r, label, val, fmt, fill, name in rows:
        put(ws, f"A{r}", label, bold=(fill is YELLOW))
        put(ws, f"B{r}", val, fmt, fill, fill is YELLOW)
        if name:
            wb.defined_names[name] = DefinedName(name, attr_text=f"BT2_WhatIf!$B${r}")
    put(ws, "A18", "Ràng buộc: 0% ≤ d ≤ 25% (khoảng chiết khấu có trong dữ liệu)")

    # a) Data Table 1 biến: cột đầu A22:A27 = d, hàng 21 chứa công thức đầu ra
    d_vals = [0.0, 0.05, 0.10, 0.15, 0.20, 0.25]
    ws["A20"] = "a) Data Table 1 biến: lợi nhuận theo mức chiết khấu (η lấy từ B9). Column input cell = B8"
    ws["A20"].font = BOLD
    put(ws, "A21", "d  \\  kết quả", bold=True)
    put(ws, "B21", "=LoiNhuan", INT, BLUE, True)
    put(ws, "C21", "=ChenhLech", INT, BLUE, True)
    for i, d in enumerate(d_vals):
        r = 22 + i
        put(ws, f"A{r}", d, PCT, bold=True)
        put(ws, f"B{r}", round(profit(d, 0.0), 2), INT)
        put(ws, f"C{r}", round(profit(d, 0.0) - ln_ht, 2), INT)
    ws["B22"] = DataTableFormula(ref="B22:C27", dt2D=False, dtr=False, r1="B8")

    # b) Data Table 2 biến: d (cột) x η (hàng)
    etas = [0.0, 0.005, 0.01, 0.02, 0.03]
    ws["A30"] = "b) Data Table 2 biến: lợi nhuận theo d (cột) và η (hàng). Row input = B9, Column input = B8"
    ws["A30"].font = BOLD
    put(ws, "A31", "=LoiNhuan", INT, BLUE, True)
    for j, e in enumerate(etas):
        put(ws, f"{get_column_letter(2 + j)}31", e, PCT, bold=True)
    for i, d in enumerate(d_vals):
        r = 32 + i
        put(ws, f"A{r}", d, PCT, bold=True)
        for j, e in enumerate(etas):
            put(ws, f"{get_column_letter(2 + j)}{r}", round(profit(d, e), 2), INT)
    ws["B32"] = DataTableFormula(ref="B32:F37", dt2D=True, dtr=True, r1="B9", r2="B8")

    # c) Bảng so sánh kịch bản (công thức) + Scenario Manager thật
    ws["A40"] = "c) So sánh kịch bản (công thức) - cùng 5 kịch bản đã nạp vào Scenario Manager (Data > What-If > Scenario Manager)"
    ws["A40"].font = BOLD
    head(ws, 41, ["Kịch bản", "d", "η", "Số dòng đơn", "Lợi nhuận (USD)", "Chênh lệch vs hiện tại"])
    scen = [("KB1_Ngung_giam_gia", 0.0, 0.0), ("KB2_Giam_gia_thap", 0.05, 0.0),
            ("KB3_Giu_hien_tai", None, 0.0), ("KB4_Giam_cao_cau_nhay", 0.20, 0.02),
            ("KB5_Giam_cao_cau_rat_nhay", 0.25, 0.03)]
    for i, (nm, d, e) in enumerate(scen):
        r = 42 + i
        put(ws, f"A{r}", nm)
        put(ws, f"B{r}", d if d is not None else "=d0", PCT)
        put(ws, f"C{r}", e, PCT)
        put(ws, f"D{r}", f"=N0*(1+C{r}*(B{r}-d0)*100)", INT)
        put(ws, f"E{r}", f"=D{r}*(HeSoA+HeSoB*B{r})", INT)
        put(ws, f"F{r}", f"=E{r}-LN_HT", INT)
    ws.scenarios = ScenarioList(
        # val ghi dạng "0.05" (không "0.0"), kèm numFmtId=10 (định dạng %) như Excel tự ghi,
        # và bỏ thuộc tính deleted/undone để Excel đọc là số chứ không phải chuỗi
        scenario=[Scenario(name=nm, locked=True, user="DSS", count=2,
                           inputCells=[InputCells(r="B8", deleted=None, undone=None, numFmtId=10,
                                                  val="%g" % (d if d is not None else round(d0, 6))),
                                       InputCells(r="B9", deleted=None, undone=None, numFmtId=10,
                                                  val="%g" % e)])
                  for nm, d, e in scen],
        current=3, show=3, sqref="B15")

    # d) Goal Seek tách riêng để không làm đổi mô hình chính (B8, B9)
    ws["A49"] = "d) Goal Seek: η tối thiểu để mức chiết khấu mới KHÔNG làm giảm lợi nhuận"
    ws["A49"].font = BOLD
    put(ws, "A50", "Mức chiết khấu mới cần xét (d)")
    put(ws, "B50", 0.15, PCT, BLUE)
    put(ws, "A51", "η cần tìm (BIẾN THAY ĐỔI cho Goal Seek)", bold=True)
    put(ws, "B51", 0.0, '0.0000%', YELLOW, True)
    put(ws, "A52", "Chênh lệch lợi nhuận vs hiện tại (đặt = 0)", bold=True)
    put(ws, "B52", "=N0*(1+B51*(B50-d0)*100)*(HeSoA+HeSoB*B50)-LN_HT", INT, YELLOW, True)
    put(ws, "A53", "Đáp án giải tích η* (để đối chiếu kết quả Goal Seek)")
    put(ws, "B53", "=IFERROR(((HeSoA+HeSoB*d0)/(HeSoA+HeSoB*B50)-1)/((B50-d0)*100),\"d = d0\")", '0.0000%')
    steps(ws, 55, [
        "B1: Data Table a: nhập công thức đầu ra ở hàng 21 (đã có), chọn vùng A21:C27, Data > What-If Analysis > Data Table, Column input cell = B8",
        "B2: Data Table b: chọn vùng A31:F37, Row input cell = B9, Column input cell = B8",
        "B3: Scenario Manager: Add 5 kịch bản (bảng A41:C46), Changing cells = B8:B9, rồi Summary với Result cells = B15,B17",
        "B4: Goal Seek: Data > What-If > Goal Seek: Set cell = B52, To value = 0, By changing cell = B51",
        "B5: Đối chiếu B51 với B53 - hai số trùng nhau (sai số dưới 0,0001%)",
    ])
    widths(ws, {"A": 56, "B": 18, "C": 18, "D": 16, "E": 18, "F": 22})
    return ws, dict(a=a, b=b, ln_ht=ln_ht, profit=profit, d0=d0, n0=n0)


# --------------------------------------------------------------------------
# 5. BT3 - điểm hoà vốn theo ngành hàng (Goal Seek)
# --------------------------------------------------------------------------
def sheet_bt3(wb, dept):
    ws = wb.create_sheet("BT3_GoalSeek")
    title(ws, "Bài toán 3: Điểm hoà vốn theo ngành hàng - sản lượng tối thiểu để không lỗ")
    textbox(ws, "H1:N14",
            f"Bối cảnh: biên lợi nhuận gộp của các ngành chênh nhau (Book Shop ≈{vn(FACTS['margin_book'])}%, "
            f"Fan Shop ≈{vn(FACTS['margin_fan'])}%). "
            "Mỗi ngành phải gánh chi phí cố định (kho, nhân sự, marketing). Câu hỏi: ngành nào đang "
            "bán sát điểm hoà vốn, tức dễ lỗ khi sản lượng giảm?\n\nMô hình: Lợi nhuận = p × Q − v × Q − F. "
            "p (giá thuần/đơn vị) và v (biến phí/đơn vị = p − lợi nhuận/đơn vị) lấy từ dữ liệu. "
            "F = k × doanh thu thuần của ngành, k là GIẢ ĐỊNH (dữ liệu không có chi phí cố định).\n\n"
            "Công cụ: Goal Seek (tìm Q để lợi nhuận = 0), Data Table (lợi nhuận theo Q).")
    first = dept.iloc[0]["department_name"]
    q0 = int(dept.iloc[0]["units"])
    rng = lambda col: f"DuLieu!${col}$22:${col}$32"  # noqa: E731
    match = f"MATCH($B$3,DuLieu!$A$22:$A$32,0)"

    rows = [
        (3, "Ngành hàng đang xét (chọn trong danh sách)", first, None, BLUE),
        (4, "Giá thuần/đơn vị (p, USD)", f"=INDEX({rng('H')},{match})", USD, None),
        (5, "Biến phí/đơn vị (v, USD)", f"=INDEX({rng('J')},{match})", USD, None),
        (6, "Số lượng bán hiện tại (Q0)", f"=INDEX({rng('C')},{match})", INT, None),
        (7, "Tỷ lệ chi phí cố định / doanh thu thuần (k) (tham số đặt)", FIXED_COST_RATIO, PCT, BLUE),
        (8, "Chi phí cố định kỳ (F, USD)", f"=B7*INDEX({rng('F')},{match})", INT, None),
        (9, "Sản lượng Q (BIẾN QUYẾT ĐỊNH)", q0, INT, YELLOW),
        (10, "Doanh thu = p × Q", "=B4*B9", INT, None),
        (11, "Tổng biến phí = v × Q", "=B5*B9", INT, None),
        (12, "Lợi nhuận (BIẾN KẾT QUẢ) = doanh thu − biến phí − F", "=B10-B11-B8", INT, YELLOW),
        (13, "Sản lượng hoà vốn (giải tích) = F / (p − v)", "=B8/(B4-B5)", INT, None),
        (14, "Biên an toàn = (Q0 − hoà vốn) / Q0", "=(B6-B13)/B6", PCT, None),
    ]
    for r, label, val, fmt, fill in rows:
        put(ws, f"A{r}", label, bold=(fill is YELLOW))
        put(ws, f"B{r}", val, fmt, fill, fill is YELLOW)
    put(ws, "A15", "Ràng buộc: 0 ≤ Q ≤ 1,5 × Q0 (năng lực bán tối đa - giả định)")
    dv = DataValidation(type="list", formula1="=DuLieu!$A$22:$A$32", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")

    ws["A17"] = "Data Table 1 biến: lợi nhuận theo sản lượng (Column input cell = B9)"
    ws["A17"].font = BOLD
    put(ws, "A18", "Q  \\  lợi nhuận", bold=True)
    put(ws, "B18", "=B12", INT, BLUE, True)
    pct_q = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5]
    p0 = float(dept.iloc[0]["net"] / dept.iloc[0]["units"])
    m0 = float(dept.iloc[0]["profit"] / dept.iloc[0]["units"])
    f0 = FIXED_COST_RATIO * float(dept.iloc[0]["net"])
    for i, pq in enumerate(pct_q):
        r = 19 + i
        put(ws, f"A{r}", f"=B6*{pq}", INT, bold=True)
        put(ws, f"B{r}", round(m0 * q0 * pq - f0, 2), INT)
    ws["B19"] = DataTableFormula(ref="B19:B24", dt2D=False, dtr=False, r1="B9")

    ws["A27"] = "Điểm hoà vốn của cả 11 ngành (công thức giải tích - dùng để đối chiếu Goal Seek)"
    ws["A27"].font = BOLD
    head(ws, 28, ["Ngành hàng", "p (USD)", "v (USD)", "Q0", "F (USD)", "Hoà vốn (đơn vị)", "Biên an toàn", "Đánh giá"])
    for i in range(11):
        r = 29 + i
        dr = 22 + i
        put(ws, f"A{r}", f"=DuLieu!A{dr}")
        put(ws, f"B{r}", f"=DuLieu!H{dr}", USD)
        put(ws, f"C{r}", f"=DuLieu!J{dr}", USD)
        put(ws, f"D{r}", f"=DuLieu!C{dr}", INT)
        put(ws, f"E{r}", f"=$B$7*DuLieu!F{dr}", INT)
        put(ws, f"F{r}", f"=E{r}/(B{r}-C{r})", INT)
        put(ws, f"G{r}", f"=(D{r}-F{r})/D{r}", PCT)
        put(ws, f"H{r}", f'=IF(G{r}<0.2,"Rủi ro",IF(G{r}<0.4,"Theo dõi","An toàn"))')
    ws["A40"] = "Ngưỡng đánh giá (biên an toàn < 20% = Rủi ro, < 40% = Theo dõi) là quy ước của mô hình, không phải chuẩn ngành."

    # Data Table 2 biến: độ nhạy biên an toàn theo k (cột) và ngành (hàng)
    ws["A50"] = "Độ nhạy: biên an toàn theo tỷ lệ chi phí cố định k (cột) và ngành hàng (hàng). Row input = B3, Column input = B7"
    ws["A50"].font = BOLD
    put(ws, "A51", "=B14", PCT, BLUE, True)
    for j in range(11):
        put(ws, f"{get_column_letter(2 + j)}51", dept.iloc[j]["department_name"], bold=True)
    k_vals = [0.03, 0.05, 0.07, 0.09, 0.11]
    for i, k in enumerate(k_vals):
        r = 52 + i
        put(ws, f"A{r}", k, PCT, bold=True)
        for j in range(11):
            d = dept.iloc[j]
            be = k * float(d["net"]) / (float(d["profit"]) / float(d["units"]))
            put(ws, f"{get_column_letter(2 + j)}{r}", round((float(d["units"]) - be) / float(d["units"]), 6), PCT)
    ws["B52"] = DataTableFormula(ref="B52:L56", dt2D=True, dtr=True, r1="B3", r2="B7")

    steps(ws, 42, [
        "B1: Biến quyết định = Q (B9); biến kết quả = lợi nhuận (B12); ràng buộc 0 ≤ Q ≤ 1,5 × Q0",
        "B2: Chọn ngành ở B3 (danh sách thả xuống); gõ lại B9 = giá trị Q0 ở B6 để bắt đầu từ hiện trạng",
        "B3: Data > What-If Analysis > Goal Seek: Set cell = B12, To value = 0, By changing cell = B9",
        "B4: Kết quả B9 trùng với B13 (hoà vốn giải tích)",
        "B5: Đối chiếu bảng A28:H39: ngành nào biên an toàn thấp nhất thì dễ lỗ nhất khi sản lượng giảm",
        "B6: Độ nhạy: chọn vùng A51:L56, Data > What-If > Data Table, Row input cell = B3, Column input cell = B7; bảng cho thấy thứ tự rủi ro các ngành có đổi khi k đổi không",
    ])
    widths(ws, {"A": 56, "B": 16, "C": 12, "D": 12, "E": 14, "F": 18, "G": 14, "H": 12})
    return ws


# --------------------------------------------------------------------------
# 6. BT4 - phân bổ đơn giữa 5 thị trường (Solver, LP)
# --------------------------------------------------------------------------
def solve_bt4(market):
    n = market["n"].to_numpy(dtype=float)
    pi = market["profit"].to_numpy()
    res = linprog(-pi, A_eq=[np.ones(len(n))], b_eq=[n.sum()],
                  bounds=list(zip((1 - DEMAND_BAND) * n, (1 + DEMAND_BAND) * n)), method="highs")
    assert res.success, f"BT4 vô nghiệm: {res.message}"
    return dict(x=res.x, obj=float(-res.fun), obj0=float(pi @ n))


def sheet_bt4(wb, market, ref):
    ws = wb.create_sheet("BT4_Solver")
    title(ws, "Bài toán 4: Phân bổ năng lực xử lý đơn giữa 5 thị trường - tối đa lợi nhuận")
    textbox(ws, "K1:Q12",
            f"Bối cảnh: lợi nhuận/đơn của Pacific Asia chỉ ${vn(FACTS['pac_profit'])}, trong khi 4 thị trường còn lại "
            f"${vn(FACTS['oth_profit_min'], 0)}-{vn(FACTS['oth_profit_max'], 0)}; doanh thu/đơn của Pacific Asia cũng thấp hơn "
            f"(${vn(FACTS['pac_sales'], 0)} so với ${vn(FACTS['oth_sales_min'], 0)}-{vn(FACTS['oth_sales_max'], 0)}). Câu hỏi: với cùng "
            "tổng số đơn, nên phân bổ nguồn lực (kho, vận chuyển, marketing) lệch về thị trường nào?\n\n"
            "Giả định: mỗi thị trường tăng/giảm tối đa ±20% so với hiện tại (nhu cầu và năng lực có "
            "giới hạn); lợi nhuận/đơn của thị trường giữ nguyên khi số đơn đổi - đây là giả định mạnh "
            "vì lợi nhuận/đơn phụ thuộc cơ cấu sản phẩm.")
    head(ws, 3, ["Thị trường", "Đơn hiện tại", "Tỷ trọng", "Lợi nhuận/đơn (USD)", "Tối thiểu",
                 "Tối đa", "Số đơn phân bổ (BIẾN QUYẾT ĐỊNH)", "Lợi nhuận (USD)", "Tỷ lệ trễ"])
    for i, r in market.iterrows():
        row = 4 + i
        dr = 13 + i
        put(ws, f"A{row}", f"=DuLieu!A{dr}")
        put(ws, f"B{row}", f"=DuLieu!B{dr}", INT)
        put(ws, f"C{row}", f"=DuLieu!C{dr}", PCT)
        put(ws, f"D{row}", f"=DuLieu!E{dr}", USD)
        put(ws, f"E{row}", f"=(1-$B$13)*B{row}", INT)
        put(ws, f"F{row}", f"=(1+$B$13)*B{row}", INT)
        put(ws, f"G{row}", int(r["n"]), INT, YELLOW, True)
        put(ws, f"H{row}", f"=D{row}*G{row}", INT)
        put(ws, f"I{row}", f"=DuLieu!F{dr}", PCT)
    put(ws, "A9", "Tổng", bold=True)
    put(ws, "B9", "=SUM(B4:B8)", INT, bold=True)
    put(ws, "C9", "=SUM(C4:C8)", PCT, bold=True)
    put(ws, "G9", "=SUM(G4:G8)", INT, bold=True)
    put(ws, "H9", "=SUM(H4:H8)", INT, YELLOW, True)
    put(ws, "I9", "=SUMPRODUCT(I4:I8,G4:G8)/G9", PCT, bold=True)
    ws["A11"] = "Tham số do người ra quyết định đặt"
    ws["A11"].font = BOLD
    put(ws, "A12", "Tổng số đơn phải phân bổ")
    put(ws, "B12", "=B9", INT)
    put(ws, "A13", "Biên nhu cầu ± mỗi thị trường (tham số đặt)")
    put(ws, "B13", DEMAND_BAND, PCT, BLUE)

    ws["A15"] = "Kết quả"
    ws["A15"].font = BOLD
    put(ws, "A16", "Lợi nhuận hiện tại (USD)")
    put(ws, "B16", "=SUMPRODUCT(D4:D8,B4:B8)", INT)
    put(ws, "A17", "Lợi nhuận sau phân bổ - BIẾN MỤC TIÊU (H9)", bold=True)
    put(ws, "B17", "=H9", INT, YELLOW, True)
    put(ws, "A18", "Chênh lệch (USD)")
    put(ws, "B18", "=B17-B16", INT)
    put(ws, "A19", "Cải thiện (%)")
    put(ws, "B19", "=B18/B16", PCT)

    ws["A21"] = "Ràng buộc (nhập vào hộp thoại Solver)"
    ws["A21"].font = BOLD
    head(ws, 22, ["Ràng buộc", "Ô", "Dấu", "Giá trị giới hạn"])
    cons = [("Số đơn mỗi thị trường ≥ tối thiểu", "G4:G8", ">=", "E4:E8"),
            ("Số đơn mỗi thị trường ≤ tối đa", "G4:G8", "<=", "F4:F8"),
            ("Tổng số đơn giữ nguyên", "G9", "=", "B12")]
    for i, r in enumerate(cons):
        for j, v in enumerate(r):
            put(ws, f"{'ABCD'[j]}{23 + i}", v)

    ws["A28"] = "Nghiệm tối ưu đối chiếu (tính độc lập bằng quy hoạch tuyến tính)"
    ws["A28"].font = BOLD
    head(ws, 29, ["Thị trường", "Số đơn tối ưu"])
    for i, r in market.iterrows():
        put(ws, f"A{30 + i}", r["market"])
        put(ws, f"B{30 + i}", round(float(ref["x"][i]), 1), INT)
    put(ws, "A35", "Lợi nhuận tối ưu (USD)", bold=True)
    put(ws, "B35", round(ref["obj"], 2), INT, bold=True)

    steps(ws, 37, [
        "B1: Biến quyết định = G4:G8; biến mục tiêu = H9; ràng buộc ở bảng A22:D25",
        "B2: Công thức liên kết: H = lợi nhuận/đơn × số đơn; H9 = SUM(H4:H8)",
        "B3: Công thức ràng buộc: G9 = SUM(G4:G8) = B12; E ≤ G ≤ F với E = (1 − B13) × hiện tại, F = (1 + B13) × hiện tại",
        "B4: Bảng tính đã lập sẵn ở trên",
        "B5: Data > Solver: Set Objective H9 (Max); By Changing G4:G8; thêm 3 ràng buộc; Simplex LP; tick Non-Negative; Solve",
    ])
    widths(ws, {"A": 46, "B": 16, "C": 12, "D": 18, "E": 14, "F": 14, "G": 24, "H": 16, "I": 12})
    solver_names(ws, "BT4_Solver!$H$9", "BT4_Solver!$G$4:$G$8", "max", [
        ("BT4_Solver!$G$4:$G$8", 3, "BT4_Solver!$E$4:$E$8"),
        ("BT4_Solver!$G$4:$G$8", 1, "BT4_Solver!$F$4:$F$8"),
        ("BT4_Solver!$G$9", 2, "BT4_Solver!$B$12"),
    ])
    return ws


# --------------------------------------------------------------------------
# 7. Tổng quan
# --------------------------------------------------------------------------
def sheet_overview(wb):
    ws = wb.active
    ws.title = "TongQuan"
    title(ws, "Phần B - Model-Driven DSS: 4 bài toán ra quyết định trên dữ liệu DataCo")
    ws["A2"] = "Chú thích màu: ô nền VÀNG = biến quyết định / biến mục tiêu; ô nền XANH NHẠT = tham số do người ra quyết định đặt hoặc tiêu đề."
    head(ws, 4, ["Bài toán", "Sheet", "Biến quyết định", "Biến mục tiêu", "Ràng buộc chính", "Công cụ Excel"])
    data = [
        ("1. Cơ cấu phương thức vận chuyển", "BT1_Solver", "Tỷ trọng đơn của 4 phương thức",
         "Tổng lợi nhuận (trừ phạt đơn trễ nếu có)", "Tỷ lệ trễ chung ≤ 45%; mỗi tỷ trọng lệch ≤ ±15 điểm %; tổng = 100%", "Solver (Simplex LP)"),
        ("2. Chính sách chiết khấu", "BT2_WhatIf", "Mức chiết khấu d; độ co giãn η",
         "Tổng lợi nhuận", "0% ≤ d ≤ 25%", "Data Table, Scenario Manager, Goal Seek"),
        ("3. Điểm hoà vốn theo ngành hàng", "BT3_GoalSeek", "Sản lượng Q của ngành",
         "Lợi nhuận ngành = 0", "0 ≤ Q ≤ 1,5 × Q0", "Goal Seek, Data Table"),
        ("4. Phân bổ đơn giữa 5 thị trường", "BT4_Solver", "Số đơn mỗi thị trường",
         "Tổng lợi nhuận", "Tổng đơn giữ nguyên; mỗi thị trường ±20%", "Solver (Simplex LP)"),
    ]
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            c = put(ws, f"{get_column_letter(1 + j)}{5 + i}", v)
            c.alignment = WRAP
    ws["A10"] = "Tham số do người ra quyết định đặt (không lấy từ dữ liệu)"
    ws["A10"].font = BOLD
    head(ws, 11, ["Mã", "Tham số", "Giá trị", "Ô", "Lý do chọn", "Điều kiện xem lại"])
    assumptions = [
        ("G1", "Chi phí phạt mỗi đơn trễ (USD)", PENALTY_PER_LATE, "BT1_Solver!B15",
         "Mặc định 0 để mô hình chỉ dùng lợi nhuận lấy từ dữ liệu; nhập số dương (ví dụ 20) để thử độ nhạy",
         "Có số chi phí chăm sóc khách hàng thật"),
        ("G2", "Ngưỡng tỷ lệ trễ chung", LATE_CAP, "BT1_Solver!B16",
         f"Mục tiêu quản trị do ban điều hành đặt, thấp hơn hiện tại {vn(FACTS['late'])}% nhưng khả thi trong biên ±15 điểm %",
         "Ban điều hành đặt SLA khác"),
        ("G3", "Biên dịch chuyển tỷ trọng phương thức", SHIFT_BAND, "BT1_Solver!B17",
         "Khách chọn phương thức theo nhu cầu; ±15 điểm % đủ rộng để ngưỡng 45% có nghiệm",
         "Có dữ liệu khách đổi phương thức"),
        ("G4", "η: % sản lượng tăng mỗi +1 điểm % chiết khấu", "0% (cơ sở) đến 3%", "BT2_WhatIf!B9",
         f"Dữ liệu cho η ≈ 0 (tương quan {vn(FACTS['corr'], 3)}) nên mọi η > 0 là giả định để thử độ nhạy",
         "Có thí nghiệm giảm giá thật (A/B test)"),
        ("G5", "Chi phí cố định / doanh thu thuần ngành", FIXED_COST_RATIO, "BT3_GoalSeek!B7",
         f"Dữ liệu không có chi phí cố định; 7% thấp hơn biên lợi nhuận {vn(FACTS['margin_all'])}% nên ngành trung bình vẫn có lãi",
         "Có báo cáo chi phí cố định theo ngành"),
        ("G6", "Biên nhu cầu mỗi thị trường", DEMAND_BAND, "BT4_Solver!B13",
         "Giới hạn năng lực và nhu cầu; ±20% là mức dịch chuyển nguồn lực trong một kỳ kế hoạch",
         "Có số năng lực kho theo thị trường"),
    ]
    for i, row in enumerate(assumptions):
        for j, v in enumerate(row):
            c = put(ws, f"{get_column_letter(1 + j)}{12 + i}", v, PCT if (j == 2 and isinstance(v, float) and v < 1) else None)
            c.alignment = WRAP
    ws["C12"].number_format = USD
    ws["A19"] = ("Lưu ý: số liệu đã loại đơn CANCELED và SUSPECTED_FRAUD (khớp Phần A); "
                 "dữ liệu giao hàng có dạng mô phỏng (số ngày thực tế gần như cố định theo phương thức) nên "
                 "kết luận về phương thức vận chuyển chỉ đúng 'trong dữ liệu này'.")
    ws["A19"].alignment = WRAP
    ws.merge_cells("A19:F19")
    ws.row_dimensions[19].height = 48
    widths(ws, {"A": 34, "B": 28, "C": 30, "D": 32, "E": 60, "F": 40})
    return ws


def main():
    mode, market, dept, band, overall = load()
    ref1 = solve_bt1(mode)
    ref20 = solve_bt1(mode, 20.0)
    ref4 = solve_bt4(market)

    wb = Workbook()
    sheet_overview(wb)
    sheet_data(wb, mode, market, dept, band, overall)
    sheet_bt1(wb, mode, ref1, ref20)
    _, ref2 = sheet_bt2(wb, band, overall)
    sheet_bt3(wb, dept)
    sheet_bt4(wb, market, ref4)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)

    print("BT1 share0 ", np.round(ref1["share0"], 4), "late0", round(ref1["late0"], 4))
    print("BT1 share* ", np.round(ref1["share"], 4), "late*", round(ref1["late"], 4))
    print("BT1 obj0 %.0f  obj* %.0f  gain %.0f" % (ref1["obj0"], ref1["obj"], ref1["obj"] - ref1["obj0"]))
    print("BT2 a=%.4f b=%.4f d0=%.4f LN_HT=%.0f" % (ref2["a"], ref2["b"], ref2["d0"], ref2["ln_ht"]))
    for d in (0, .05, .10, .15, .20, .25):
        print("   d=%.2f profit(eta=0)=%.0f" % (d, ref2["profit"](d, 0.0)))
    d1 = 0.15
    eta = ((ref2["a"] + ref2["b"] * ref2["d0"]) / (ref2["a"] + ref2["b"] * d1) - 1) / ((d1 - ref2["d0"]) * 100)
    print("BT2 eta* at d=15%% = %.6f" % eta)
    print("BT4 x*", np.round(ref4["x"], 0), "obj0 %.0f obj* %.0f gain %.0f" % (ref4["obj0"], ref4["obj"], ref4["obj"] - ref4["obj0"]))
    print("Đã ghi", OUT)


if __name__ == "__main__":
    main()
