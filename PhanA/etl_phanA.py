"""
ETL Phần A - DSS Chuỗi cung ứng DataCo
Làm sạch dữ liệu thô -> dựng mô hình Star Schema -> xuất CSV cho Tableau.

Chạy:
    python3 etl_phanA.py
Đầu vào : ../dataset/dataco-supply-chain/DataCoSupplyChainDataset.csv
Đầu ra  : ./tableau_data/{fact_sales, dim_customer, dim_product, dim_date,
          dim_location, dim_delivery}.csv  +  kpi_summary.json
"""
import json
import os
import zipfile
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "..", "dataset", "dataco-supply-chain")
SRC = os.path.join(DATA_DIR, "DataCoSupplyChainDataset.csv")
SRC_ZIP = SRC + ".zip"          # dùng khi chưa giải nén CSV
OUT = os.path.join(HERE, "tableau_data")
os.makedirs(OUT, exist_ok=True)

# Trạng thái đơn bị LOẠI khỏi mô hình phân tích (quyết định nghiệp vụ, xem Mục 3.3 báo cáo):
#  - CANCELED        : đơn đã huỷ, không phát sinh doanh thu thực
#  - SUSPECTED_FRAUD : đơn nghi gian lận, không đáng tin để tính doanh thu/lợi nhuận
# Hai nhóm này cũng chính là toàn bộ dòng « Shipping canceled » (đơn không được giao),
# nếu giữ lại sẽ bị tính là « đúng hạn » và làm sai tỷ lệ giao trễ.
EXCLUDED_STATUS = ["CANCELED", "SUSPECTED_FRAUD"]

# ---------------------------------------------------------------------------
# 1. NẠP & LÀM SẠCH
# ---------------------------------------------------------------------------
if os.path.exists(SRC):
    df = pd.read_csv(SRC, encoding="latin-1")
else:
    with zipfile.ZipFile(SRC_ZIP) as z:
        df = pd.read_csv(z.open("DataCoSupplyChainDataset.csv"), encoding="latin-1")
raw_rows, raw_cols = df.shape

# 1.1 Bỏ cột rác / không mang thông tin phân tích
drop_cols = [
    "Product Description",   # 100% trống
    "Order Zipcode",         # 86% trống
    "Customer Email",        # 1 giá trị giả (XXXXXXXXX)
    "Customer Password",     # 1 giá trị giả
    "Product Status",        # toàn bộ = 0
    "Product Image",         # URL ảnh, không dùng phân tích
    "Benefit per order",     # trùng y hệt 'Order Profit Per Order'
]
df = df.drop(columns=[c for c in drop_cols if c in df.columns])

# 1.2 Bỏ dòng trùng hoàn toàn (nếu có) + kiểm tra khoá
dups = int(df.duplicated().sum())
df = df.drop_duplicates()

# 1.3 Chuẩn hoá chuỗi: trim khoảng trắng thừa ở các cột text
for c in df.select_dtypes(include="object").columns:
    df[c] = df[c].str.strip()

# 1.4 Xử lý giá trị trống
df["Customer Lname"] = df["Customer Lname"].fillna("Unknown")
df["Customer Zipcode"] = df["Customer Zipcode"].fillna(0)

# 1.5 Chuẩn hoá ngày tháng (M/D/YYYY H:M)
df["order_dt"] = pd.to_datetime(df["order date (DateOrders)"], format="%m/%d/%Y %H:%M")
df["ship_dt"] = pd.to_datetime(df["shipping date (DateOrders)"], format="%m/%d/%Y %H:%M")

# 1.6 Trường dẫn xuất: độ trễ giao hàng (ngày thực - ngày kế hoạch)
df["shipping_delay"] = df["Days for shipping (real)"] - df["Days for shipment (scheduled)"]

# 1.7 Loại đơn CANCELED / SUSPECTED_FRAUD (lọc theo Order Status; làm SAU các bước làm
# sạch nên các số liệu làm sạch ở trên vẫn là số liệu của dữ liệu thô)
excl_mask = df["Order Status"].isin(EXCLUDED_STATUS)
_by_order = df.groupby("Order Id")
before_filter = {     # KPI của dữ liệu thô (trước khi lọc) để báo cáo so sánh tác động
    "rows": int(len(df)), "orders": int(df["Order Id"].nunique()),
    "customers": int(df["Customer Id"].nunique()),
    "total_sales": round(float(df["Sales"].sum()), 2),
    "total_profit": round(float(df["Order Profit Per Order"].sum()), 2),
    "profit_margin_pct": round(float(df["Order Profit Per Order"].sum() / df["Sales"].sum() * 100), 2),
    "late_delivery_rate_order_pct": round(float(_by_order["Late_delivery_risk"].max().mean() * 100), 2),
    "aov": round(float(_by_order["Order Item Total"].sum().mean()), 2),
    "dim_delivery_rows": int(len(df[["Shipping Mode", "Delivery Status", "Type", "Order Status"]]
                                 .drop_duplicates())),
}
excluded = {
    "before": before_filter,
    "statuses": EXCLUDED_STATUS,
    "rows": int(excl_mask.sum()),
    "orders": int(df.loc[excl_mask, "Order Id"].nunique()),
    "rows_by_status": {s: int((df["Order Status"] == s).sum()) for s in EXCLUDED_STATUS},
    "shipping_canceled_rows": int((df["Delivery Status"] == "Shipping canceled").sum()),
    "sales": round(float(df.loc[excl_mask, "Sales"].sum()), 2),
    "profit": round(float(df.loc[excl_mask, "Order Profit Per Order"].sum()), 2),
}
# bảo đảm không đơn nào có trạng thái hỗn hợp (nếu có, lọc theo dòng sẽ làm đơn bị cụt)
_mixed = df.groupby("Order Id")["Order Status"].nunique().gt(1).sum()
assert _mixed == 0, f"{_mixed} đơn có nhiều hơn một Order Status"
df = df[~excl_mask].reset_index(drop=True)
assert not (df["Delivery Status"] == "Shipping canceled").any(), "còn dòng Shipping canceled"

# ---------------------------------------------------------------------------
# 2. DỰNG CÁC BẢNG DIMENSION
# ---------------------------------------------------------------------------
# 2.1 Dim_Customer (khoá tự nhiên: Customer Id)
dim_customer = (
    df[["Customer Id", "Customer Fname", "Customer Lname", "Customer Segment",
        "Customer City", "Customer State", "Customer Country", "Customer Zipcode"]]
    .drop_duplicates(subset="Customer Id")
    .rename(columns={
        "Customer Id": "customer_key", "Customer Fname": "first_name",
        "Customer Lname": "last_name", "Customer Segment": "segment",
        "Customer City": "city", "Customer State": "state",
        "Customer Country": "country", "Customer Zipcode": "zipcode"})
    .sort_values("customer_key").reset_index(drop=True)
)

# 2.2 Dim_Product (khoá tự nhiên: Product Card Id) - phân cấp Product>Category>Department
dim_product = (
    df[["Product Card Id", "Product Name", "Category Id", "Category Name",
        "Department Id", "Department Name", "Product Price"]]
    .drop_duplicates(subset="Product Card Id")
    .rename(columns={
        "Product Card Id": "product_key", "Product Name": "product_name",
        "Category Id": "category_id", "Category Name": "category_name",
        "Department Id": "department_id", "Department Name": "department_name",
        "Product Price": "product_price"})
    .sort_values("product_key").reset_index(drop=True)
)

# 2.3 Dim_Date (role-playing: dùng chung cho ngày đặt & ngày giao)
all_dates = pd.concat([df["order_dt"], df["ship_dt"]]).dt.normalize().drop_duplicates()
dim_date = pd.DataFrame({"full_date": sorted(all_dates)})
dim_date["date_key"] = dim_date["full_date"].dt.strftime("%Y%m%d").astype(int)
dim_date["year"] = dim_date["full_date"].dt.year
dim_date["quarter"] = "Q" + dim_date["full_date"].dt.quarter.astype(str)
dim_date["month"] = dim_date["full_date"].dt.month
dim_date["month_name"] = dim_date["full_date"].dt.strftime("%b")
dim_date["day"] = dim_date["full_date"].dt.day
dim_date["day_of_week"] = dim_date["full_date"].dt.dayofweek + 1
dim_date["weekday_name"] = dim_date["full_date"].dt.strftime("%a")
dim_date["is_weekend"] = dim_date["full_date"].dt.dayofweek >= 5
dim_date = dim_date[["date_key", "full_date", "year", "quarter", "month",
                     "month_name", "day", "day_of_week", "weekday_name", "is_weekend"]]

# 2.4 Dim_Location (nơi giao hàng: Market>Region>Country>State>City) - khoá thay thế
loc_cols = ["Market", "Order Region", "Order Country", "Order State", "Order City"]
dim_location = df[loc_cols].drop_duplicates().reset_index(drop=True)
dim_location.insert(0, "location_key", dim_location.index + 1)
dim_location = dim_location.rename(columns={
    "Market": "market", "Order Region": "region", "Order Country": "country",
    "Order State": "state", "Order City": "city"})

# 2.5 Dim_Delivery (chiều hỗn hợp: giao hàng + thanh toán + trạng thái đơn) - khoá thay thế
del_cols = ["Shipping Mode", "Delivery Status", "Type", "Order Status"]
dim_delivery = df[del_cols].drop_duplicates().reset_index(drop=True)
dim_delivery.insert(0, "delivery_key", dim_delivery.index + 1)
dim_delivery = dim_delivery.rename(columns={
    "Shipping Mode": "shipping_mode", "Delivery Status": "delivery_status",
    "Type": "payment_type", "Order Status": "order_status"})

# ---------------------------------------------------------------------------
# 3. DỰNG BẢNG FACT (gắn khoá ngoại về các dimension)
# ---------------------------------------------------------------------------
fact = df.copy()
fact["order_date_key"] = fact["order_dt"].dt.strftime("%Y%m%d").astype(int)
fact["ship_date_key"] = fact["ship_dt"].dt.strftime("%Y%m%d").astype(int)

# gắn location_key
fact = fact.merge(
    dim_location.rename(columns={
        "market": "Market", "region": "Order Region", "country": "Order Country",
        "state": "Order State", "city": "Order City"}),
    on=loc_cols, how="left")
# gắn delivery_key
fact = fact.merge(
    dim_delivery.rename(columns={
        "shipping_mode": "Shipping Mode", "delivery_status": "Delivery Status",
        "payment_type": "Type", "order_status": "Order Status"}),
    on=del_cols, how="left")

fact_sales = fact[[
    "Order Item Id", "Order Id",            # khoá/degenerate
    "Customer Id", "Product Card Id",       # FK -> customer, product
    "order_date_key", "ship_date_key",      # FK -> date (role-playing)
    "location_key", "delivery_key",         # FK -> location, delivery
    # --- measures ---
    "Order Item Quantity", "Order Item Product Price", "Order Item Discount",
    "Order Item Discount Rate", "Sales", "Order Item Total",
    "Order Profit Per Order", "Order Item Profit Ratio",
    "Days for shipping (real)", "Days for shipment (scheduled)",
    "shipping_delay", "Late_delivery_risk",
]].rename(columns={
    "Order Item Id": "order_item_id", "Order Id": "order_id",
    "Customer Id": "customer_key", "Product Card Id": "product_key",
    "Order Item Quantity": "quantity", "Order Item Product Price": "unit_price",
    "Order Item Discount": "discount", "Order Item Discount Rate": "discount_rate",
    "Sales": "sales", "Order Item Total": "net_sales",
    "Order Profit Per Order": "profit", "Order Item Profit Ratio": "profit_ratio",
    "Days for shipping (real)": "days_ship_real",
    "Days for shipment (scheduled)": "days_ship_scheduled",
    "Late_delivery_risk": "late_delivery_risk",
})

# ---------------------------------------------------------------------------
# 3b. DỰNG BẢNG FACT THỨ HAI: FACT_ORDER_DELIVERY (grain = 1 đơn hàng / lần vận chuyển)
#     Lý do: các chỉ số giao hàng (ngày giao, trễ hẹn) ở cấp ĐƠN, lặp lại trên mọi
#     mặt hàng cùng đơn -> nếu để ở grain mặt hàng sẽ bị đếm trùng. Tách thành fact
#     riêng dạng accumulating snapshot (vòng đời: đặt hàng -> giao hàng).
# ---------------------------------------------------------------------------
grp = fact_sales.groupby("order_id")
fact_order_delivery = grp.agg(
    customer_key=("customer_key", "first"),
    order_date_key=("order_date_key", "first"),
    ship_date_key=("ship_date_key", "first"),
    location_key=("location_key", "first"),
    delivery_key=("delivery_key", "first"),
    days_ship_scheduled=("days_ship_scheduled", "first"),
    days_ship_real=("days_ship_real", "first"),
    shipping_delay=("shipping_delay", "first"),
    late_delivery_flag=("late_delivery_risk", "max"),
    order_item_count=("order_item_id", "count"),
    order_quantity=("quantity", "sum"),
    order_sales=("sales", "sum"),
    order_net_sales=("net_sales", "sum"),
    order_profit=("profit", "sum"),
).reset_index()
fact_order_delivery["on_time_flag"] = 1 - fact_order_delivery["late_delivery_flag"]
# đưa cột khoá lên đầu cho dễ đọc
fact_order_delivery = fact_order_delivery[[
    "order_id", "customer_key", "order_date_key", "ship_date_key",
    "location_key", "delivery_key",
    "days_ship_scheduled", "days_ship_real", "shipping_delay",
    "late_delivery_flag", "on_time_flag",
    "order_item_count", "order_quantity", "order_sales", "order_net_sales", "order_profit",
]]

# ---------------------------------------------------------------------------
# 4. XUẤT CSV
# ---------------------------------------------------------------------------
tables = {
    "fact_sales": fact_sales, "fact_order_delivery": fact_order_delivery,
    "dim_customer": dim_customer, "dim_product": dim_product,
    "dim_date": dim_date, "dim_location": dim_location, "dim_delivery": dim_delivery,
}
for name, t in tables.items():
    t.to_csv(os.path.join(OUT, f"{name}.csv"), index=False)

# ---------------------------------------------------------------------------
# 5. TÍNH KPI & INSIGHT (ghi ra JSON để đưa vào báo cáo)
# ---------------------------------------------------------------------------
f = fact_sales
n_orders = int(f["order_id"].nunique())
summary = {
    "raw_rows": int(raw_rows), "raw_cols": int(raw_cols),
    "clean_rows": int(len(f)), "dup_removed": dups,
    "excluded": excluded,
    "date_min": str(df["order_dt"].min().date()), "date_max": str(df["order_dt"].max().date()),
    "n_orders": n_orders, "n_customers": int(dim_customer.shape[0]),
    "n_products": int(dim_product.shape[0]),
    "tables": {k: list(v.shape) for k, v in tables.items()},
    # KPI
    "total_sales": round(float(f["sales"].sum()), 2),
    "net_sales": round(float(f["net_sales"].sum()), 2),
    "total_profit": round(float(f["profit"].sum()), 2),
    "profit_margin_pct": round(float(f["profit"].sum() / f["sales"].sum() * 100), 2),
    "late_delivery_rate_pct": round(float(f["late_delivery_risk"].mean() * 100), 2),
    "avg_shipping_delay": round(float(f["shipping_delay"].mean()), 2),
    "loss_order_lines": int((f["profit"] < 0).sum()),
    "loss_order_line_pct": round(float((f["profit"] < 0).mean() * 100), 2),
    "aov": round(float(f.groupby("order_id")["net_sales"].sum().mean()), 2),
    "total_quantity": int(f["quantity"].sum()),
    "avg_discount_rate_pct": round(float(f["discount_rate"].mean() * 100), 2),
    # KPI tính ở grain ĐƠN HÀNG (từ fact_order_delivery) - chính xác hơn cho giao hàng
    "late_delivery_rate_order_pct": round(float(fact_order_delivery["late_delivery_flag"].mean() * 100), 2),
    "on_time_rate_order_pct": round(float(fact_order_delivery["on_time_flag"].mean() * 100), 2),
    "avg_shipping_delay_order": round(float(fact_order_delivery["shipping_delay"].mean()), 2),
    "avg_items_per_order": round(float(fact_order_delivery["order_item_count"].mean()), 2),
}
# breakdowns cho insight
pm = f.merge(dim_product[["product_key", "department_name", "category_name"]], on="product_key")
summary["sales_by_department"] = (pm.groupby("department_name")["sales"].sum()
                                  .sort_values(ascending=False).round(0).head(6).to_dict())
lm = f.merge(dim_location[["location_key", "market"]], on="location_key")
summary["sales_by_market"] = (lm.groupby("market")["sales"].sum()
                              .sort_values(ascending=False).round(0).to_dict())
dm = f.merge(dim_delivery[["delivery_key", "shipping_mode"]], on="delivery_key")
summary["late_rate_by_shipmode_pct"] = (dm.groupby("shipping_mode")["late_delivery_risk"]
                                        .mean().mul(100).round(1).sort_values(ascending=False).to_dict())
summary["profit_by_market"] = (lm.groupby("market")["profit"].sum()
                               .sort_values(ascending=False).round(0).to_dict())

with open(os.path.join(OUT, "kpi_summary.json"), "w") as fp:
    json.dump(summary, fp, indent=2, ensure_ascii=False)

print(json.dumps(summary, indent=2, ensure_ascii=False))
print("\nĐã xuất các bảng:", {k: v.shape for k, v in tables.items()})
