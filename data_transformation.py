"""
data_transformation.py
APL Logistics – Customer, Product & Profitability Intelligence
Transforms raw APL_Logistics.csv → data/APL_Logistics_Transformed.csv

Run from project root:  python data_transformation.py
Verified against raw data: 180,519 records | $36,784,734 revenue | $3,966,903 profit
"""

import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ─── Column Rename Map ────────────────────────────────────────────────────────
RENAME_MAP = {
    "Type":                        "payment_type",
    "Days for shipping (real)":    "shipping_days_actual",
    "Days for shipment (scheduled)":"shipping_days_scheduled",
    "Benefit per order":           "benefit_per_order",
    "Sales per customer":          "sales_per_customer",
    "Delivery Status":             "delivery_status",
    "Late_delivery_risk":          "late_delivery_risk",
    "Category Id":                 "category_id",
    "Category Name":               "category_name",
    "Customer City":               "customer_city",
    "Customer Country":            "customer_country",
    "Customer Fname":              "customer_fname",
    "Customer Id":                 "customer_id",
    "Customer Lname":              "customer_lname",
    "Customer Segment":            "customer_segment",
    "Customer State":              "customer_state",
    "Customer Street":             "customer_street",
    "Customer Zipcode":            "customer_zipcode",
    "Department Id":               "department_id",
    "Department Name":             "department_name",
    "Latitude":                    "latitude",
    "Longitude":                   "longitude",
    "Market":                      "market",
    "Order City":                  "order_city",
    "Order Country":               "order_country",
    "Order Customer Id":           "order_customer_id",
    "Order Item Discount":         "order_item_discount",
    "Order Item Discount Rate":    "order_item_discount_rate",
    "Order Item Product Price":    "order_item_product_price",
    "Order Item Profit Ratio":     "order_item_profit_ratio",
    "Order Item Quantity":         "order_item_quantity",
    "Sales":                       "sales",
    "Order Item Total":            "order_item_total",
    "Order Profit Per Order":      "order_profit_per_order",
    "Order Region":                "order_region",
    "Order State":                 "order_state",
    "Order Status":                "order_status",
    "Product Name":                "product_name",
    "Product Price":               "product_price",
    "Shipping Mode":               "shipping_mode",
}

# Shipping cost proxy (USD per unit, by mode)
SHIPPING_COST_MAP = {
    "Same Day":       15.0,
    "First Class":    10.0,
    "Second Class":    7.0,
    "Standard Class":  4.5,
}

DISCOUNT_BINS   = [-0.001, 0.0, 0.05, 0.10, 0.15, 0.20, 0.25]
DISCOUNT_LABELS = ["No Discount", "1-5%", "6-10%", "11-15%", "16-20%", "21-25%"]

DROP_COLUMNS = [
    "customer_street", "customer_zipcode",
    "latitude", "longitude",
]


# ─── Tier Assignment Helpers ──────────────────────────────────────────────────
def assign_customer_tier(total_profit, quartiles):
    if total_profit < 0:
        return "Loss Customer"
    if total_profit < quartiles[0.25]:
        return "Low Value"
    if total_profit < quartiles[0.50]:
        return "Mid Value"
    if total_profit < quartiles[0.75]:
        return "High Value"
    return "Premium"


def assign_product_tier(margin_pct, quartiles):
    if pd.isna(margin_pct) or margin_pct < 0:
        return "Loss Product"
    if margin_pct < quartiles[0.25]:
        return "Low Margin"
    if margin_pct < quartiles[0.75]:
        return "Moderate Margin"
    return "High Margin"


# ─── Main Pipeline ────────────────────────────────────────────────────────────
def main():
    print("=" * 65)
    print("APL LOGISTICS – DATA TRANSFORMATION PIPELINE")
    print("=" * 65)

    # 1. Ingest
    df = pd.read_csv(
        "data/APL_Logistics.csv", encoding="latin1", low_memory=False
    )
    raw_rows = len(df)
    print(f"[LOAD]  Raw records      : {raw_rows:,}  rows × {df.shape[1]} columns")

    # 2. Rename
    df.rename(columns=RENAME_MAP, inplace=True)

    # 3. Standardise strings
    obj_cols = df.select_dtypes(include="object").columns
    for col in obj_cols:
        df[col] = df[col].astype(str).str.strip()

    df["order_status"]      = df["order_status"].str.upper()
    df["delivery_status"]   = df["delivery_status"].str.title()
    df["shipping_mode"]     = df["shipping_mode"].str.title()
    df["customer_segment"]  = df["customer_segment"].str.title()
    df["payment_type"]      = df["payment_type"].str.upper()
    df["market"] = (
        df["market"].str.strip()
        .replace({"Usca": "USCA", "Latam": "LATAM",
                  "usca": "USCA", "latam": "LATAM"})
    )

    # 4. Numeric coercions
    df["customer_zipcode"] = pd.to_numeric(df["customer_zipcode"], errors="coerce")
    df["customer_lname"].fillna("Unknown", inplace=True)

    # 5. Drop non-positive sales (data-entry errors)
    before = len(df)
    df = df[df["sales"] > 0].copy()
    removed = before - len(df)
    print(f"[CLEAN] Removed          : {removed:,} non-positive sales records")
    print(f"[INFO]  Retained records : {len(df):,}")

    # ── 6. DERIVED FINANCIAL FEATURES ────────────────────────────────────────
    df["gross_margin_pct"] = np.where(
        df["sales"] != 0,
        (df["order_profit_per_order"] / df["sales"]) * 100,
        np.nan,
    )
    df["discount_amount"]         = df["order_item_discount"]
    df["revenue_after_discount"]  = df["order_item_total"]
    df["effective_unit_price"] = np.where(
        df["order_item_quantity"] > 0,
        df["order_item_total"] / df["order_item_quantity"],
        df["order_item_product_price"],
    )
    df["unit_profit"] = np.where(
        df["order_item_quantity"] > 0,
        df["order_profit_per_order"] / df["order_item_quantity"],
        0.0,
    )

    # ── 7. SHIPPING / DELIVERY FEATURES ──────────────────────────────────────
    df["shipping_delay_days"] = (
        df["shipping_days_actual"] - df["shipping_days_scheduled"]
    )
    df["is_late_delivery"] = (df["late_delivery_risk"] == 1).astype(int)
    df["shipping_efficiency"] = np.where(
        df["shipping_days_scheduled"] > 0,
        df["shipping_days_actual"] / df["shipping_days_scheduled"],
        1.0,
    )
    df["shipping_cost_proxy"] = df["shipping_mode"].map(SHIPPING_COST_MAP).fillna(4.5)
    df["shipping_cost_total"] = df["shipping_cost_proxy"] * df["order_item_quantity"]
    df["profit_after_shipping"] = df["order_profit_per_order"] - df["shipping_cost_total"]

    # ── 8. DISCOUNT FEATURES ──────────────────────────────────────────────────
    df["discount_impact_on_profit"] = df["discount_amount"] - df["order_profit_per_order"]
    df["discount_erodes_profit"] = (
        df["discount_amount"] > df["order_profit_per_order"]
    ).astype(int)
    df["net_margin_after_discount"] = np.where(
        df["order_item_total"] != 0,
        (df["order_profit_per_order"] / df["order_item_total"]) * 100,
        np.nan,
    )
    df["discount_band"] = pd.cut(
        df["order_item_discount_rate"],
        bins=DISCOUNT_BINS,
        labels=DISCOUNT_LABELS,
    )

    # ── 9. PROFITABILITY CLASSIFICATION ──────────────────────────────────────
    conditions = [
        df["order_profit_per_order"] < 0,
        df["order_profit_per_order"] == 0,
        (df["order_profit_per_order"] > 0) & (df["gross_margin_pct"] < 10),
        (df["gross_margin_pct"] >= 10) & (df["gross_margin_pct"] < 25),
        df["gross_margin_pct"] >= 25,
    ]
    classes = ["Loss-Making", "Break-Even", "Low-Margin", "Moderate-Margin", "High-Margin"]
    df["profitability_class"] = np.select(conditions, classes, default="Unknown")

    # ── 10. ORDER HELPER FLAGS ────────────────────────────────────────────────
    df["is_express_shipping"]  = df["shipping_mode"].isin(["Same Day", "First Class"]).astype(int)
    sales_95th                 = df["sales"].quantile(0.95)
    df["is_high_value_order"]  = (df["sales"] >= sales_95th).astype(int)
    df["is_order_cancelled"]   = (df["order_status"] == "CANCELED").astype(int)
    df["customer_name"]        = df["customer_fname"].str.strip() + " " + df["customer_lname"].str.strip()

    # ── 11. MARGIN EROSION RISK SCORE ─────────────────────────────────────────
    norm_disc = (df["order_item_discount_rate"].clip(0, 0.25) / 0.25) * 50
    df["margin_erosion_risk"] = (
        norm_disc + df["late_delivery_risk"] * 20 + df["is_order_cancelled"] * 30
    ).clip(0, 100)

    # ── 12. CUSTOMER AGGREGATION (tier assignment) ────────────────────────────
    cust_agg = (
        df.groupby("customer_id")
        .agg(
            cust_total_sales   =("sales", "sum"),
            cust_total_profit  =("order_profit_per_order", "sum"),
            cust_order_count   =("sales", "count"),
            cust_avg_discount  =("order_item_discount_rate", "mean"),
            cust_avg_margin    =("gross_margin_pct", "mean"),
        )
        .reset_index()
    )
    cust_agg["cust_avg_order_value"] = cust_agg["cust_total_sales"] / cust_agg["cust_order_count"]
    cust_agg["cust_profit_margin"]   = np.where(
        cust_agg["cust_total_sales"] != 0,
        cust_agg["cust_total_profit"] / cust_agg["cust_total_sales"] * 100,
        np.nan,
    )
    pos_cust_profit   = cust_agg.loc[cust_agg["cust_total_profit"] >= 0, "cust_total_profit"]
    customer_quartiles = pos_cust_profit.quantile([0.25, 0.50, 0.75])
    cust_agg["customer_value_tier"] = cust_agg["cust_total_profit"].apply(
        lambda v: assign_customer_tier(v, customer_quartiles)
    )
    df = df.merge(cust_agg, on="customer_id", how="left")

    # ── 13. PRODUCT AGGREGATION ───────────────────────────────────────────────
    prod_agg = (
        df.groupby("product_name")
        .agg(
            prod_total_sales  =("sales", "sum"),
            prod_total_profit =("order_profit_per_order", "sum"),
            prod_order_count  =("sales", "count"),
            prod_avg_margin   =("gross_margin_pct", "mean"),
        )
        .reset_index()
    )
    prod_agg["prod_profit_margin"] = np.where(
        prod_agg["prod_total_sales"] != 0,
        prod_agg["prod_total_profit"] / prod_agg["prod_total_sales"] * 100,
        np.nan,
    )
    pos_prod_margin   = prod_agg.loc[prod_agg["prod_profit_margin"] >= 0, "prod_profit_margin"]
    product_quartiles = pos_prod_margin.quantile([0.25, 0.75])
    prod_agg["product_margin_tier"] = prod_agg["prod_profit_margin"].apply(
        lambda v: assign_product_tier(v, product_quartiles)
    )
    df = df.merge(prod_agg, on="product_name", how="left")

    # ── 14. MARKET AGGREGATION ────────────────────────────────────────────────
    mkt_agg = (
        df.groupby("market")
        .agg(
            mkt_total_sales  =("sales", "sum"),
            mkt_total_profit =("order_profit_per_order", "sum"),
            mkt_order_count  =("sales", "count"),
        )
        .reset_index()
    )
    mkt_agg["mkt_profit_margin"] = np.where(
        mkt_agg["mkt_total_sales"] != 0,
        mkt_agg["mkt_total_profit"] / mkt_agg["mkt_total_sales"] * 100,
        np.nan,
    )
    df = df.merge(mkt_agg, on="market", how="left")

    # ── 15. CATEGORY AGGREGATION ──────────────────────────────────────────────
    cat_agg = (
        df.groupby("category_name")
        .agg(
            cat_total_sales  =("sales", "sum"),
            cat_total_profit =("order_profit_per_order", "sum"),
            cat_avg_discount =("order_item_discount_rate", "mean"),
        )
        .reset_index()
    )
    cat_agg["cat_margin_pct"] = np.where(
        cat_agg["cat_total_sales"] != 0,
        cat_agg["cat_total_profit"] / cat_agg["cat_total_sales"] * 100,
        np.nan,
    )
    df = df.merge(cat_agg, on="category_name", how="left")

    # ── 16. DROP REDUNDANT COLUMNS ────────────────────────────────────────────
    df.drop(columns=[c for c in DROP_COLUMNS if c in df.columns], inplace=True)

    # ── 17. ROUND FLOATS ──────────────────────────────────────────────────────
    float_cols = df.select_dtypes(include="float64").columns
    df[float_cols] = df[float_cols].round(4)

    # ─── OUTPUT ───────────────────────────────────────────────────────────────
    out_path = "data/APL_Logistics_Transformed.csv"
    df.to_csv(out_path, index=False)

    print(f"[INFO]  Transformed      : {df.shape[0]:,} rows × {df.shape[1]} columns")
    print(f"[INFO]  Engineered feats : {df.shape[1] - len(RENAME_MAP)} new fields")
    print(f"[DONE]  Saved -> {out_path}")

    # ─── FINANCIAL VERIFICATION SUMMARY ──────────────────────────────────────
    total_rev    = df["sales"].sum()
    total_prof   = df["order_profit_per_order"].sum()
    total_disc   = df["order_item_discount"].sum()
    overall_marg = total_prof / total_rev * 100
    avg_disc_r   = df["order_item_discount_rate"].mean() * 100
    loss_n       = (df["profitability_class"] == "Loss-Making").sum()
    late_pct     = df["is_late_delivery"].mean() * 100

    print()
    print("=" * 65)
    print("FINANCIAL VERIFICATION SUMMARY")
    print("=" * 65)
    print(f"  Total Revenue          : ${total_rev:>18,.2f}")
    print(f"  Total Profit           : ${total_prof:>18,.2f}")
    print(f"  Overall Profit Margin  : {overall_marg:>17.2f}%")
    print(f"  Total Discounts Given  : ${total_disc:>18,.2f}")
    print(f"  Average Discount Rate  : {avg_disc_r:>17.2f}%")
    print(f"  Loss-Making Orders     : {loss_n:>18,}")
    print(f"  Loss-Making Rate       : {loss_n/len(df)*100:>17.2f}%")
    print(f"  Late Delivery Rate     : {late_pct:>17.2f}%")
    print(f"  Discount / Profit Ratio: {total_disc/total_prof*100:>17.1f}%")
    print("=" * 65)


if __name__ == "__main__":
    main()
