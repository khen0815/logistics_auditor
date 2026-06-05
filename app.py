import math

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="E-commerce Logistics Auditor", layout="wide")

st.title("E-commerce Logistics Auditor")
st.caption("Audit volumetric shipping penalties and SKU revenue concentration across South African e-commerce orders.")

uploaded_file = st.sidebar.file_uploader("Upload shipping CSV", type=["csv"])

if uploaded_file is not None:
    data = pd.read_csv(uploaded_file)
else:
    data = pd.read_csv("mock_shipping_data.csv")

required_columns = {
    "Order_ID",
    "SKU",
    "Revenue",
    "Province",
    "Actual_Weight_kg",
    "Length_cm",
    "Width_cm",
    "Height_cm",
    "Billed_Shipping_Cost_ZAR",
}

missing_columns = required_columns.difference(data.columns)
if missing_columns:
    st.error(f"CSV is missing required columns: {', '.join(sorted(missing_columns))}")
    st.stop()

for column in [
    "Revenue",
    "Actual_Weight_kg",
    "Length_cm",
    "Width_cm",
    "Height_cm",
    "Billed_Shipping_Cost_ZAR",
]:
    data[column] = pd.to_numeric(data[column], errors="coerce")

data = data.dropna(subset=required_columns).copy()

data["Volumetric_Weight_kg"] = (
    data["Length_cm"] * data["Width_cm"] * data["Height_cm"]
) / 5000

data["Volumetric_Penalty"] = data["Volumetric_Weight_kg"] > data["Actual_Weight_kg"]
data["Excess_Weight_kg"] = (
    data["Volumetric_Weight_kg"] - data["Actual_Weight_kg"]
).clip(lower=0)
data["Estimated_Loss_ZAR"] = data["Excess_Weight_kg"] * 15

sku_revenue = (
    data.groupby("SKU", as_index=False)["Revenue"]
    .sum()
    .sort_values("Revenue", ascending=False)
    .reset_index(drop=True)
)

sku_count = len(sku_revenue)
a_cutoff = math.ceil(sku_count * 0.20)
b_cutoff = a_cutoff + math.ceil(sku_count * 0.30)

sku_revenue["ABC_Class"] = "C-Class (Dead Stock)"
sku_revenue.loc[: a_cutoff - 1, "ABC_Class"] = "A-Class (High Volume)"
sku_revenue.loc[a_cutoff:b_cutoff - 1, "ABC_Class"] = "B-Class"

data = data.merge(sku_revenue[["SKU", "ABC_Class"]], on="SKU", how="left")

penalty_orders = data[data["Volumetric_Penalty"]].copy()
total_orders = len(data)
total_loss = penalty_orders["Estimated_Loss_ZAR"].sum()
c_class_count = sku_revenue[sku_revenue["ABC_Class"] == "C-Class (Dead Stock)"]["SKU"].nunique()

metric_1, metric_2, metric_3 = st.columns(3)
metric_1.metric("Total Orders", f"{total_orders:,}")
metric_2.metric("Money Lost to Volumetric Penalties (ZAR)", f"R{total_loss:,.2f}")
metric_3.metric("Count of C-Class Dead Stock SKUs", f"{c_class_count:,}")

st.divider()

province_inefficiencies = (
    penalty_orders.groupby("Province", as_index=False)
    .agg(
        Estimated_Loss_ZAR=("Estimated_Loss_ZAR", "sum"),
        Penalty_Orders=("Order_ID", "count"),
    )
    .sort_values("Estimated_Loss_ZAR", ascending=False)
)

st.subheader("Shipping Inefficiencies by Province")
fig = px.bar(
    province_inefficiencies,
    x="Province",
    y="Estimated_Loss_ZAR",
    color="Penalty_Orders",
    labels={
        "Estimated_Loss_ZAR": "Estimated Loss (ZAR)",
        "Penalty_Orders": "Penalty Orders",
    },
    text_auto=".2s",
)
fig.update_layout(xaxis_title="Province", yaxis_title="Estimated Loss (ZAR)")
st.plotly_chart(fig, use_container_width=True)

st.subheader("Orders Triggering Volumetric Penalties")
st.dataframe(
    penalty_orders[
        [
            "Order_ID",
            "SKU",
            "Province",
            "Actual_Weight_kg",
            "Volumetric_Weight_kg",
            "Excess_Weight_kg",
            "Estimated_Loss_ZAR",
            "Billed_Shipping_Cost_ZAR",
            "ABC_Class",
        ]
    ].sort_values("Estimated_Loss_ZAR", ascending=False),
    use_container_width=True,
    hide_index=True,
)

with st.expander("ABC Inventory Analysis by SKU"):
    st.dataframe(sku_revenue, use_container_width=True, hide_index=True)
