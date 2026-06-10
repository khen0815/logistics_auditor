import io
import math
import re
from itertools import permutations

import pandas as pd
import pdfplumber
import plotly.express as px
import streamlit as st
from aramex_html_processor import clean_selected_table, select_shipment_table
from fpdf import FPDF
from supabase import create_client


st.set_page_config(
    page_title="360° Margin Diagnostic Engine",
    layout="wide",
    initial_sidebar_state="expanded",
)


STANDARD_COLUMNS = [
    "Order_ID",
    "Waybill_ID",
    "SKU",
    "Province",
    "Actual_Weight_KG",
    "Billed_Vol_KG",
    "Billed_Cost_ZAR",
    "Length_cm",
    "Width_cm",
    "Height_cm",
]

NUMERIC_COLUMNS = [
    "Actual_Weight_KG",
    "Billed_Vol_KG",
    "Billed_Cost_ZAR",
    "Length_cm",
    "Width_cm",
    "Height_cm",
]

COLUMN_SYNONYMS = {
    "Order_ID": [
        "order id", "order_id", "order no", "order number", "order", "waybill", "waybill no",
        "waybill number", "tracking", "tracking number", "shipment id", "parcel id", "reference",
    ],
    "SKU": ["sku", "product", "product code", "item", "item code", "description", "product description"],
    "Province": ["province", "destination province", "dest province", "region", "zone", "destination", "ship province"],
    "Actual_Weight_KG": [
        "actual kg", "actual weight", "actual weight kg", "actual_weight_kg", "wgt", "weight", "weight kg",
        "mass", "actual mass", "dead weight", "scale weight", "parcel weight", "submitted weight kg",
    ],
    "Billed_Vol_KG": [
        "vol wgt", "vol weight", "volumetric weight", "volumetric weight kg", "billed volumetric",
        "billed volumetric weight", "billed vol", "billed vol kg", "billed_vol_kg", "chargeable weight",
        "billed weight", "csv volumetric weight", "vol kg", "charged weight kg",
    ],
    "Billed_Cost_ZAR": [
        "billed amount", "cost zar", "billed cost", "billed cost zar", "amount", "charge", "shipping cost",
        "shipping cost zar", "billed shipping cost", "billed shipping cost zar", "total", "invoice amount",
        "courier cost", "freight charge", "charged rate incl vat",
    ],
    "Length_cm": ["l", "len", "length", "length cm", "length (cm)", "length_cm", "parcel length"],
    "Width_cm": ["w", "wid", "width", "width cm", "width (cm)", "width_cm", "parcel width"],
    "Height_cm": ["h", "hei", "height", "height cm", "height (cm)", "height_cm", "parcel height"],
    "Dimensions": ["dimensions", "dims", "parcel dimensions", "size", "l x w x h", "lxwxh", "submitted dimensions"],
    "Revenue": ["revenue", "sales value", "item price", "selling price", "order value"],
}

BOB_GO_DIRECT_MAPPINGS = {
    "submitted weight kg": "Actual_Weight_KG",
    "charged weight kg": "Billed_Vol_KG",
    "charged rate incl vat": "Billed_Cost_ZAR",
}

BOB_GO_DIMENSION_PATTERN = re.compile(
    r"^\s*(?P<waybill>[^\s\-–—]+).*?"
    r"(?P<length>\d+(?:\.\d+)?)\s*cm\s*x\s*"
    r"(?P<width>\d+(?:\.\d+)?)\s*cm\s*x\s*"
    r"(?P<height>\d+(?:\.\d+)?)\s*cm",
    re.I,
)

DEFAULT_PACKAGING_MATRIX = pd.DataFrame(
    [
        {"Package": "A4 Flyer", "L": 30.0, "W": 21.0, "H": 2.0, "cost": 1.50, "Fragile/Void Fill Required": False},
        {"Package": "A3 Flyer", "L": 42.0, "W": 30.0, "H": 2.0, "cost": 2.50, "Fragile/Void Fill Required": False},
        {"Package": "Small Box", "L": 25.0, "W": 15.0, "H": 10.0, "cost": 5.00, "Fragile/Void Fill Required": False},
        {"Package": "Medium Box", "L": 35.0, "W": 25.0, "H": 15.0, "cost": 8.00, "Fragile/Void Fill Required": False},
        {"Package": "Large Corrugated Box", "L": 45.0, "W": 35.0, "H": 25.0, "cost": 12.00, "Fragile/Void Fill Required": False},
    ]
)

DISTANT_ZONE_PATTERN = re.compile(r"limpopo|mpumalanga|northern cape|eastern cape|free state|rural|remote|outlying", re.I)
FLYER_PATTERN = re.compile(r"flyer|satchel|bag", re.I)
CORRUGATED_PATTERN = re.compile(r"box|corrugated|carton", re.I)


@st.cache_resource
def get_supabase_client():
    try:
        return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        return None


def normalize_label(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def build_alias_lookup():
    lookup = {}
    for standard_name, aliases in COLUMN_SYNONYMS.items():
        lookup[normalize_label(standard_name)] = standard_name
        for alias in aliases:
            lookup[normalize_label(alias)] = standard_name
    return lookup


ALIAS_LOOKUP = build_alias_lookup()


def parse_numeric(value):
    if pd.isna(value):
        return pd.NA
    cleaned = re.sub(r"[^0-9.\-]", "", str(value))
    return pd.to_numeric(cleaned, errors="coerce")


def parse_dimensions(value):
    if pd.isna(value):
        return None
    numbers = re.findall(r"\d+(?:\.\d+)?", str(value))
    if len(numbers) < 3:
        return None
    return [float(numbers[0]), float(numbers[1]), float(numbers[2])]


def parse_bob_go_dimensions(value):
    if pd.isna(value):
        return {"Waybill_ID": pd.NA, "Length_cm": pd.NA, "Width_cm": pd.NA, "Height_cm": pd.NA}
    match = BOB_GO_DIMENSION_PATTERN.search(str(value))
    if not match:
        return {"Waybill_ID": pd.NA, "Length_cm": pd.NA, "Width_cm": pd.NA, "Height_cm": pd.NA}
    return {
        "Waybill_ID": match.group("waybill"),
        "Length_cm": float(match.group("length")),
        "Width_cm": float(match.group("width")),
        "Height_cm": float(match.group("height")),
    }


def safe_float(value, default=0.0):
    value = pd.to_numeric(value, errors="coerce")
    return float(value) if pd.notna(value) else default


def safe_text(value, default="Unknown"):
    if pd.isna(value) or str(value).strip() == "":
        return default
    return str(value).strip()


def format_currency(value, currency_prefix):
    return f"{currency_prefix}{safe_float(value):,.2f}"


def format_currency_whole(value, currency_prefix):
    return f"{currency_prefix}{safe_float(value):,.0f}"


def display_currency_columns(df, currency_symbol):
    return df.rename(
        columns={
            "Billed_Cost_ZAR": f"Billed_Cost_{currency_symbol}",
            "Estimated_Loss_ZAR": f"Estimated_Loss_{currency_symbol}",
            "Recoverable_Overcharge_ZAR": f"Recoverable_Overcharge_{currency_symbol}",
            "Avoidable_Volumetric_Leak_ZAR": f"Avoidable_Volumetric_Leak_{currency_symbol}",
            "SKU_Total_Shipping_Cost_ZAR": f"SKU_Total_Shipping_Cost_{currency_symbol}",
            "SKU_Total_Leakage_ZAR": f"SKU_Total_Leakage_{currency_symbol}",
            "Multi_Item_Leak_ZAR": f"Multi_Item_Leak_{currency_symbol}",
        }
    )


def clean_and_standardize_data(df):
    """Map messy courier exports into a deterministic standard schema without crashing on bad rows."""
    cleaned = df.copy()
    cleaned.columns = [str(col).strip() for col in cleaned.columns]

    bob_go_dimensions_col = next((col for col in ["Charged dimensions", "Submitted dimensions"] if col in cleaned.columns), None)
    bob_go_parsed = None
    if bob_go_dimensions_col:
        bob_go_parsed = pd.DataFrame(
            cleaned[bob_go_dimensions_col].apply(parse_bob_go_dimensions).tolist(),
            index=cleaned.index,
            columns=["Waybill_ID", "Length_cm", "Width_cm", "Height_cm"],
        )

    rename_map = {}
    used_standard_names = set()
    for original in cleaned.columns:
        normalized = normalize_label(original)
        standard = BOB_GO_DIRECT_MAPPINGS.get(normalized) or ALIAS_LOOKUP.get(normalized)
        if standard and standard not in used_standard_names:
            rename_map[original] = standard
            used_standard_names.add(standard)
    cleaned = cleaned.rename(columns=rename_map)

    try:
        if bob_go_parsed is not None:
            for col in ["Waybill_ID", "Length_cm", "Width_cm", "Height_cm"]:
                if col not in cleaned.columns:
                    cleaned[col] = bob_go_parsed[col]
                else:
                    cleaned[col] = cleaned[col].combine_first(bob_go_parsed[col])
            if "Order_ID" not in cleaned.columns:
                cleaned["Order_ID"] = cleaned["Waybill_ID"]
            else:
                cleaned["Order_ID"] = cleaned["Order_ID"].combine_first(cleaned["Waybill_ID"])

        if "Dimensions" in cleaned.columns:
            parsed_dims = cleaned["Dimensions"].apply(parse_dimensions)
            if "Length_cm" not in cleaned.columns:
                cleaned["Length_cm"] = parsed_dims.apply(lambda dims: dims[0] if dims else pd.NA)
            if "Width_cm" not in cleaned.columns:
                cleaned["Width_cm"] = parsed_dims.apply(lambda dims: dims[1] if dims else pd.NA)
            if "Height_cm" not in cleaned.columns:
                cleaned["Height_cm"] = parsed_dims.apply(lambda dims: dims[2] if dims else pd.NA)

        for col in STANDARD_COLUMNS:
            if col not in cleaned.columns:
                cleaned[col] = pd.NA

        cleaned["Order_ID"] = cleaned["Order_ID"].fillna(pd.Series([f"ROW-{i + 1}" for i in range(len(cleaned))], index=cleaned.index))
        cleaned["SKU"] = cleaned["SKU"].fillna("Unknown SKU")
        cleaned["Province"] = cleaned["Province"].fillna("Unknown")

        for col in NUMERIC_COLUMNS:
            cleaned[col] = cleaned[col].apply(parse_numeric)

        cleaned["Data_Quality_Issue"] = cleaned[NUMERIC_COLUMNS].isna().any(axis=1)
        cleaned["Data_Quality_Note"] = cleaned.apply(
            lambda row: "Missing critical numeric courier data; row skipped from pillar math."
            if row["Data_Quality_Issue"] else "OK",
            axis=1,
        )
    except Exception as exc:
        st.warning(f"Some dirty-data cleanup failed gracefully and affected rows were marked for review: {exc}")
        for col in STANDARD_COLUMNS:
            if col not in cleaned.columns:
                cleaned[col] = pd.NA
        cleaned["Data_Quality_Issue"] = True
        cleaned["Data_Quality_Note"] = "Cleanup exception; row requires manual review."

    return cleaned


def fits_in_package(item_dims, package_dims):
    if any(pd.isna(dim) or safe_float(dim) <= 0 for dim in item_dims + package_dims):
        return False
    return any(all(item <= package for item, package in zip(ordering, package_dims)) for ordering in permutations(item_dims))


def clean_packaging_matrix(packaging_matrix):
    matrix = packaging_matrix.copy()
    required_columns = ["Package", "L", "W", "H", "cost"]
    for col in required_columns:
        if col not in matrix.columns:
            matrix[col] = pd.NA
    for col in ["L", "W", "H", "cost"]:
        matrix[col] = pd.to_numeric(matrix[col], errors="coerce")
    matrix = matrix.dropna(subset=required_columns)
    matrix = matrix[(matrix["L"] > 0) & (matrix["W"] > 0) & (matrix["H"] > 0)]
    matrix["Volume_cm3"] = matrix["L"] * matrix["W"] * matrix["H"]
    return matrix.sort_values("Volume_cm3").reset_index(drop=True)


def run_single_item_repack(row, packaging_matrix, penalty_rate, divisor):
    actual_weight = safe_float(row["Actual_Weight_KG"])
    billed_vol = safe_float(row["Billed_Vol_KG"])
    item_dims = [safe_float(row["Length_cm"]), safe_float(row["Width_cm"]), safe_float(row["Height_cm"])]
    billed_weight = max(actual_weight, billed_vol)
    current_penalty = max(0, billed_weight - actual_weight) * penalty_rate

    base = {
        "Packaging_Flag": False,
        "Packaging_Reason": "No lower-cost flyer fit found.",
        "Recommended_Package": "No recommendation",
        "Optimized_Vol_KG": pd.NA,
        "Avoidable_Volumetric_Leak_ZAR": 0.0,
    }

    try:
        if actual_weight <= 0 or billed_vol <= 0 or any(dim <= 0 for dim in item_dims):
            return {**base, "Packaging_Reason": "Missing or invalid single-item dimensions/weight."}

        if packaging_matrix.empty:
            return {**base, "Packaging_Reason": "Packaging matrix has no valid numeric package rows."}

        for _, package in packaging_matrix.iterrows():
            usable_factor = 0.85 if bool(package.get("Fragile/Void Fill Required", False)) else 1.0
            package_dims = [safe_float(package["L"]) * usable_factor, safe_float(package["W"]) * usable_factor, safe_float(package["H"]) * usable_factor]
            if fits_in_package(item_dims, package_dims):
                package_name = str(package["Package"])
                optimized_vol = (safe_float(package["L"]) * safe_float(package["W"]) * safe_float(package["H"])) / max(divisor, 1)
                optimized_billed_weight = max(actual_weight, optimized_vol)
                optimized_penalty = max(0, optimized_billed_weight - actual_weight) * penalty_rate
                net_savings = max(0, current_penalty - optimized_penalty - safe_float(package["cost"]))
                flyer_fit = bool(FLYER_PATTERN.search(package_name))
                return {
                    "Packaging_Flag": flyer_fit and net_savings > 0,
                    "Packaging_Reason": "Single-item order fits into a standard flyer." if flyer_fit and net_savings > 0 else "Fits smaller packaging but no net flyer leak after material cost.",
                    "Recommended_Package": package_name,
                    "Optimized_Vol_KG": optimized_vol,
                    "Avoidable_Volumetric_Leak_ZAR": net_savings if flyer_fit else 0.0,
                }
        return {**base, "Packaging_Reason": "No configured package fits this order; marked as custom/oversized.", "Recommended_Package": "Custom/Oversized"}
    except Exception as exc:
        order_id = row.get("Order_ID", "unknown")
        st.warning(f"Packaging calculation failed for row {row.name} / order {order_id}: {exc}")
        return {**base, "Packaging_Reason": f"Packaging calculation error: {exc}", "Recommended_Package": "Review manually"}


def find_standard_column(df, aliases):
    normalized_lookup = {normalize_label(column): column for column in df.columns}
    for alias in aliases:
        normalized_alias = normalize_label(alias)
        if normalized_alias in normalized_lookup:
            return normalized_lookup[normalized_alias]
    for column in df.columns:
        normalized_column = normalize_label(column)
        if any(normalize_label(alias) in normalized_column for alias in aliases):
            return column
    return None


def generate_client_summary(df, anomalies_df):
    shipment_count = len(df)
    if anomalies_df is None or anomalies_df.empty:
        return (
            f"Courier audit complete: {shipment_count:,} shipments were reviewed.\n\n"
            "No material financial or weight anomalies were detected in the current ruleset. "
            "This means the file did not show surcharge spikes or same-weight base-charge mismatches above the configured thresholds."
        )

    financial_recovery = safe_float(anomalies_df.get("Financial_Excess_ZAR", pd.Series(dtype=float)).sum())
    weight_recovery = safe_float(anomalies_df.get("Recoverable_Overcharge_ZAR", pd.Series(dtype=float)).sum())
    total_recovery = financial_recovery + weight_recovery

    if "Financial_Excess_ZAR" in anomalies_df.columns and anomalies_df["Financial_Excess_ZAR"].notna().any():
        largest = anomalies_df.sort_values("Financial_Excess_ZAR", ascending=False).iloc[0]
        largest_amount = safe_float(largest.get("Financial_Excess_ZAR", 0))
    else:
        largest = anomalies_df.iloc[0]
        largest_amount = safe_float(largest.get("Recoverable_Overcharge_ZAR", 0))

    weight = safe_float(largest.get("Financial_Weight_KG", largest.get("Actual_Weight_KG", 0)))
    destination = safe_text(largest.get("Destination", largest.get("Province", "Unknown destination")), "Unknown destination")
    reason = safe_text(largest.get("Financial_Anomaly_Reason", largest.get("Anomaly_Reason", "unexplained courier charge")), "unexplained courier charge")

    return (
        f"Courier audit complete: {shipment_count:,} shipments were reviewed.\n\n"
        f"The audit identified {format_currency(total_recovery, currency_prefix)} in potential margin recovery across financial, routing, and weight anomalies.\n\n"
        f"The largest anomaly was a {weight:.2f} kg package to {destination} that incurred an unexplained surcharge/excess charge of {format_currency(largest_amount, currency_prefix)}. "
        f"Reason flagged: {reason}\n\n"
        "Recommended next step: query these rows with the courier account manager and request supporting charge/routing evidence for each flagged shipment."
    )


def calculate_financial_anomalies(df):
    """Find financial anomalies with hard guards around every sparse-array calculation."""
    if df is None or df.empty:
        return pd.DataFrame()

    try:
        working = df.copy()
        other_col = find_standard_column(working, ["Other Charges", "Other Charge", "Surcharge", "Surcharges", "Additional Charge"])
        base_col = find_standard_column(working, ["cost", "Cost", "Base Charge", "Base", "Basic Charge", "Freight Charge", "Billed_Cost_ZAR", "Financial_Base_Charge_ZAR"])
        weight_col = find_standard_column(working, ["weight", "Weight", "Actual_Weight_KG", "Actual Weight", "Charged weight", "Billed_Weight_KG", "Financial_Weight_KG"])
        tracking_col = find_standard_column(working, ["Tracking_Number", "HAWB", "Shipper Ref.", "Shipper Ref", "Waybill", "Waybill Number", "Tracking Number", "Order_ID"])
        shipper_ref_col = find_standard_column(working, ["Shipper Ref.", "Shipper Ref", "Shipper Reference", "Reference"])
        destination_col = find_standard_column(working, ["Destination", "Destination City", "Receiver City", "Consignee City", "Province"])
        start_town_col = "Start Town" if "Start Town" in working.columns else None
        service_col = "Srv" if "Srv" in working.columns else None

        working["Tracking_Number"] = working[tracking_col].apply(lambda value: safe_text(value, "")) if tracking_col else ""
        working["Shipper_Ref"] = working[shipper_ref_col].apply(lambda value: safe_text(value, "")) if shipper_ref_col else ""
        working["Destination"] = working[destination_col].apply(lambda value: safe_text(value, "")) if destination_col else ""
        working["Start Town"] = working[start_town_col].apply(lambda value: safe_text(value, "")) if start_town_col else ""
        working["Srv"] = working[service_col].apply(lambda value: safe_text(value, "")) if service_col else ""
        working["Financial_Other_Charges_ZAR"] = clean_money_series(working[other_col]) if other_col else 0.0
        working["Financial_Base_Charge_ZAR"] = clean_money_series(working[base_col]) if base_col else 0.0
        working["Financial_Weight_KG"] = clean_numeric_series(working[weight_col]) if weight_col else 0.0
        working["cost"] = pd.to_numeric(working["Financial_Base_Charge_ZAR"], errors="coerce")

        anomaly_frames = []

        valid_other = pd.to_numeric(working["Financial_Other_Charges_ZAR"], errors="coerce").dropna()
        valid_other = valid_other[valid_other > 0]
        if len(valid_other) > 1:
            try:
                median_other = valid_other.median()
                std_other = valid_other.std()
                if pd.isna(median_other):
                    raise ValueError("empty other-charge median")
                statistical_threshold = median_other + (3 * std_other) if pd.notna(std_other) and std_other > 0 else median_other
                other_threshold = max(50, statistical_threshold)
                all_other = pd.to_numeric(working["Financial_Other_Charges_ZAR"], errors="coerce").fillna(0.0)
                other_spike = all_other > other_threshold
                if other_spike.any():
                    spike_rows = working.loc[other_spike].copy()
                    spike_rows["Financial_Anomaly_Flag"] = True
                    spike_rows["Financial_Anomaly_Reason"] = "Other Charges spike above normal dataset threshold."
                    spike_rows["Financial_Excess_ZAR"] = (all_other.loc[other_spike] - max(median_other, 0)).clip(lower=0)
                    anomaly_frames.append(spike_rows)
            except ValueError:
                pass
            except Exception:
                pass

        comparable = working.copy()
        comparable["Financial_Weight_KG"] = pd.to_numeric(comparable["Financial_Weight_KG"], errors="coerce")
        comparable["cost"] = pd.to_numeric(comparable["cost"], errors="coerce")
        comparable = comparable[(comparable["Financial_Weight_KG"] > 0) & (comparable["cost"] > 0)].copy()

        if len(comparable) > 1:
            comparable["weight"] = comparable["Financial_Weight_KG"].round(3)
            group_cols = ["Destination", "Srv", "weight"]
            if start_town_col:
                group_cols.insert(0, "Start Town")
            comparable = comparable.dropna(subset=group_cols)
            comparable = comparable[(comparable["Destination"] != "") & (comparable["Srv"] != "")].copy()
            if start_town_col:
                comparable = comparable[comparable["Start Town"] != ""].copy()

            for _, group in comparable.groupby(group_cols, dropna=True):
                if len(group) <= 1:
                    continue

                valid_costs = pd.to_numeric(group.get("cost", pd.Series(dtype=float)), errors="coerce").dropna()
                valid_costs = valid_costs[valid_costs > 0]
                if len(valid_costs) <= 1:
                    continue

                try:
                    mode_values = valid_costs.mode(dropna=True)
                    baseline_cost = mode_values.iloc[0] if not mode_values.empty else valid_costs.median()
                    if pd.isna(baseline_cost) or baseline_cost <= 0:
                        raise ValueError("invalid strict-route baseline cost")

                    excess_base = (valid_costs - baseline_cost).clip(lower=0)
                    routing_spike = (excess_base > 50) & (valid_costs > baseline_cost * 1.5)
                    if not routing_spike.any():
                        continue

                    routing_indexes = valid_costs.index[routing_spike]
                    if routing_indexes.empty:
                        continue

                    routing_rows = working.loc[routing_indexes].copy()
                    routing_rows["Financial_Anomaly_Flag"] = True
                    routing_rows["Financial_Anomaly_Reason"] = "Shipment has unusually high Base Charge versus identical route, service, and weight group."
                    routing_rows["Financial_Excess_ZAR"] = excess_base.loc[routing_indexes]
                    anomaly_frames.append(routing_rows)
                except ValueError:
                    continue
                except Exception:
                    continue

        if not anomaly_frames:
            return pd.DataFrame()

        anomaly_rows = pd.concat(anomaly_frames, ignore_index=False, sort=False)
        if anomaly_rows.empty:
            return pd.DataFrame()

        anomaly_rows = (
            anomaly_rows.sort_values("Financial_Excess_ZAR", ascending=False)
            .drop_duplicates(subset=["Tracking_Number", "Financial_Anomaly_Reason", "Financial_Excess_ZAR"], keep="first")
            .reset_index(drop=True)
        )
        priority_cols = ["Tracking_Number", "Destination", "Shipper_Ref"]
        remaining_cols = [col for col in anomaly_rows.columns if col not in priority_cols]
        return anomaly_rows[[col for col in priority_cols if col in anomaly_rows.columns] + remaining_cols]
    except Exception as exc:
        st.warning(f"Financial anomaly calculation skipped safely: {exc}")
        return pd.DataFrame()


def classify_velocity(data):
    sku_summary = (
        data.groupby("SKU", as_index=False)
        .agg(
            SKU_Order_Frequency=("Order_ID", "nunique"),
            SKU_Line_Count=("Order_ID", "count"),
            SKU_Total_Shipping_Cost_ZAR=("Billed_Cost_ZAR", "sum"),
            SKU_Distant_Zone_Lines=("Is_Distant_Zone", "sum"),
            SKU_Total_Leakage_ZAR=("Estimated_Loss_ZAR", "sum"),
        )
        .sort_values(["SKU_Order_Frequency", "SKU_Line_Count"], ascending=False)
        .reset_index(drop=True)
    )
    sku_count = len(sku_summary)
    if sku_count == 0:
        sku_summary["Velocity_Class"] = pd.Series(dtype=str)
        return sku_summary

    a_count = max(1, math.ceil(sku_count * 0.20))
    c_count = max(1, math.ceil(sku_count * 0.30))
    c_start = max(a_count, sku_count - c_count)
    sku_summary["Velocity_Class"] = "B"
    sku_summary.loc[: a_count - 1, "Velocity_Class"] = "A"
    sku_summary.loc[c_start:, "Velocity_Class"] = "C"
    return sku_summary


def run_triple_pillar_engine(raw_data, packaging_matrix, penalty_rate, volumetric_divisor, negotiated_divisor):
    data = clean_and_standardize_data(raw_data)
    cleaned_packaging_matrix = clean_packaging_matrix(packaging_matrix)
    valid = data[~data["Data_Quality_Issue"]].copy()
    skipped = data[data["Data_Quality_Issue"]].copy()

    if valid.empty:
        return valid, skipped, pd.DataFrame()

    valid["Order_ID"] = valid["Order_ID"].apply(lambda value: safe_text(value, "Unknown Order"))
    valid["SKU"] = valid["SKU"].apply(lambda value: safe_text(value, "Unknown SKU"))
    valid["Province"] = valid["Province"].apply(lambda value: safe_text(value, "Unknown"))

    valid["Physical_Volume_cm3"] = valid["Length_cm"] * valid["Width_cm"] * valid["Height_cm"]
    valid["Calculated_Vol_KG"] = valid["Physical_Volume_cm3"] / max(volumetric_divisor, 1)
    valid["Billed_Weight_KG"] = valid[["Actual_Weight_KG", "Billed_Vol_KG"]].max(axis=1)
    valid["Excess_Weight_KG"] = (valid["Billed_Weight_KG"] - valid["Actual_Weight_KG"]).clip(lower=0).fillna(0)
    valid["Estimated_Loss_ZAR"] = valid["Excess_Weight_KG"] * penalty_rate

    order_line_counts = valid.groupby("Order_ID")["SKU"].transform("count")
    valid["Is_Multi_Item"] = order_line_counts > 1
    valid["Is_Distant_Zone"] = valid["Province"].str.contains(DISTANT_ZONE_PATTERN, na=False)

    ratio = valid["Billed_Vol_KG"] / valid["Actual_Weight_KG"].replace(0, pd.NA)
    impossible_dimensions = (
        valid[["Length_cm", "Width_cm", "Height_cm"]].le(0).any(axis=1)
        | valid[["Length_cm", "Width_cm", "Height_cm"]].gt(200).any(axis=1)
        | ((valid["Physical_Volume_cm3"] / valid["Actual_Weight_KG"].replace(0, pd.NA)) > 40000)
    ).fillna(False)
    extreme_billing = (ratio > 5).fillna(False)
    valid["Anomaly_Flag"] = extreme_billing | impossible_dimensions
    valid["Anomaly_Reason"] = ""
    valid.loc[extreme_billing, "Anomaly_Reason"] = "Billed volumetric weight exceeds 500% of actual weight."
    valid.loc[impossible_dimensions, "Anomaly_Reason"] = valid.loc[impossible_dimensions, "Anomaly_Reason"].mask(
        valid.loc[impossible_dimensions, "Anomaly_Reason"].eq(""),
        "Physically impossible or likely mis-keyed dimensions.",
    )
    valid.loc[extreme_billing & impossible_dimensions, "Anomaly_Reason"] = "Extreme billed weight and impossible dimensions."
    valid["Recoverable_Overcharge_ZAR"] = valid["Estimated_Loss_ZAR"].where(valid["Anomaly_Flag"], 0).fillna(0)

    valid["Packaging_Flag"] = False
    valid["Packaging_Reason"] = "No packaging leak detected."
    valid["Recommended_Package"] = "No recommendation"
    valid["Optimized_Vol_KG"] = pd.NA
    valid["Avoidable_Volumetric_Leak_ZAR"] = 0.0

    single_mask = ~valid["Is_Multi_Item"]
    if single_mask.any():
        try:
            single_results = valid.loc[single_mask].apply(
                lambda row: run_single_item_repack(row, cleaned_packaging_matrix, penalty_rate, negotiated_divisor), axis=1
            )
            for idx, result in single_results.items():
                for key, value in result.items():
                    valid.at[idx, key] = value
        except Exception as exc:
            st.error(f"Packaging optimization failed before completion: {exc}")
            valid.loc[single_mask, "Packaging_Reason"] = f"Packaging optimization failed: {exc}"
            valid.loc[single_mask, "Recommended_Package"] = "Review manually"

    multi_orders = (
        valid[valid["Is_Multi_Item"]]
        .groupby("Order_ID", as_index=False)
        .agg(
            Combined_Physical_Volume_cm3=("Physical_Volume_cm3", "sum"),
            Combined_Actual_Weight_KG=("Actual_Weight_KG", "sum"),
            Billed_Vol_KG=("Billed_Vol_KG", "max"),
            Line_Count=("SKU", "count"),
        )
    )
    if not multi_orders.empty:
        multi_orders["Allowed_Vol_KG_With_20pct_Buffer"] = (multi_orders["Combined_Physical_Volume_cm3"] * 1.20) / max(volumetric_divisor, 1)
        multi_orders["Multi_Item_Excess_Vol_KG"] = (multi_orders["Billed_Vol_KG"] - multi_orders["Allowed_Vol_KG_With_20pct_Buffer"]).clip(lower=0)
        multi_orders["Multi_Item_Leak_ZAR"] = multi_orders["Multi_Item_Excess_Vol_KG"] * penalty_rate
        multi_lookup = multi_orders.set_index("Order_ID").to_dict("index")
        for order_id, info in multi_lookup.items():
            order_mask = valid["Order_ID"].eq(order_id)
            if info["Multi_Item_Leak_ZAR"] > 0:
                per_line_leak = info["Multi_Item_Leak_ZAR"] / max(info["Line_Count"], 1)
                valid.loc[order_mask, "Packaging_Flag"] = True
                valid.loc[order_mask, "Packaging_Reason"] = "Multi-Item Packaging Bloat: billed volumetric weight exceeds combined item volume plus 20% void-fill allowance."
                valid.loc[order_mask, "Recommended_Package"] = "Review pack station: mixed-basket carton selection"
                valid.loc[order_mask, "Optimized_Vol_KG"] = info["Allowed_Vol_KG_With_20pct_Buffer"]
                valid.loc[order_mask, "Avoidable_Volumetric_Leak_ZAR"] = per_line_leak

    sku_summary = classify_velocity(valid)
    valid = valid.merge(sku_summary, on="SKU", how="left")
    high_shipping_threshold = sku_summary["SKU_Total_Shipping_Cost_ZAR"].median() if not sku_summary.empty else 0
    valid["Capital_Trap_Flag"] = (
        valid["Velocity_Class"].eq("C")
        & (valid["SKU_Total_Shipping_Cost_ZAR"] >= high_shipping_threshold)
        & (valid["SKU_Distant_Zone_Lines"] > 0)
    ).fillna(False)

    return valid, skipped, sku_summary


def extract_pdf_rows(pdf_source):
    rows = []
    with pdfplumber.open(pdf_source) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table or len(table) < 2:
                    continue
                headers = [str(header or "").strip() for header in table[0]]
                for row in table[1:]:
                    if row and any(row):
                        rows.append({headers[i]: row[i] for i in range(min(len(headers), len(row)))})
            text = page.extract_text() or ""
            for line in text.splitlines():
                dims = parse_dimensions(line)
                order_match = re.search(r"\b[A-Z]{1,4}[- ]?\d{4,}|\b\d{8,}\b", line)
                numbers = re.findall(r"\d+(?:\.\d+)?", line)
                if order_match and dims and len(numbers) >= 5:
                    rows.append({"Order_ID": order_match.group(0), "Dimensions": " x ".join(map(str, dims)), "Actual_Weight_KG": numbers[-3], "Billed_Cost_ZAR": numbers[-1]})
    if not rows:
        raise ValueError("No invoice rows found")
    return pd.DataFrame(rows)


def extract_html_rows(file_source):
    try:
        tables = pd.read_html(file_source)
    except ValueError as exc:
        raise ValueError("No HTML tables found in the uploaded file.") from exc
    except Exception as exc:
        raise ValueError(f"Could not parse uploaded HTML file: {exc}") from exc

    if not tables:
        raise ValueError("No HTML tables found in the uploaded file.")

    debug_rows = []
    for index, table in enumerate(tables):
        score, matched_columns = score_table_for_display(table)
        debug_rows.append(f"table {index}: score={score}, shape={table.shape}, matched={matched_columns}")
    st.caption("HTML table scan: " + " | ".join(debug_rows[:6]))
    if len(debug_rows) > 6:
        st.caption(f"HTML table scan found {len(debug_rows)} tables; showing first 6 in the caption.")

    selected_table = select_shipment_table(tables)
    cleaned = clean_selected_table(selected_table)
    if cleaned.empty:
        raise ValueError("HTML file was parsed, but the selected shipment table produced no usable rows.")
    return cleaned


def score_table_for_display(table):
    try:
        from aramex_html_processor import score_table

        return score_table(table)
    except Exception as exc:
        return 0, [f"score failed: {exc}"]


@st.cache_data(show_spinner="Loading courier shipment data...")
def load_data(file_name, file_bytes):
    if file_bytes is None:
        return pd.read_csv("mock_shipping_data.csv", sep=None, engine="python", encoding="utf-8-sig")
    file_buffer = io.BytesIO(file_bytes)
    lower_name = file_name.lower()
    if lower_name.endswith(".csv"):
        return pd.read_csv(file_buffer, sep=None, engine="python", encoding="utf-8-sig")
    if lower_name.endswith(".xlsx"):
        return pd.read_excel(file_buffer)
    if lower_name.endswith(".pdf"):
        return extract_pdf_rows(file_buffer)
    if lower_name.endswith((".html", ".htm")):
        return extract_html_rows(file_buffer)
    raise ValueError("Unsupported file type. Upload CSV, XLSX, PDF, HTML, or HTM.")


class MarginPDF(FPDF):
    NAVY = (31, 41, 55)
    TEXT = (38, 38, 38)
    MUTED = (90, 99, 110)
    BORDER = (215, 221, 229)
    ZEBRA = (246, 248, 250)
    WHITE = (255, 255, 255)

    def header(self):
        self.set_y(10)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*self.NAVY)
        self.cell(0, 6, "Margin Diagnostic & Recovery Blueprint", ln=1, align="R")
        self.set_draw_color(*self.BORDER)
        self.line(self.l_margin, 20, self.w - self.r_margin, 20)
        self.ln(8)

    def footer(self):
        self.set_y(-14)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*self.MUTED)
        self.cell(95, 6, "Prepared by Kyle | Logistics Audit Advisory", align="L")
        self.cell(0, 6, f"Page {self.page_no()} of {{nb}}", align="R")


def generate_dispute_pack_xlsx(dispute_rows, currency_symbol, currency_prefix):
    """Generate a client-ready Excel dispute pack with only flagged courier overcharges."""
    overcharge_label = f"Overcharge Amount ({currency_symbol})"
    export_columns = {
        "Order_ID": "Waybill Number",
        "SKU": "Item SKU",
        "Actual_Weight_KG": "Actual Weight (kg)",
        "Billed_Vol_KG": "Billed Volumetric (kg)",
        "Recoverable_Overcharge_ZAR": overcharge_label,
        "Anomaly_Reason": "Dispute Reason",
    }
    clean_columns = list(export_columns.values())
    available_columns = [col for col in export_columns if col in dispute_rows.columns]

    dispute_pack = dispute_rows[available_columns].copy() if available_columns else pd.DataFrame()
    dispute_pack = dispute_pack.rename(columns=export_columns)
    for clean_col in clean_columns:
        if clean_col not in dispute_pack.columns:
            dispute_pack[clean_col] = ""
    dispute_pack = dispute_pack[clean_columns]

    for col in ["Actual Weight (kg)", "Billed Volumetric (kg)", overcharge_label]:
        dispute_pack[col] = pd.to_numeric(dispute_pack[col], errors="coerce").fillna(0)

    total_recoverable = dispute_pack[overcharge_label].sum()
    totals_row = {
        "Waybill Number": "",
        "Item SKU": "",
        "Actual Weight (kg)": "",
        "Billed Volumetric (kg)": "TOTAL RECOVERABLE:",
        overcharge_label: total_recoverable,
        "Dispute Reason": "",
    }
    dispute_pack = pd.concat([dispute_pack, pd.DataFrame([totals_row])], ignore_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = "Dispute Pack"
        dispute_pack.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format({"bold": True, "bg_color": "#E5E7EB", "font_color": "#111827", "border": 1, "align": "center", "valign": "vcenter"})
        text_format = workbook.add_format({"border": 1, "valign": "top"})
        weight_format = workbook.add_format({"num_format": "0.00", "border": 1, "valign": "top"})
        currency_format = workbook.add_format({"num_format": f"{currency_prefix} #,##0.00", "border": 1, "valign": "top"})
        total_label_format = workbook.add_format({"bold": True, "border": 1, "align": "right", "valign": "top", "bg_color": "#F3F4F6"})
        total_currency_format = workbook.add_format({"bold": True, "num_format": f"{currency_prefix} #,##0.00", "border": 1, "valign": "top", "bg_color": "#F3F4F6"})

        for col_idx, column_name in enumerate(dispute_pack.columns):
            worksheet.write(0, col_idx, column_name, header_format)

        actual_weight_col = dispute_pack.columns.get_loc("Actual Weight (kg)")
        billed_vol_col = dispute_pack.columns.get_loc("Billed Volumetric (kg)")
        overcharge_col = dispute_pack.columns.get_loc(overcharge_label)
        last_excel_row = len(dispute_pack)

        for row_idx in range(1, last_excel_row + 1):
            is_total_row = row_idx == last_excel_row
            for col_idx, column_name in enumerate(dispute_pack.columns):
                value = dispute_pack.iloc[row_idx - 1, col_idx]
                if is_total_row and col_idx == billed_vol_col:
                    worksheet.write(row_idx, col_idx, value, total_label_format)
                elif is_total_row and col_idx == overcharge_col:
                    worksheet.write_number(row_idx, col_idx, safe_float(value), total_currency_format)
                elif col_idx in [actual_weight_col, billed_vol_col] and not is_total_row:
                    worksheet.write_number(row_idx, col_idx, safe_float(value), weight_format)
                elif col_idx == overcharge_col and not is_total_row:
                    worksheet.write_number(row_idx, col_idx, safe_float(value), currency_format)
                else:
                    worksheet.write(row_idx, col_idx, value, text_format)

        for col_idx, column_name in enumerate(dispute_pack.columns):
            column_values = dispute_pack[column_name].astype(str).tolist()
            max_content_width = max([len(str(column_name))] + [len(value) for value in column_values])
            if column_name == "Dispute Reason":
                width = min(max(max_content_width + 4, 45), 85)
            else:
                width = min(max(max_content_width + 4, 14), 32)
            worksheet.set_column(col_idx, col_idx, width)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, last_excel_row, len(dispute_pack.columns) - 1)

    output.seek(0)
    return output.getvalue()


def generate_margin_pdf(data, anomaly_rows, packaging_rows, capital_trap_skus, courier_name, divisor, pillar_a_loss, pillar_b_loss, currency_prefix):
    pdf = MarginPDF()
    pdf.alias_nb_pages()
    pdf.set_margins(15, 24, 15)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    def ensure_space(height):
        if pdf.get_y() + height > pdf.page_break_trigger:
            pdf.add_page()

    def money(value):
        return format_currency(value, currency_prefix)

    def truncate(value, limit):
        value = "" if pd.isna(value) else str(value)
        return value if len(value) <= limit else value[: limit - 3] + "..."

    def h1(text):
        ensure_space(18)
        pdf.set_font("Helvetica", "B", 18)
        pdf.set_text_color(*MarginPDF.NAVY)
        pdf.multi_cell(0, 9, text, align="C")
        pdf.ln(2)

    def h2(text):
        ensure_space(15)
        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(*MarginPDF.NAVY)
        pdf.multi_cell(0, 7, text)
        pdf.set_draw_color(*MarginPDF.BORDER)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
        pdf.ln(3)

    def body(text):
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*MarginPDF.TEXT)
        pdf.multi_cell(0, 6, str(text))
        pdf.ln(1.5)

    def table(headers, rows, widths, aligns=None, limits=None):
        aligns = aligns or ["L"] * len(headers)
        limits = limits or [24] * len(headers)
        row_h = 7
        ensure_space(row_h * 2)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(*MarginPDF.NAVY)
        pdf.set_text_color(*MarginPDF.WHITE)
        pdf.set_draw_color(*MarginPDF.NAVY)
        for header, width, align in zip(headers, widths, aligns):
            pdf.cell(width, row_h, truncate(header, 25), border=1, align=align, fill=True)
        pdf.ln(row_h)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_draw_color(*MarginPDF.BORDER)
        if not rows:
            rows = [["No records flagged"] + [""] * (len(headers) - 1)]
        for idx, row in enumerate(rows):
            ensure_space(row_h)
            pdf.set_fill_color(*(MarginPDF.ZEBRA if idx % 2 else MarginPDF.WHITE))
            pdf.set_text_color(*MarginPDF.TEXT)
            for value, width, align, limit in zip(row, widths, aligns, limits):
                pdf.cell(width, row_h, truncate(value, limit), border=1, align=align, fill=True)
            pdf.ln(row_h)
        pdf.ln(4)

    h1("360° Margin Diagnostic & Recovery Blueprint")
    body(f"Courier configuration: {courier_name}. Volumetric divisor used: {divisor}.")
    body(f"Total direct exposure identified: {money(pillar_a_loss + pillar_b_loss)}. Courier-dispute exposure: {money(pillar_a_loss)}. Packaging workflow exposure: {money(pillar_b_loss)}.")

    h2("Chapter 1: Courier Billing Anomalies")
    body("These are the exact orders to query with the courier. They show extreme volumetric billing or physically impossible dimensions.")
    table(
        ["Order ID", "SKU", "Actual kg", "Billed Vol kg", "Reason", "Recoverable"],
        [[r["Order_ID"], r["SKU"], f"{safe_float(r['Actual_Weight_KG']):.2f}", f"{safe_float(r['Billed_Vol_KG']):.2f}", r["Anomaly_Reason"], money(r["Recoverable_Overcharge_ZAR"])] for _, r in anomaly_rows.head(30).iterrows()],
        [25, 38, 20, 23, 52, 25],
        ["L", "L", "R", "R", "L", "R"],
        [18, 26, 10, 12, 34, 14],
    )

    h2("Chapter 2: Warehouse Packaging Inefficiencies")
    body("Single-item leaks show flyer-fit opportunities. Mixed-basket leaks show orders where the billed volumetric weight exceeds combined product volume plus a 20% void-fill allowance.")
    table(
        ["Order", "SKU", "Basket", "Reason", "Recommended", "Avoidable"],
        [[r["Order_ID"], r["SKU"], "Multi" if r["Is_Multi_Item"] else "Single", r["Packaging_Reason"], r["Recommended_Package"], money(r["Avoidable_Volumetric_Leak_ZAR"])] for _, r in packaging_rows.head(30).iterrows()],
        [23, 34, 16, 52, 38, 24],
        ["L", "L", "L", "L", "L", "R"],
        [16, 24, 8, 34, 26, 14],
    )

    h2("Chapter 3: ABC Inventory Intelligence")
    body("These C-Class SKUs have low order velocity and high shipping-cost exposure to distant zones. Treat them as capital traps until margin and reorder policy are reviewed.")
    table(
        ["SKU", "Class", "Orders", "Shipping Cost", "Distant Lines", "Action"],
        [[r["SKU"], r["Velocity_Class"], f"{int(safe_float(r['SKU_Order_Frequency'])):,}", money(r["SKU_Total_Shipping_Cost_ZAR"]), f"{int(safe_float(r['SKU_Distant_Zone_Lines'])):,}", "Bundle / liquidate / reduce MOQ"] for _, r in capital_trap_skus.head(30).iterrows()],
        [40, 18, 20, 28, 25, 48],
        ["L", "L", "R", "R", "R", "L"],
        [28, 8, 10, 14, 12, 30],
    )

    h2("Methodology")
    body("All calculations are deterministic. Dirty numeric fields are coerced safely; rows with missing critical data are skipped from pillar math and shown in the data-quality table. Multi-item logic uses combined physical item volume plus a 20% allowance for void fill/bubble wrap before flagging packaging bloat.")
    return bytes(pdf.output())


st.markdown(
    """
    <style>
    .stApp {background: #ffffff; color: #1f2937;}
    [data-testid="stSidebar"] {background: #f8fafc; border-right: 1px solid #e5e7eb;}
    h1, h2, h3, p, label, span {color: #1f2937;}
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #d1d5db;
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
    }
    div[data-testid="stMetricLabel"] p {color: #374151; font-weight: 700;}
    div[data-testid="stMetricValue"] {color: #111827;}
    .info-card {background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 14px; padding: 18px; color: #1f2937;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("360° Margin Diagnostic Engine")
st.caption("A light-mode diagnostic dashboard for South African and Botswana e-commerce brands using Bob Go, The Courier Guy, and messy courier exports.")

selected_region = st.sidebar.selectbox("Select Region", ["South Africa", "Botswana"])
REGION_CONFIG = {
    "South Africa": {"currency_symbol": "ZAR", "currency_prefix": "R", "avg_shipping_cost": 120},
    "Botswana": {"currency_symbol": "BWP", "currency_prefix": "P", "avg_shipping_cost": 80},
}
region_config = REGION_CONFIG[selected_region]
currency_symbol = region_config["currency_symbol"]
currency_prefix = region_config["currency_prefix"]
avg_shipping_cost = region_config["avg_shipping_cost"]
st.sidebar.caption(f"Currency: {currency_symbol} | Average shipping cost: {format_currency(avg_shipping_cost, currency_prefix)}")
st.sidebar.divider()

supabase = get_supabase_client()
if supabase is None:
    st.sidebar.info("Supabase is not configured. CRM save/history is disabled; local analysis still works.")

st.sidebar.header("Courier Rate Card Configuration")
courier_provider = st.sidebar.selectbox(
    "Courier Provider",
    ["The Courier Guy (Divisor: 5000)", "Bob Go Aggregated (Divisor: 4000)", "Aramex (Divisor: 5000)", "Custom"],
    help="Select the courier rate-card family. This sets the default volumetric divisor but you can still override it below.",
)
provider_divisors = {
    "The Courier Guy (Divisor: 5000)": 5000,
    "Bob Go Aggregated (Divisor: 4000)": 4000,
    "Aramex (Divisor: 5000)": 5000,
    "Custom": 5000,
}
volumetric_divisor = st.sidebar.number_input(
    "Volumetric Divisor",
    min_value=1000,
    max_value=10000,
    value=provider_divisors[courier_provider],
    step=100,
    help="Courier formula divisor: Length × Width × Height ÷ Divisor = volumetric kilograms. Keep this aligned with the client's actual rate card.",
)
excess_penalty_per_kg = st.sidebar.number_input(
    f"Average Excess Penalty per Kg ({currency_symbol})",
    min_value=0.0,
    value=15.0,
    step=0.5,
    format="%.2f",
    help=f"The estimated {currency_symbol} cost per excess billed kilogram. Used to convert volumetric leakage into {currency_symbol} exposure.",
)
negotiated_divisor = st.sidebar.slider(
    "Packaging Simulation Divisor",
    3000,
    7000,
    int(volumetric_divisor),
    step=100,
    help="Used in the packaging simulation to test negotiated courier terms without removing your custom divisor logic.",
)
st.sidebar.divider()
uploaded_file = st.sidebar.file_uploader(
    "Upload courier shipment CSV, Excel, PDF, or HTML",
    type=["csv", "xlsx", "pdf", "html", "htm"],
    help="Upload messy Bob Go, The Courier Guy, or Aramex HTML exports. The auto-mapper will normalize chaotic column names automatically.",
)

selected_client_id = None
selected_client_name = "Local Client"
audit_month = "June"
audit_year = 2026
if supabase is not None:
    st.sidebar.divider()
    st.sidebar.header("Cloud CRM")
    clients = supabase.table("clients").select("id, client_name").order("client_name").execute().data or []
    client_options = {client["client_name"]: client["id"] for client in clients}
    selected_client_name = st.sidebar.selectbox("Select Client", list(client_options.keys()) if client_options else ["No clients available"], help="Select the client account for saving this audit snapshot.")
    selected_client_id = client_options.get(selected_client_name)
    audit_month = st.sidebar.text_input("Audit Month", value="June", help="Month label used when saving this audit to the CRM.")
    audit_year = st.sidebar.number_input("Audit Year", min_value=2020, max_value=2100, value=2026, step=1, help="Year label used when saving this audit to the CRM.")

try:
    file_name = uploaded_file.name if uploaded_file else None
    file_bytes = uploaded_file.getvalue() if uploaded_file else None
    raw_data = load_data(file_name, file_bytes)
    st.toast(f"Upload/load checkpoint complete: {len(raw_data):,} rows loaded.")
    st.write(f"✅ Upload/load checkpoint complete: {len(raw_data):,} rows loaded from {file_name or 'mock data'}.")
except Exception as exc:
    st.error(f"The courier file could not be loaded: {exc}")
    st.stop()

with st.expander("Packaging Matrix Configuration", expanded=False):
    packaging_matrix = st.data_editor(
        DEFAULT_PACKAGING_MATRIX,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Package": st.column_config.TextColumn("Package", help="Name of the packaging option used in the repack simulation."),
            "L": st.column_config.NumberColumn("Length (cm)", min_value=0.1, help="Outer package length in centimeters."),
            "W": st.column_config.NumberColumn("Width (cm)", min_value=0.1, help="Outer package width in centimeters."),
            "H": st.column_config.NumberColumn("Height (cm)", min_value=0.1, help="Outer package height in centimeters."),
            "cost": st.column_config.NumberColumn(f"Cost ({currency_symbol})", min_value=0.0, help="Packaging material cost subtracted from savings."),
            "Fragile/Void Fill Required": st.column_config.CheckboxColumn("Fragile/Void Fill Required", help="If checked, the model reserves 15% internal space for protection."),
        },
    )

st.toast("Packaging checkpoint: starting packaging and diagnostic calculations.")
st.write("⏳ Packaging checkpoint: starting packaging and diagnostic calculations...")
try:
    analysis_data, skipped_rows, sku_summary = run_triple_pillar_engine(
        raw_data,
        packaging_matrix,
        excess_penalty_per_kg,
        volumetric_divisor,
        negotiated_divisor,
    )
    st.toast(f"Packaging checkpoint complete: {len(analysis_data):,} valid rows analysed.")
    st.write(f"✅ Packaging checkpoint complete: {len(analysis_data):,} valid rows analysed.")
except Exception as exc:
    st.error(f"Diagnostic engine failed: {exc}")
    st.exception(exc)
    raise

if analysis_data.empty:
    st.error("No valid rows remained after dirty-data cleanup. Check that the file contains weight, billed volumetric weight, cost, and dimensions.")
    if not skipped_rows.empty:
        st.dataframe(skipped_rows, use_container_width=True, hide_index=True)
    st.stop()

anomaly_rows = analysis_data[analysis_data["Anomaly_Flag"]].sort_values("Recoverable_Overcharge_ZAR", ascending=False)
packaging_rows = analysis_data[analysis_data["Packaging_Flag"]].sort_values("Avoidable_Volumetric_Leak_ZAR", ascending=False)
capital_trap_skus = sku_summary[
    (sku_summary["Velocity_Class"].eq("C"))
    & (sku_summary["SKU_Total_Shipping_Cost_ZAR"] >= (sku_summary["SKU_Total_Shipping_Cost_ZAR"].median() if not sku_summary.empty else 0))
    & (sku_summary["SKU_Distant_Zone_Lines"] > 0)
].sort_values("SKU_Total_Shipping_Cost_ZAR", ascending=False)

pillar_a_loss = anomaly_rows["Recoverable_Overcharge_ZAR"].sum()
pillar_b_loss = packaging_rows["Avoidable_Volumetric_Leak_ZAR"].sum()
capital_trap_count = analysis_data["Capital_Trap_Flag"].sum()
multi_item_orders = analysis_data[analysis_data["Is_Multi_Item"]]["Order_ID"].nunique()

st.sidebar.divider()
consultant_password = st.sidebar.text_input(
    " ",
    type="password",
    placeholder="",
    label_visibility="collapsed",
    key="consultant_mode_password",
)

if consultant_password == "pixie":
    st.sidebar.divider()
    st.sidebar.header("Consultant Module")

    estimated_monthly_spend = (
        pd.to_numeric(analysis_data["Billed_Cost_ZAR"], errors="coerce")
        .fillna(0)
        .sum()
    )
    direct_monthly_leakage = float(pillar_a_loss + pillar_b_loss)
    modelled_monthly_risk = estimated_monthly_spend * 0.05
    annual_margin_risk = modelled_monthly_risk * 12

    pricing_tiers = {
        "Tier 1": {
            "label": "Tier 1: Courier Control",
            "condition": f"< {format_currency_whole(50_000, currency_prefix)} monthly courier spend",
            "setup_fee": 9_500,
            "retainer_fee": 7_500,
        },
        "Tier 2": {
            "label": "Tier 2: Margin Protection",
            "condition": f"{format_currency_whole(50_000, currency_prefix)}–{format_currency_whole(150_000, currency_prefix)} monthly courier spend",
            "setup_fee": 18_000,
            "retainer_fee": 14_500,
        },
        "Tier 3": {
            "label": "Tier 3: Logistics Governance",
            "condition": f"> {format_currency_whole(150_000, currency_prefix)} monthly courier spend",
            "setup_fee": 32_000,
            "retainer_fee": 29_000,
        },
    }

    if estimated_monthly_spend < 50_000:
        calculated_tier_key = "Tier 1"
    elif estimated_monthly_spend <= 150_000:
        calculated_tier_key = "Tier 2"
    else:
        calculated_tier_key = "Tier 3"

    calculated_tier = pricing_tiers[calculated_tier_key]

    if direct_monthly_leakage < calculated_tier["retainer_fee"]:
        value_anchor = max(direct_monthly_leakage, modelled_monthly_risk)
        recommended_retainer = min(
            calculated_tier["retainer_fee"],
            max(2_500, int(round((value_anchor * 0.50) / 500) * 500)),
        )
        recommended_setup = min(
            calculated_tier["setup_fee"],
            max(4_500, int(round((direct_monthly_leakage * 1.50) / 500) * 500)),
        )
        pricing_note = "Beta / value-protected pricing recommended because observed leakage is below the list retainer."
    else:
        recommended_retainer = calculated_tier["retainer_fee"]
        recommended_setup = calculated_tier["setup_fee"]
        pricing_note = "List pricing is commercially defensible against observed leakage."

    st.sidebar.caption(f"Detected: {calculated_tier['label']}")
    st.sidebar.caption(calculated_tier["condition"])
    st.sidebar.caption(pricing_note)

    setup_fee = st.sidebar.number_input(
        "Setup Fee / Historical Audit Fee",
        min_value=0,
        value=recommended_setup,
        step=500,
        format="%d",
        key="consultant_setup_fee",
        help="Editable for manual overrides or beta discounts.",
    )

    retainer_fee = st.sidebar.number_input(
        "Monthly Retainer Fee",
        min_value=0,
        value=recommended_retainer,
        step=500,
        format="%d",
        key="consultant_retainer_fee",
        help="Editable for manual overrides or beta discounts.",
    )

    annual_retainer = retainer_fee * 12

    st.sidebar.subheader("Pricing Dashboard")
    col1, col2 = st.sidebar.columns(2)
    col1.metric("Setup", format_currency_whole(setup_fee, currency_prefix))
    col2.metric("Retainer", f"{format_currency_whole(retainer_fee, currency_prefix)}/mo")

    st.sidebar.metric("Estimated Monthly Courier Spend", format_currency_whole(estimated_monthly_spend, currency_prefix))
    st.sidebar.metric("Observed Direct Leakage", format_currency_whole(direct_monthly_leakage, currency_prefix))
    st.sidebar.metric(
        "Annual Margin Risk",
        format_currency_whole(annual_margin_risk, currency_prefix),
        help="Calculated as 5% of estimated monthly courier spend, annualised.",
    )
    st.sidebar.info(
        f"List tier: {format_currency_whole(calculated_tier['retainer_fee'], currency_prefix)}/mo. "
        f"Your proposed retainer: {format_currency_whole(annual_retainer, currency_prefix)}/year. "
        f"Cost of internal hire: ~{format_currency_whole(300_000, currency_prefix)}/year."
    )

    if retainer_fee > direct_monthly_leakage and direct_monthly_leakage > 0:
        st.sidebar.warning(
            "Retainer is above the observed direct leakage. Pitch this as governance, prevention, and beta monitoring — not as immediate monthly savings."
        )


k1, k2, k3, k4 = st.columns(4)
k1.metric("Pillar A: Courier Recoverable", format_currency(pillar_a_loss, currency_prefix), f"{len(anomaly_rows):,} rows", help="Money potentially recoverable from courier billing anomalies such as extreme billed volumetric weight or impossible dimensions.")
k2.metric("Pillar B: Packaging Leak", format_currency(pillar_b_loss, currency_prefix), f"{len(packaging_rows):,} rows", help="Avoidable warehouse-side leakage from single-item flyer opportunities and multi-item packaging bloat.")
k3.metric("Pillar C: Capital Traps", f"{int(capital_trap_count):,}", "flagged lines", help="C-Class low-velocity SKUs that also incur high shipping costs to distant or remote zones.")
k4.metric("Mixed-Basket Orders", f"{multi_item_orders:,}", "multi-item", help="Calculates the combined physical volume of all items in a single order plus a 20% buffer for packaging material.")

summary_tab, anomaly_tab, packaging_tab, velocity_tab = st.tabs(["Executive Summary", "Courier Anomalies", "Packaging Leaks", "SKU Velocity"])

with summary_tab:
    st.subheader("Executive Summary")
    st.markdown(
        f"""
        <div class="info-card">
        The engine standardized <b>{len(raw_data):,}</b> raw courier rows and retained <b>{len(analysis_data):,}</b> rows for deterministic pillar analysis.
        Total direct leakage identified is <b>{format_currency(pillar_a_loss + pillar_b_loss, currency_prefix)}</b>, split into courier-dispute exposure and warehouse packaging exposure.
        </div>
        """,
        unsafe_allow_html=True,
    )
    leakage_column = f"Leakage_{currency_symbol}"
    chart_data = pd.DataFrame([
        {"Pillar": "A: Courier Billing Anomalies", leakage_column: pillar_a_loss},
        {"Pillar": "B: Packaging Inefficiency", leakage_column: pillar_b_loss},
    ])
    st.plotly_chart(px.bar(chart_data, x="Pillar", y=leakage_column, text_auto=".2s", color="Pillar", color_discrete_sequence=["#2563eb", "#16a34a"], template="plotly_white"), use_container_width=True)

    pdf_report = generate_margin_pdf(analysis_data, anomaly_rows, packaging_rows, capital_trap_skus, courier_provider, volumetric_divisor, pillar_a_loss, pillar_b_loss, currency_prefix)
    c1, c2, c3 = st.columns(3)
    c1.download_button("Download PDF Blueprint", pdf_report, "Margin_Diagnostic_Recovery_Blueprint.pdf", "application/pdf", use_container_width=True, help="Downloads the white-background PDF report with three pillar chapters and zebra-striped tables.")
    c2.download_button("Download Full Diagnostic CSV", analysis_data.to_csv(index=False).encode("utf-8"), "full_margin_diagnostic.csv", "text/csv", use_container_width=True, help="Exports every cleaned and calculated row with all pillar flags.")
    if supabase is not None:
        if c3.button("Save Audit to CRM", disabled=selected_client_id is None, use_container_width=True, help="Saves the top-line diagnostic values to Supabase for client history tracking."):
            supabase.table("audits").insert({"client_id": selected_client_id, "audit_month": audit_month, "audit_year": int(audit_year), "total_loss_zar": float(pillar_a_loss + pillar_b_loss), "dead_stock_value_zar": float(capital_trap_skus["SKU_Total_Shipping_Cost_ZAR"].sum() if not capital_trap_skus.empty else 0)}).execute()
            st.success(f"Saved {audit_month} {audit_year} audit for {selected_client_name}.")

    with st.expander("Skipped Dirty-Data Rows"):
        st.caption("Rows shown here were not used in pillar calculations because critical numeric data was missing or invalid.")
        st.dataframe(display_currency_columns(skipped_rows, currency_symbol), use_container_width=True, hide_index=True)
    st.subheader("Full Cleaned Diagnostic Data")
    st.dataframe(display_currency_columns(analysis_data, currency_symbol), use_container_width=True, hide_index=True)

with anomaly_tab:
    st.subheader("Pillar A: Invoice & Billing Anomalies")
    with st.expander("ELI5: What does this mean and why should I care?", expanded=True):
        st.write("This tab finds orders where the courier bill looks suspicious. If a parcel weighs 1kg but the courier bills it like it takes up the space of 5kg or more, you may be paying an unfair 'empty air' tax. These rows are your dispute list.")
    cols = ["Order_ID", "SKU", "Province", "Actual_Weight_KG", "Billed_Vol_KG", "Anomaly_Reason", "Recoverable_Overcharge_ZAR", "Billed_Cost_ZAR"]
    st.dataframe(display_currency_columns(anomaly_rows[[col for col in cols if col in anomaly_rows.columns]], currency_symbol), use_container_width=True, hide_index=True)
    st.download_button(
        "Export Dispute Pack",
        generate_dispute_pack_xlsx(anomaly_rows, currency_symbol, currency_prefix),
        "courier_dispute_pack.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Downloads a professional Excel dispute pack containing only flagged courier overcharges.",
    )

with packaging_tab:
    st.subheader("Pillar B: Workflow & Packaging Inefficiency")
    with st.expander("ELI5: What does this mean and why should I care?", expanded=True):
        st.write("This tab separates single-item mistakes from mixed-basket mistakes. A single t-shirt in a big box should become a flyer rule. For multi-item orders, we do not guess a box; we check whether the courier's billed size is bigger than all products combined plus a fair 20% packing buffer.")
    cols = ["Order_ID", "SKU", "Is_Multi_Item", "Packaging_Reason", "Recommended_Package", "Billed_Vol_KG", "Optimized_Vol_KG", "Avoidable_Volumetric_Leak_ZAR"]
    st.dataframe(display_currency_columns(packaging_rows[[col for col in cols if col in packaging_rows.columns]], currency_symbol), use_container_width=True, hide_index=True)
    if not packaging_rows.empty:
        avoidable_column = f"Avoidable_Volumetric_Leak_{currency_symbol}"
        by_reason = display_currency_columns(packaging_rows, currency_symbol).groupby("Packaging_Reason", as_index=False)[avoidable_column].sum()
        st.plotly_chart(px.pie(by_reason, names="Packaging_Reason", values=avoidable_column, color_discrete_sequence=px.colors.qualitative.Safe, template="plotly_white"), use_container_width=True)
    st.download_button("Download Packaging Leak CSV", packaging_rows.to_csv(index=False).encode("utf-8"), "packaging_leak_rows.csv", "text/csv", help="Exports single-item flyer opportunities and multi-item packaging bloat rows.")

if consultant_password == "pixie":
    st.divider()
    st.subheader("Automated Pitch Engine")
    pitch_company_name = st.text_input(
        "Company Name",
        value=selected_client_name if selected_client_name != "Local Client" else "",
        key="pitch_company_name",
    )
    pitch_founder_name = st.text_input(
        "Founder Name",
        key="pitch_founder_name",
    )
    pitch_brand_context = st.text_area(
        "Brand Context",
        placeholder="e.g., Premium athleisure brand dominating the Cape Town market with sustainable materials...",
        key="pitch_brand_context",
    )

    worst_sku = "your highest-risk SKU"
    if not capital_trap_skus.empty and "SKU" in capital_trap_skus.columns:
        worst_sku = safe_text(capital_trap_skus.iloc[0]["SKU"], worst_sku)
    elif not packaging_rows.empty and "SKU" in packaging_rows.columns:
        worst_sku = safe_text(packaging_rows.iloc[0]["SKU"], worst_sku)

    total_recoverable = safe_float(pillar_a_loss)
    total_avoidable = safe_float(pillar_b_loss)
    annualized_loss = (total_recoverable + total_avoidable) * 12
    company_for_pitch = pitch_company_name.strip() or "your company"
    founder_for_pitch = pitch_founder_name.strip() or "there"
    context_for_pitch = pitch_brand_context.strip() or "the brand positioning and operational momentum you are building"

    pitch_text = f"""Hi {founder_for_pitch},

I just finished running the historical logistics audit on the dataset you sent over. First off, I love what {company_for_pitch} is doing—{context_for_pitch} is a massive differentiator right now.

I wanted to get straight to the numbers. The engine flagged two major structural leaks in your current shipping setup:

1. Courier Glitches: I have isolated exactly {format_currency(total_recoverable, currency_prefix)} in pure courier overcharges from last month. I've attached the Dispute Pack—you can forward this straight to your account manager to get that cash credited back.
2. Packaging Bloat: We found that oversized packaging, specifically on orders containing the {worst_sku}, is triggering an additional {format_currency(total_avoidable, currency_prefix)} in unnecessary volumetric penalties.

If we annualize these figures, {company_for_pitch} is quietly losing over {format_currency(annualized_loss, currency_prefix)} this year just to un-optimized dispatch logistics.

I’d love to jump on a quick 15-minute Google Meet this week to walk you through the visual dashboard and show you exactly how to plug these holes.

Let me know what day works best for you."""

    st.code(pitch_text, language="text")

with velocity_tab:
    st.subheader("Pillar C: SKU Velocity & Trapped Capital")
    with st.expander("ELI5: What does this mean and why should I care?", expanded=True):
        st.write("This tab tells you which products are slow movers. If a slow product also costs a lot to ship to far-away zones, it traps cash in inventory and courier fees. These SKUs should be bundled, liquidated, or reordered less often.")
    v1, v2, v3 = st.columns(3)
    v1.metric("A-Class SKUs", f"{sku_summary[sku_summary['Velocity_Class'].eq('A')]['SKU'].nunique():,}", help="Highest-frequency SKUs. These usually deserve stock protection and fast fulfillment.")
    v2.metric("B-Class SKUs", f"{sku_summary[sku_summary['Velocity_Class'].eq('B')]['SKU'].nunique():,}", help="Middle-frequency SKUs. Monitor these for movement up or down.")
    v3.metric("C-Class SKUs", f"{sku_summary[sku_summary['Velocity_Class'].eq('C')]['SKU'].nunique():,}", help="Bottom 30% of SKUs by order frequency. These may be dead-stock or slow-moving inventory.")
    st.dataframe(display_currency_columns(sku_summary, currency_symbol), use_container_width=True, hide_index=True)
    if not sku_summary.empty:
        shipping_cost_column = f"SKU_Total_Shipping_Cost_{currency_symbol}"
        sku_chart_data = display_currency_columns(sku_summary, currency_symbol)
        st.plotly_chart(px.treemap(sku_chart_data, path=["Velocity_Class", "SKU"], values=shipping_cost_column, color="SKU_Order_Frequency", color_continuous_scale="Blues", template="plotly_white"), use_container_width=True)
    st.download_button("Download SKU Velocity CSV", sku_summary.to_csv(index=False).encode("utf-8"), "sku_velocity_matrix.csv", "text/csv", help="Exports ABC velocity classes and shipping-cost exposure by SKU.")
