import numpy as np
import pandas as pd

OUTPUT_FILE = "realistic_courier_export.csv"
N_ROWS = 500
RANDOM_SEED = 42

np.random.seed(RANDOM_SEED)

PROVINCES = [
    "Western Cape",
    "Gauteng",
    "KwaZulu-Natal",
    "Eastern Cape",
    "Free State",
    "Limpopo",
    "Mpumalanga",
    "North West",
    "Northern Cape",
]

PROVINCE_PROBS = [0.34, 0.31, 0.12, 0.06, 0.04, 0.04, 0.04, 0.03, 0.02]

ZONE_BASE_COST = {
    "Western Cape": 58,
    "Gauteng": 68,
    "KwaZulu-Natal": 76,
    "Eastern Cape": 82,
    "Free State": 84,
    "Limpopo": 92,
    "Mpumalanga": 90,
    "North West": 88,
    "Northern Cape": 98,
}

ZONE_RATE_PER_KG = {
    "Western Cape": 10,
    "Gauteng": 13,
    "KwaZulu-Natal": 15,
    "Eastern Cape": 16,
    "Free State": 17,
    "Limpopo": 18,
    "Mpumalanga": 18,
    "North West": 17,
    "Northern Cape": 20,
}

PRODUCTS = {
    "Supplements": [
        {"sku": "SUP-WHEY-1KG", "weight": 1.05, "dims": (16, 16, 24)},
        {"sku": "SUP-CREAT-300G", "weight": 0.36, "dims": (9, 9, 12)},
        {"sku": "SUP-VIT-60CAP", "weight": 0.14, "dims": (6, 6, 11)},
        {"sku": "SUP-PRE-250G", "weight": 0.32, "dims": (10, 10, 13)},
        {"sku": "SUP-COLL-400G", "weight": 0.48, "dims": (10, 10, 17)},
    ],
    "Fitness Gear": [
        {"sku": "FIT-BAND-SET", "weight": 0.42, "dims": (18, 14, 5)},
        {"sku": "FIT-GLOVES-M", "weight": 0.24, "dims": (22, 14, 4)},
        {"sku": "FIT-SKIPROPE", "weight": 0.28, "dims": (18, 12, 5)},
        {"sku": "FIT-YOGAMAT", "weight": 1.15, "dims": (62, 13, 13)},
        {"sku": "FIT-SHAKER", "weight": 0.18, "dims": (9, 9, 22)},
    ],
    "Apparel": [
        {"sku": "APP-TEE-BLK", "weight": 0.20, "dims": (28, 22, 3)},
        {"sku": "APP-TEE-WHT", "weight": 0.20, "dims": (28, 22, 3)},
        {"sku": "APP-HOODIE", "weight": 0.72, "dims": (34, 28, 7)},
        {"sku": "APP-LEGGINGS", "weight": 0.28, "dims": (30, 22, 4)},
        {"sku": "APP-CAP", "weight": 0.16, "dims": (22, 18, 10)},
    ],
}

CATEGORY_PROBS = {
    "Supplements": 0.48,
    "Fitness Gear": 0.22,
    "Apparel": 0.30,
}


def volumetric_kg(length_cm, width_cm, height_cm, divisor=5000):
    return (length_cm * width_cm * height_cm) / divisor


def chargeable_weight(actual_weight, billed_vol_weight):
    return max(actual_weight, billed_vol_weight)


def courier_cost(chargeable_kg, province):
    """Realistic simplified courier pricing: base fee includes first 2kg, then per-kg increments."""
    base = ZONE_BASE_COST[province]
    rate = ZONE_RATE_PER_KG[province]
    excess_kg = max(0, np.ceil(chargeable_kg - 2.0))
    fuel_surcharge = 0.12
    return round((base + excess_kg * rate) * (1 + fuel_surcharge), 2)


def choose_product():
    category = np.random.choice(
        list(CATEGORY_PROBS.keys()),
        p=list(CATEGORY_PROBS.values()),
    )
    product = np.random.choice(PRODUCTS[category])
    return category, product


def add_noise(value, pct=0.06, minimum=0.01):
    noisy = value * np.random.normal(1, pct)
    return max(minimum, noisy)


def make_standard_row(order_id):
    category, product = choose_product()
    province = np.random.choice(PROVINCES, p=PROVINCE_PROBS)

    actual_weight = add_noise(product["weight"], pct=0.05)
    length, width, height = product["dims"]
    length = add_noise(length, pct=0.04, minimum=3)
    width = add_noise(width, pct=0.04, minimum=3)
    height = add_noise(height, pct=0.05, minimum=1)

    natural_vol = volumetric_kg(length, width, height)
    billed_vol = max(natural_vol, actual_weight * np.random.uniform(0.88, 1.08))
    billed_vol = max(0.1, billed_vol)
    cost = courier_cost(chargeable_weight(actual_weight, billed_vol), province)

    return {
        "Order_ID": order_id,
        "SKU": product["sku"],
        "Product_Category": category,
        "Actual_Weight_KG": round(actual_weight, 2),
        "Length_cm": round(length, 1),
        "Width_cm": round(width, 1),
        "Height_cm": round(height, 1),
        "Billed_Vol_KG": round(billed_vol, 2),
        "Billed_Cost_ZAR": cost,
        "Destination_Province": province,
        "_Leak_Type": "Normal",
        "_Expected_Cost_ZAR": cost,
    }


def build_order_ids(n_rows):
    """Create 500 line items with roughly 30% of orders represented by multi-row baskets."""
    target_multi_rows = int(n_rows * 0.30)
    order_ids = []
    order_counter = 10001

    multi_rows_created = 0
    while multi_rows_created < target_multi_rows:
        basket_size = np.random.choice([2, 3], p=[0.78, 0.22])
        if multi_rows_created + basket_size > target_multi_rows:
            basket_size = target_multi_rows - multi_rows_created
        order_id = f"LA-{order_counter}"
        order_ids.extend([order_id] * basket_size)
        order_counter += 1
        multi_rows_created += basket_size

    while len(order_ids) < n_rows:
        order_ids.append(f"LA-{order_counter}")
        order_counter += 1

    np.random.shuffle(order_ids)
    return order_ids


def inject_subtle_billing_anomalies(df, n=13):
    """Inject decimal/scanner glitches capped below R250 estimated overcharge per line."""
    candidate_idx = df.index[df["_Leak_Type"].eq("Normal")].to_numpy()
    chosen = np.random.choice(candidate_idx, size=n, replace=False)

    for idx in chosen:
        row = df.loc[idx]
        actual = float(row["Actual_Weight_KG"])
        province = row["Destination_Province"]
        expected_cost = courier_cost(chargeable_weight(actual, max(actual, volumetric_kg(row["Length_cm"], row["Width_cm"], row["Height_cm"]))), province)

        # Common believable cases: decimal shift or scanner overread, not extreme fantasy values.
        possible_billed_weights = [
            round(actual * 10, 1),
            round(actual + np.random.uniform(2.2, 4.2), 1),
            round(np.random.choice([3.5, 4.0, 4.5, 6.0, 8.0, 10.0, 12.0]), 1),
        ]
        np.random.shuffle(possible_billed_weights)

        selected_weight = None
        selected_cost = None
        selected_overcharge = None
        for billed_weight in possible_billed_weights:
            billed_weight = min(max(billed_weight, actual + 1.2), 14.0)
            anomalous_cost = courier_cost(billed_weight, province)
            overcharge = anomalous_cost - expected_cost
            if 25 <= overcharge <= 245:
                selected_weight = billed_weight
                selected_cost = anomalous_cost
                selected_overcharge = overcharge
                break

        if selected_weight is None:
            selected_weight = min(actual + 3.0, 8.0)
            selected_cost = courier_cost(selected_weight, province)
            selected_overcharge = selected_cost - expected_cost

        df.loc[idx, "Billed_Vol_KG"] = round(selected_weight, 2)
        df.loc[idx, "Billed_Cost_ZAR"] = round(selected_cost, 2)
        df.loc[idx, "_Expected_Cost_ZAR"] = round(expected_cost, 2)
        df.loc[idx, "_Leak_Type"] = "Subtle Billing Anomaly"
        df.loc[idx, "_Injected_Leak_ZAR"] = round(max(0, selected_overcharge), 2)

    return df


def inject_packaging_inefficiencies(df, n=36):
    """Inject realistic small-box/flyer mismatch, mostly apparel and compact supplements."""
    mask = (
        df["_Leak_Type"].eq("Normal")
        & df["Product_Category"].isin(["Apparel", "Supplements"])
        & (df["Actual_Weight_KG"] <= 1.1)
    )
    candidate_idx = df.index[mask].to_numpy()
    chosen = np.random.choice(candidate_idx, size=n, replace=False)

    for idx in chosen:
        row = df.loc[idx]
        actual = float(row["Actual_Weight_KG"])
        province = row["Destination_Province"]

        expected_vol = max(actual, volumetric_kg(row["Length_cm"], row["Width_cm"], row["Height_cm"]))
        expected_cost = courier_cost(expected_vol, province)

        # Small corrugated boxes commonly used instead of flyer bags.
        box_options = [
            (24, 18, 10),  # 0.86kg volumetric
            (28, 22, 12),  # 1.48kg volumetric
            (30, 24, 14),  # 2.02kg volumetric
            (32, 25, 15),  # 2.40kg volumetric
        ]
        length, width, height = box_options[np.random.choice(len(box_options), p=[0.18, 0.42, 0.30, 0.10])]
        length = add_noise(length, pct=0.025)
        width = add_noise(width, pct=0.025)
        height = add_noise(height, pct=0.025)
        billed_vol = volumetric_kg(length, width, height)

        bloated_cost = courier_cost(chargeable_weight(actual, billed_vol), province)
        leak = bloated_cost - expected_cost

        # Keep most packaging leaks in the R15-R25 range by nudging the billed weight if needed.
        if leak < 10:
            billed_vol = max(billed_vol, 2.25)
            bloated_cost = courier_cost(chargeable_weight(actual, billed_vol), province)
            leak = bloated_cost - expected_cost
        if leak > 35:
            billed_vol = min(billed_vol, 2.1)
            bloated_cost = courier_cost(chargeable_weight(actual, billed_vol), province)
            leak = bloated_cost - expected_cost

        df.loc[idx, "Length_cm"] = round(length, 1)
        df.loc[idx, "Width_cm"] = round(width, 1)
        df.loc[idx, "Height_cm"] = round(height, 1)
        df.loc[idx, "Billed_Vol_KG"] = round(max(billed_vol, actual), 2)
        df.loc[idx, "Billed_Cost_ZAR"] = round(bloated_cost, 2)
        df.loc[idx, "_Expected_Cost_ZAR"] = round(expected_cost, 2)
        df.loc[idx, "_Leak_Type"] = "Packaging Inefficiency"
        df.loc[idx, "_Injected_Leak_ZAR"] = round(max(0, leak), 2)

    return df


def tune_total_leak(df, target_min=2500, target_max=4500):
    """Small deterministic top-up if the simplified tariff model lands below the target leak band."""
    leak = df.get("_Injected_Leak_ZAR", pd.Series(0, index=df.index)).fillna(0)
    total = float(leak.sum())

    if target_min <= total <= target_max:
        return df

    # If below target, add a few modest R35-R80 anomalies. Still cap each line below R250.
    if total < target_min:
        needed = target_min - total + 150
        candidates = df.index[df["_Leak_Type"].eq("Normal")].to_numpy().copy()
        np.random.shuffle(candidates)
        added = 0
        for idx in candidates[:20]:
            row = df.loc[idx]
            province = row["Destination_Province"]
            actual = float(row["Actual_Weight_KG"])
            expected_cost = float(row["Billed_Cost_ZAR"])
            new_weight = min(max(actual + np.random.uniform(3.0, 6.0), 4.0), 12.0)
            new_cost = courier_cost(new_weight, province)
            extra = new_cost - expected_cost
            if 30 <= extra <= 120:
                df.loc[idx, "Billed_Vol_KG"] = round(new_weight, 2)
                df.loc[idx, "Billed_Cost_ZAR"] = round(new_cost, 2)
                df.loc[idx, "_Expected_Cost_ZAR"] = round(expected_cost, 2)
                df.loc[idx, "_Leak_Type"] = "Subtle Billing Anomaly"
                df.loc[idx, "_Injected_Leak_ZAR"] = round(extra, 2)
                added += extra
            if added >= needed:
                break

    return df


def main():
    order_ids = build_order_ids(N_ROWS)
    rows = [make_standard_row(order_id) for order_id in order_ids]
    df = pd.DataFrame(rows)
    df["_Injected_Leak_ZAR"] = 0.0

    df = inject_subtle_billing_anomalies(df, n=13)
    df = inject_packaging_inefficiencies(df, n=36)
    df = tune_total_leak(df, target_min=2500, target_max=4500)

    # Exactly 5 dirty missing actual-weight cells, avoiding injected anomaly rows so leak math remains interpretable.
    clean_candidates = df.index[df["_Leak_Type"].eq("Normal")].to_numpy()
    dirty_idx = np.random.choice(clean_candidates, size=5, replace=False)
    df.loc[dirty_idx, "Actual_Weight_KG"] = np.nan
    df.loc[dirty_idx, "_Leak_Type"] = "Dirty Missing Weight"

    # Sort by order ID so multi-item baskets are easy to inspect in the app.
    df = df.sort_values(["Order_ID", "SKU"]).reset_index(drop=True)

    total_injected_leak = df["_Injected_Leak_ZAR"].sum()
    anomaly_count = (df["_Leak_Type"] == "Subtle Billing Anomaly").sum()
    packaging_count = (df["_Leak_Type"] == "Packaging Inefficiency").sum()
    dirty_count = df["Actual_Weight_KG"].isna().sum()
    multi_order_share = df.groupby("Order_ID").size().gt(1).sum() / df["Order_ID"].nunique()
    max_line_leak = df["_Injected_Leak_ZAR"].max()

    # Export only the columns expected by the Streamlit app.
    export_cols = [
        "Order_ID",
        "SKU",
        "Product_Category",
        "Actual_Weight_KG",
        "Length_cm",
        "Width_cm",
        "Height_cm",
        "Billed_Vol_KG",
        "Billed_Cost_ZAR",
        "Destination_Province",
    ]
    df[export_cols].to_csv(OUTPUT_FILE, index=False)

    print(f"Generated {OUTPUT_FILE} with {len(df)} rows")
    print(f"Subtle billing anomalies: {anomaly_count}")
    print(f"Packaging inefficiencies: {packaging_count}")
    print(f"Blank Actual_Weight_KG cells: {dirty_count}")
    print(f"Multi-row order share: {multi_order_share:.1%}")
    print(f"Injected total leak estimate: R{total_injected_leak:,.2f}")
    print(f"Maximum single-line injected leak: R{max_line_leak:,.2f}")


if __name__ == "__main__":
    main()
