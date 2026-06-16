"""Generate a realistic synthetic courier export for the Logistics Margin Diagnostic app.

Run locally from the logistics_auditor folder:
    python generate_mock_data.py

Output:
    synthetic_courier_export.csv
"""

from __future__ import annotations

import numpy as np
import pandas as pd


RNG_SEED = 42
ROW_COUNT = 500
OUTPUT_FILE = "synthetic_courier_export.csv"
VOLUMETRIC_DIVISOR = 5000
BASE_RATE_ZAR = 58.0
PER_KG_RATE_ZAR = 12.5

PROVINCES = [
    "Gauteng",
    "Western Cape",
    "KwaZulu-Natal",
    "Eastern Cape",
    "Free State",
    "Limpopo",
    "Mpumalanga",
    "North West",
    "Northern Cape",
]

PROVINCE_MULTIPLIER = {
    "Gauteng": 1.00,
    "Western Cape": 1.12,
    "KwaZulu-Natal": 1.10,
    "Eastern Cape": 1.22,
    "Free State": 1.16,
    "Limpopo": 1.25,
    "Mpumalanga": 1.18,
    "North West": 1.18,
    "Northern Cape": 1.35,
}

PRODUCT_CATALOG = [
    # Supplements: dense, moderate weights, regular box/pouch sizes.
    {"SKU": "SUP-WHEY-1KG", "Product_Category": "Supplements", "weight": 1.15, "dims": (14, 14, 22)},
    {"SKU": "SUP-WHEY-2KG", "Product_Category": "Supplements", "weight": 2.25, "dims": (18, 18, 28)},
    {"SKU": "SUP-CREATINE-300G", "Product_Category": "Supplements", "weight": 0.38, "dims": (10, 10, 12)},
    {"SKU": "SUP-PREWORKOUT", "Product_Category": "Supplements", "weight": 0.45, "dims": (11, 11, 13)},
    {"SKU": "SUP-VITAMINS", "Product_Category": "Supplements", "weight": 0.18, "dims": (6, 6, 11)},
    # Fitness gear: heavier or awkward items.
    {"SKU": "FIT-RESIST-BAND", "Product_Category": "Fitness Gear", "weight": 0.28, "dims": (16, 12, 4)},
    {"SKU": "FIT-WRIST-WRAP", "Product_Category": "Fitness Gear", "weight": 0.20, "dims": (18, 10, 4)},
    {"SKU": "FIT-YOGA-MAT", "Product_Category": "Fitness Gear", "weight": 1.05, "dims": (62, 14, 14)},
    {"SKU": "FIT-SHAKER", "Product_Category": "Fitness Gear", "weight": 0.16, "dims": (9, 9, 22)},
    {"SKU": "FIT-LIFT-GLOVES", "Product_Category": "Fitness Gear", "weight": 0.24, "dims": (18, 12, 5)},
    # Apparel: light, should usually fit flyers.
    {"SKU": "APP-TEE-BLACK", "Product_Category": "Apparel", "weight": 0.22, "dims": (28, 22, 2)},
    {"SKU": "APP-TEE-WHITE", "Product_Category": "Apparel", "weight": 0.22, "dims": (28, 22, 2)},
    {"SKU": "APP-HOODIE", "Product_Category": "Apparel", "weight": 0.65, "dims": (35, 28, 5)},
    {"SKU": "APP-LEGGINGS", "Product_Category": "Apparel", "weight": 0.30, "dims": (30, 22, 3)},
    {"SKU": "APP-CAP", "Product_Category": "Apparel", "weight": 0.18, "dims": (24, 18, 10)},
]

CATEGORY_PROBABILITIES = {
    "Supplements": 0.46,
    "Fitness Gear": 0.24,
    "Apparel": 0.30,
}


def volumetric_kg(length_cm: float, width_cm: float, height_cm: float) -> float:
    return round((length_cm * width_cm * height_cm) / VOLUMETRIC_DIVISOR, 2)


def courier_cost(billed_vol_kg: float, actual_weight_kg: float, province: str, rng: np.random.Generator) -> float:
    billable_weight = max(billed_vol_kg, actual_weight_kg)
    noise = rng.normal(0, 4.5)
    cost = (BASE_RATE_ZAR + PER_KG_RATE_ZAR * billable_weight) * PROVINCE_MULTIPLIER[province] + noise
    return round(max(cost, 45.0), 2)


def jitter(value: float, rng: np.random.Generator, pct: float = 0.08, minimum: float = 0.01) -> float:
    return round(max(value * rng.normal(1.0, pct), minimum), 2)


def choose_product(rng: np.random.Generator, category: str | None = None) -> dict:
    if category is None:
        categories = list(CATEGORY_PROBABILITIES.keys())
        probs = list(CATEGORY_PROBABILITIES.values())
        category = rng.choice(categories, p=probs)
    options = [product for product in PRODUCT_CATALOG if product["Product_Category"] == category]
    return dict(rng.choice(options))


def create_normal_line(order_id: str, rng: np.random.Generator, category: str | None = None) -> dict:
    product = choose_product(rng, category)
    base_l, base_w, base_h = product["dims"]
    length = jitter(base_l, rng, pct=0.06)
    width = jitter(base_w, rng, pct=0.06)
    height = jitter(base_h, rng, pct=0.10)
    actual_weight = jitter(product["weight"], rng, pct=0.12)
    billed_vol = max(volumetric_kg(length, width, height), round(actual_weight * rng.uniform(0.95, 1.35), 2))
    province = rng.choice(PROVINCES, p=[0.34, 0.22, 0.15, 0.07, 0.04, 0.05, 0.05, 0.04, 0.04])
    return {
        "Order_ID": order_id,
        "SKU": product["SKU"],
        "Product_Category": product["Product_Category"],
        "Actual_Weight_KG": actual_weight,
        "Length_cm": length,
        "Width_cm": width,
        "Height_cm": height,
        "Billed_Vol_KG": billed_vol,
        "Billed_Cost_ZAR": courier_cost(billed_vol, actual_weight, province, rng),
        "Destination_Province": province,
    }


def create_billing_anomaly(order_id: str, rng: np.random.Generator) -> dict:
    product = dict(rng.choice([p for p in PRODUCT_CATALOG if p["SKU"] in ["FIT-WRIST-WRAP", "APP-TEE-BLACK", "SUP-VITAMINS", "FIT-RESIST-BAND"]]))
    base_l, base_w, base_h = product["dims"]
    length = jitter(base_l, rng, pct=0.04)
    width = jitter(base_w, rng, pct=0.04)
    height = jitter(base_h, rng, pct=0.05)
    actual_weight = jitter(product["weight"], rng, pct=0.08)
    billed_vol = round(rng.uniform(55, 95), 2)
    province = rng.choice(PROVINCES)
    return {
        "Order_ID": order_id,
        "SKU": product["SKU"],
        "Product_Category": product["Product_Category"],
        "Actual_Weight_KG": actual_weight,
        "Length_cm": length,
        "Width_cm": width,
        "Height_cm": height,
        "Billed_Vol_KG": billed_vol,
        "Billed_Cost_ZAR": courier_cost(billed_vol, actual_weight, province, rng),
        "Destination_Province": province,
    }


def create_packaging_inefficiency(order_id: str, rng: np.random.Generator) -> dict:
    product = dict(rng.choice([p for p in PRODUCT_CATALOG if p["Product_Category"] == "Apparel"]))
    actual_weight = jitter(product["weight"], rng, pct=0.10)
    # Deliberately represent a large corrugated box rather than the true apparel size.
    length = round(rng.uniform(45, 62), 2)
    width = round(rng.uniform(35, 48), 2)
    height = round(rng.uniform(22, 35), 2)
    billed_vol = volumetric_kg(length, width, height)
    province = rng.choice(PROVINCES)
    return {
        "Order_ID": order_id,
        "SKU": product["SKU"],
        "Product_Category": product["Product_Category"],
        "Actual_Weight_KG": actual_weight,
        "Length_cm": length,
        "Width_cm": width,
        "Height_cm": height,
        "Billed_Vol_KG": billed_vol,
        "Billed_Cost_ZAR": courier_cost(billed_vol, actual_weight, province, rng),
        "Destination_Province": province,
    }


def generate_dataset(row_count: int = ROW_COUNT, seed: int = RNG_SEED) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    rows: list[dict] = []
    order_counter = 100000

    # 15 critical billing anomalies: small physical parcels billed at extreme volumetric weights.
    for _ in range(15):
        rows.append(create_billing_anomaly(f"ORD-{order_counter}", rng))
        order_counter += 1

    # 40 single-item apparel packaging inefficiencies: small products packed as large corrugated boxes.
    for _ in range(40):
        rows.append(create_packaging_inefficiency(f"ORD-{order_counter}", rng))
        order_counter += 1

    # Multi-item baskets: 60 orders x 2-4 lines = about 180 rows, comfortably >=30% of final rows.
    multi_item_order_count = 60
    for _ in range(multi_item_order_count):
        order_id = f"ORD-{order_counter}"
        order_counter += 1
        basket_size = int(rng.choice([2, 3, 4], p=[0.45, 0.40, 0.15]))
        basket_categories = rng.choice(["Supplements", "Fitness Gear", "Apparel"], size=basket_size, p=[0.42, 0.24, 0.34])
        basket_lines = [create_normal_line(order_id, rng, category=str(category)) for category in basket_categories]

        # In roughly one third of mixed baskets, inflate billed volumetric kg to create packaging bloat cases.
        if rng.random() < 0.34:
            combined_physical_vol_kg = sum(
                volumetric_kg(line["Length_cm"], line["Width_cm"], line["Height_cm"]) for line in basket_lines
            )
            inflated_billed_vol = round(combined_physical_vol_kg * rng.uniform(1.45, 2.60), 2)
            for line in basket_lines:
                line["Billed_Vol_KG"] = max(line["Billed_Vol_KG"], inflated_billed_vol)
                line["Billed_Cost_ZAR"] = courier_cost(line["Billed_Vol_KG"], line["Actual_Weight_KG"], line["Destination_Province"], rng)

        rows.extend(basket_lines)

    # Fill the remaining rows with ordinary single-item shipments.
    while len(rows) < row_count:
        rows.append(create_normal_line(f"ORD-{order_counter}", rng))
        order_counter += 1

    df = pd.DataFrame(rows[:row_count])

    # Dirty data injection: blank a few critical fields to test fail-safe cleaning.
    dirty_actual_weight_count = 8
    dirty_cost_count = 8
    actual_weight_indices = rng.choice(df.index, size=dirty_actual_weight_count, replace=False)
    remaining_indices = np.array([idx for idx in df.index if idx not in set(actual_weight_indices)])
    cost_indices = rng.choice(remaining_indices, size=dirty_cost_count, replace=False)
    df.loc[actual_weight_indices, "Actual_Weight_KG"] = np.nan
    df.loc[cost_indices, "Billed_Cost_ZAR"] = np.nan

    # Shuffle rows so injected scenarios are not grouped at the top of the CSV.
    df = df.sample(frac=1, random_state=seed).reset_index(drop=True)

    return df[
        [
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
    ]


def main() -> None:
    df = generate_dataset()
    df.to_csv(OUTPUT_FILE, index=False)

    multi_item_order_share = (df.groupby("Order_ID").size() > 1).mean()
    multi_item_order_ids = df.groupby("Order_ID").size().loc[lambda sizes: sizes > 1].index
    multi_item_row_share = df["Order_ID"].isin(multi_item_order_ids).mean()

    print(f"Generated {len(df):,} rows -> {OUTPUT_FILE}")
    print(f"Unique orders: {df['Order_ID'].nunique():,}")
    print(f"Multi-item order share: {multi_item_order_share:.1%}")
    print(f"Multi-item row share: {multi_item_row_share:.1%}")
    print(f"Blank Actual_Weight_KG cells: {df['Actual_Weight_KG'].isna().sum():,}")
    print(f"Blank Billed_Cost_ZAR cells: {df['Billed_Cost_ZAR'].isna().sum():,}")


if __name__ == "__main__":
    main()
