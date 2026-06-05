import random

import pandas as pd


random.seed(42)

skus = [
    "Cotton T-Shirt",
    "Performance Hoodie",
    "Yoga Leggings",
    "Running Shorts",
    "Sports Bra",
    "Denim Jacket",
    "Chino Pants",
    "Whey Protein 1kg",
    "Creatine Monohydrate",
    "Omega-3 Capsules",
    "Multivitamin Pack",
    "Resistance Bands",
    "Gym Gloves",
    "Trail Socks",
    "Compression Top",
    "Puffer Vest",
    "Canvas Sneakers",
    "Hydration Bottle",
    "Meal Replacement Shake",
    "BCAA Powder",
]

provinces = [
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

price_ranges = {
    "Cotton T-Shirt": (149, 299),
    "Performance Hoodie": (499, 899),
    "Yoga Leggings": (349, 699),
    "Running Shorts": (199, 399),
    "Sports Bra": (249, 499),
    "Denim Jacket": (699, 1299),
    "Chino Pants": (399, 799),
    "Whey Protein 1kg": (449, 799),
    "Creatine Monohydrate": (249, 449),
    "Omega-3 Capsules": (129, 299),
    "Multivitamin Pack": (149, 349),
    "Resistance Bands": (99, 249),
    "Gym Gloves": (149, 299),
    "Trail Socks": (79, 179),
    "Compression Top": (299, 599),
    "Puffer Vest": (599, 1199),
    "Canvas Sneakers": (399, 899),
    "Hydration Bottle": (129, 279),
    "Meal Replacement Shake": (299, 599),
    "BCAA Powder": (249, 499),
}

rows = []
for order_number in range(1001, 1501):
    sku = random.choice(skus)
    low_price, high_price = price_ranges[sku]
    rows.append(
        {
            "Order_ID": f"SA-{order_number}",
            "SKU": sku,
            "Revenue": round(random.uniform(low_price, high_price), 2),
            "Province": random.choices(
                provinces,
                weights=[32, 22, 17, 8, 5, 5, 4, 4, 3],
                k=1,
            )[0],
            "Actual_Weight_kg": round(random.uniform(0.5, 5.0), 2),
            "Length_cm": random.randint(10, 60),
            "Width_cm": random.randint(10, 60),
            "Height_cm": random.randint(10, 60),
            "Billed_Shipping_Cost_ZAR": round(random.uniform(80, 250), 2),
        }
    )

shipping_data = pd.DataFrame(rows)
shipping_data.to_csv("mock_shipping_data.csv", index=False)
print("Created mock_shipping_data.csv with 500 rows.")
