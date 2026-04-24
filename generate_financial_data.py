"""
Financial Data Generator
Creates regional sales/revenue data with intentional anomalies for detection demo.
"""

import pandas as pd
import random
import os
from datetime import date, timedelta

random.seed(99)

REGIONS = ["North", "South", "East", "West", "Central"]
PRODUCTS = ["Product_A", "Product_B", "Product_C", "Product_D", "Product_E"]
CHANNELS = ["Online", "Retail", "Wholesale", "Direct"]
MONTHS = pd.date_range("2023-01-01", "2023-12-31", freq="MS").to_list()

records = []
record_id = 1

for month in MONTHS:
    for region in REGIONS:
        for product in PRODUCTS:
            for channel in CHANNELS:
                units = random.randint(100, 1000)
                unit_price = random.uniform(50, 500)
                revenue = round(units * unit_price, 2)
                cost = round(revenue * random.uniform(0.4, 0.7), 2)
                profit = round(revenue - cost, 2)

                # Inject anomalies (~8% of records)
                anomaly_type = ""
                issue = random.random()

                if issue < 0.02:
                    revenue = -abs(revenue)        # Negative revenue
                    anomaly_type = "Negative Revenue"
                elif issue < 0.04:
                    revenue = revenue * 10         # 10x spike (data entry error)
                    anomaly_type = "Revenue Spike"
                elif issue < 0.05:
                    revenue = 0                    # Zero revenue
                    anomaly_type = "Zero Revenue"
                elif issue < 0.06:
                    cost = revenue * 1.5           # Cost > Revenue (loss anomaly)
                    profit = round(revenue - cost, 2)
                    anomaly_type = "Cost Exceeds Revenue"
                elif issue < 0.07:
                    units = 0                      # Units sold = 0 but revenue exists
                    anomaly_type = "Units Zero but Revenue Non-Zero"
                elif issue < 0.08:
                    region = ""                    # Missing region
                    anomaly_type = "Missing Region"

                records.append({
                    "Record_ID": f"TXN{str(record_id).zfill(5)}",
                    "Month": month.strftime("%Y-%m"),
                    "Region": region,
                    "Product": product,
                    "Channel": channel,
                    "Units_Sold": units,
                    "Unit_Price": round(unit_price, 2),
                    "Revenue": revenue,
                    "Cost": cost,
                    "Profit": profit,
                    "Anomaly_Type": anomaly_type
                })
                record_id += 1

df = pd.DataFrame(records)
os.makedirs("data", exist_ok=True)
df.to_csv("data/financial_raw_data.csv", index=False)

total = len(df)
anomalies = df[df["Anomaly_Type"] != ""]
print(f"✅ Generated {total} records → data/financial_raw_data.csv")
print(f"   Anomalies injected: {len(anomalies)} ({len(anomalies)*100/total:.1f}%)")
print(f"   Types: {anomalies['Anomaly_Type'].value_counts().to_dict()}")
