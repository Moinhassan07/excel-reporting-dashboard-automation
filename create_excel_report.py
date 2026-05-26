import pandas as pd
import os

os.makedirs("output", exist_ok=True)

ops_df = pd.read_csv("data/ev_operations.csv")
lookup_df = pd.read_csv("data/lookup_data.csv")

ops_df["trip_date"] = pd.to_datetime(ops_df["trip_date"])
ops_df["week"] = ops_df["trip_date"].dt.isocalendar().week
ops_df["month"] = ops_df["trip_date"].dt.strftime("%Y-%m")

cleaned_df = ops_df.merge(lookup_df, on="bike_id", how="left")

weekly_km = (
    cleaned_df.groupby(["week", "city"], as_index=False)["distance_km"]
    .sum()
    .rename(columns={"distance_km": "total_weekly_km"})
)

monthly_km = (
    cleaned_df.groupby(["month", "city"], as_index=False)["distance_km"]
    .sum()
    .rename(columns={"distance_km": "total_monthly_km"})
)

city_summary = (
    cleaned_df.groupby("city", as_index=False)
    .agg(
        total_distance_km=("distance_km", "sum"),
        avg_battery_used_percent=("battery_used_percent", "mean"),
        total_charge_cycles=("charge_cycles", "sum"),
        active_bikes=("status", lambda x: (x == "active").sum())
    )
)

vehicle_summary = (
    cleaned_df.groupby("vehicle_type", as_index=False)
    .agg(
        total_distance_km=("distance_km", "sum"),
        avg_battery_used_percent=("battery_used_percent", "mean"),
        bike_count=("bike_id", "count")
    )
)

output_file = "output/excel_reporting_dashboard.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    ops_df.to_excel(writer, sheet_name="Raw_Data", index=False)
    lookup_df.to_excel(writer, sheet_name="Lookup_Data", index=False)
    cleaned_df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    weekly_km.to_excel(writer, sheet_name="Weekly_KM", index=False)
    monthly_km.to_excel(writer, sheet_name="Monthly_KM", index=False)
    city_summary.to_excel(writer, sheet_name="City_Summary", index=False)
    vehicle_summary.to_excel(writer, sheet_name="Vehicle_Summary", index=False)

print(f"Excel workbook created successfully: {output_file}")