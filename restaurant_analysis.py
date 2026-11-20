
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# ==========================================
# RESTAURANT ORDERS AND MENU PERFORMANCE
# ANALYSIS PROJECT
# ==========================================

print("=" * 55)
print(" RESTAURANT ORDERS AND MENU PERFORMANCE ANALYSIS")
print("=" * 55)

# STEP 1: Load Excel Dataset

file_name = "Restaurant_Orders_600_Records.xlsx"

try:
    df = pd.read_excel(file_name)
    print("\nDataset loaded successfully!")

except FileNotFoundError:
    print("\nERROR: Excel file not found!")
    print("Keep the Excel file in the same folder as this Python file.")
    raise SystemExit

# STEP 2: Display Basic Dataset Information

print("\n--- DATASET INFORMATION ---")

print("\nFirst 5 Records:")
print(df.head())

print("\nDataset Shape:")
print(df.shape)

print("\nColumn Names:")
print(df.columns.tolist())

print("\nDataset Information:")
df.info()

# STEP 3: Data Cleaning

print("\n--- DATA CLEANING ---")

# Remove extra spaces from column names
df.columns = df.columns.str.strip()

# Check required columns
required_columns = [
    "Order_ID",
    "Food_Item",
    "Category",
    "Quantity",
    "Price"
]

if not all(col in df.columns for col in required_columns):
    print("ERROR: Required columns are missing.")
    raise SystemExit

# Remove duplicate records
df.drop_duplicates(inplace=True)

# Remove records with missing important values
df.dropna(subset=required_columns, inplace=True)

# Convert numerical columns
df["Quantity"] = pd.to_numeric(
    df["Quantity"], errors="coerce"
)

df["Price"] = pd.to_numeric(
    df["Price"], errors="coerce"
)

# Remove invalid numerical values
df.dropna(subset=["Quantity", "Price"], inplace=True)

df = df[
    (df["Quantity"] > 0) &
    (df["Price"] >= 0)
].copy()

# STEP 4: Calculate Total Sales

df["Total_Sales"] = df["Quantity"] * df["Price"]

print("\nData cleaning completed!")

print("Total Valid Records:", len(df))

# STEP 5: Overall Restaurant Performance

total_orders = df["Order_ID"].nunique()

total_quantity = df["Quantity"].sum()

total_revenue = df["Total_Sales"].sum()

average_order_value = (
    total_revenue / total_orders
    if total_orders > 0 else 0
)

print("\n--- OVERALL PERFORMANCE ---")

print("Total Orders:", total_orders)

print("Total Items Sold:", total_quantity)

print("Total Revenue: Rs.", round(total_revenue, 2))

print(
    "Average Revenue Per Order: Rs.",
    round(average_order_value, 2)
)

# STEP 6: Food Item Performance

item_performance = df.groupby("Food_Item").agg(
    Total_Quantity=("Quantity", "sum"),
    Total_Revenue=("Total_Sales", "sum"),
    Number_of_Orders=("Order_ID", "nunique")
).reset_index()

item_performance = item_performance.sort_values(
    by="Total_Revenue",
    ascending=False
)

print("\n--- FOOD ITEM PERFORMANCE ---")
print(item_performance.to_string(index=False))

# STEP 7: Category Performance

category_performance = df.groupby("Category").agg(
    Total_Quantity=("Quantity", "sum"),
    Total_Revenue=("Total_Sales", "sum"),
    Number_of_Orders=("Order_ID", "nunique")
).reset_index()

category_performance = category_performance.sort_values(
    by="Total_Revenue",
    ascending=False
)

print("\n--- CATEGORY PERFORMANCE ---")
print(category_performance.to_string(index=False))

# STEP 8: Best Selling Food Item

best_selling_item = (
    df.groupby("Food_Item")["Quantity"]
    .sum()
    .idxmax()
)

print("\nBest Selling Food Item:", best_selling_item)

# STEP 9: Highest Revenue Generating Food Item

highest_revenue_item = (
    df.groupby("Food_Item")["Total_Sales"]
    .sum()
    .idxmax()
)

print(
    "Highest Revenue Generating Item:",
    highest_revenue_item
)

# STEP 10: Highest Performing Category

best_category = (
    df.groupby("Category")["Total_Sales"]
    .sum()
    .idxmax()
)

print("Best Performing Category:", best_category)

# STEP 11: Create Charts Folder

os.makedirs("charts", exist_ok=True)

# CHART 1: Revenue by Food Item

item_revenue = df.groupby("Food_Item")[
    "Total_Sales"
].sum().sort_values(ascending=False)

plt.figure(figsize=(10, 6))

item_revenue.plot(
    kind="bar",
    edgecolor="black"
)

plt.title("Revenue by Food Item")
plt.xlabel("Food Items")
plt.ylabel("Revenue (Rs.)")
plt.xticks(rotation=35)
plt.tight_layout()

plt.savefig("charts/food_item_revenue.png")
plt.show()

# CHART 2: Quantity Sold by Food Item

item_quantity = df.groupby("Food_Item")[
    "Quantity"
].sum().sort_values(ascending=False)

plt.figure(figsize=(10, 6))

item_quantity.plot(
    kind="bar",
    edgecolor="black"
)

plt.title("Quantity Sold by Food Item")
plt.xlabel("Food Items")
plt.ylabel("Quantity Sold")
plt.xticks(rotation=35)
plt.tight_layout()

plt.savefig("charts/food_item_quantity.png")
plt.show()

# CHART 3: Revenue by Category

category_revenue = df.groupby("Category")[
    "Total_Sales"
].sum()

plt.figure(figsize=(8, 8))

plt.pie(
    category_revenue,
    labels=category_revenue.index,
    autopct="%1.1f%%",
    startangle=90
)

plt.title("Revenue Contribution by Category")
plt.tight_layout()

plt.savefig("charts/category_revenue.png")
plt.show()

# CHART 4: Orders by Category

category_orders = df.groupby("Category")[
    "Order_ID"
].nunique()

plt.figure(figsize=(9, 6))

category_orders.plot(
    kind="bar",
    edgecolor="black"
)

plt.title("Number of Orders by Category")
plt.xlabel("Category")
plt.ylabel("Number of Orders")
plt.xticks(rotation=30)
plt.tight_layout()

plt.savefig("charts/category_orders.png")
plt.show()

# CHART 5: Food Item Sales Comparison

plt.figure(figsize=(10, 6))

plt.scatter(
    item_quantity,
    item_revenue[item_quantity.index],
    s=100
)

plt.title("Food Item Quantity vs Revenue")
plt.xlabel("Quantity Sold")
plt.ylabel("Revenue (Rs.)")
plt.tight_layout()

plt.savefig("charts/quantity_vs_revenue.png")
plt.show()

# STEP 12: Export Analysis Report to Excel

with pd.ExcelWriter(
    "Restaurant_Analysis_Report.xlsx",
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Cleaned_Data",
        index=False
    )

    item_performance.to_excel(
        writer,
        sheet_name="Item_Performance",
        index=False
    )

    category_performance.to_excel(
        writer,
        sheet_name="Category_Performance",
        index=False
    )

    summary = pd.DataFrame({
        "Metric": [
            "Total Orders",
            "Total Items Sold",
            "Total Revenue",
            "Average Revenue Per Order",
            "Best Selling Item",
            "Highest Revenue Item",
            "Best Performing Category"
        ],

        "Value": [
            total_orders,
            total_quantity,
            round(total_revenue, 2),
            round(average_order_value, 2),
            best_selling_item,
            highest_revenue_item,
            best_category
        ]
    })

    summary.to_excel(
        writer,
        sheet_name="Overall_Summary",
        index=False
    )

# STEP 13: Final Message

print("\n" + "=" * 55)
print(" PROJECT COMPLETED SUCCESSFULLY!")
print("=" * 55)

print("\nGenerated Files:")
print("1. Restaurant_Analysis_Report.xlsx")
print("2. charts/food_item_revenue.png")
print("3. charts/food_item_quantity.png")
print("4. charts/category_revenue.png")
print("5. charts/category_orders.png")
print("6. charts/quantity_vs_revenue.png")

print("\nThank you for using Restaurant Analysis System!")
