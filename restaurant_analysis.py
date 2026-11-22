
import pandas as pd
import os 

# ==========================================
# RESTAURANT ORDERS AND MENU PERFORMANCE
# ANALYSIS PROJECT
# ==========================================

print("=" * 55)
print(" RESTAURANT ORDERS AND MENU PERFORMANCE ANALYSIS")
print("=" * 55)




file_name = "Restaurant_Orders_600_Records.xlsx"

try:
    df = pd.read_excel(file_name)
    print("\nDataset loaded successfully!")

except FileNotFoundError:
    print("\nERROR: Excel file not found!")
    print("Keep the Excel file in the same folder as this Python file.")
    raise SystemExit



print("\n--- DATASET INFORMATION ---")

print("\nFirst 5 Records:")
print(df.head())

print("\nDataset Shape:")
print(df.shape)

print("\nColumn Names:")
print(df.columns.tolist())

print("\nDataset Information:")
df.info()


# STEP 3: DATA CLEANING

print("\n--- DATA CLEANING ---")

# Remove extra spaces from column names
df.columns = df.columns.str.strip()

# Required columns
required_columns = [
    "Order_ID",
    "Food_Item",
    "Category",
    "Quantity",
    "Price"
]

# Check required columns
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

# Remove invalid quantities and prices
df = df[
    (df["Quantity"] > 0) &
    (df["Price"] >= 0)
].copy()

print("\nData cleaning completed!")
print("Total Valid Records:", len(df))


# STEP 4: CALCULATE TOTAL SALES

# Revenue = Quantity x Price

df["Total_Sales"] = df["Quantity"] * df["Price"]

print("\nTotal sales column created successfully!")


# STEP 5: OVERALL RESTAURANT PERFORMANCE

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


# STEP 6: FOOD ITEM PERFORMANCE

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


# STEP 7: CATEGORY PERFORMANCE

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


# STEP 8: BEST SELLING FOOD ITEM

best_selling_item = (
    df.groupby("Food_Item")["Quantity"]
    .sum()
    .idxmax()
)

print("\nBest Selling Food Item:", best_selling_item)


# STEP 9: HIGHEST REVENUE GENERATING FOOD ITEM

highest_revenue_item = (
    df.groupby("Food_Item")["Total_Sales"]
    .sum()
    .idxmax()
)

print(
    "Highest Revenue Generating Item:",
    highest_revenue_item
)


# STEP 10: BEST PERFORMING CATEGORY

best_category = (
    df.groupby("Category")["Total_Sales"]
    .sum()
    .idxmax()
)

print("Best Performing Category:", best_category)


# STEP 11: QUANTITY VS REVENUE ANALYSIS

item_quantity = df.groupby("Food_Item")[
    "Quantity"
].sum()

item_revenue = df.groupby("Food_Item")[
    "Total_Sales"
].sum()

correlation = item_quantity.corr(item_revenue)

print("\n--- QUANTITY VS REVENUE ANALYSIS ---")

print(
    "Correlation between Quantity and Revenue:",
    round(correlation, 3)
)

if correlation > 0:
    print("There is a positive relationship between quantity and revenue.")

elif correlation < 0:
    print("There is a negative relationship between quantity and revenue.")

else:
    print("There is no linear relationship between quantity and revenue.")


# STEP 12: CREATE EXCEL ANALYSIS REPORT

print("\n--- GENERATING EXCEL REPORT ---")

with pd.ExcelWriter(
    "Restaurant_Analysis_Report.xlsx",
    engine="openpyxl"
) as writer:

    # Sheet 1: Cleaned Data
    df.to_excel(
        writer,
        sheet_name="Cleaned_Data",
        index=False
    )

    # Sheet 2: Item Performance
    item_performance.to_excel(
        writer,
        sheet_name="Item_Performance",
        index=False
    )

    # Sheet 3: Category Performance
    category_performance.to_excel(
        writer,
        sheet_name="Category_Performance",
        index=False
    )

    # Sheet 4: Overall Summary
    summary = pd.DataFrame({

        "Metric": [
            "Total Orders",
            "Total Items Sold",
            "Total Revenue",
            "Average Revenue Per Order",
            "Best Selling Item",
            "Highest Revenue Item",
            "Best Performing Category",
            "Quantity-Revenue Correlation"
        ],

        "Value": [
            total_orders,
            total_quantity,
            round(total_revenue, 2),
            round(average_order_value, 2),
            best_selling_item,
            highest_revenue_item,
            best_category,
            round(correlation, 3)
        ]

    })

    summary.to_excel(
        writer,
        sheet_name="Overall_Summary",
        index=False
    )


# STEP 13: FINAL MESSAGE

print("\n" + "=" * 55)

print(" PROJECT COMPLETED SUCCESSFULLY!")

print("=" * 55)

print("\nGenerated File:")

print("1. Restaurant_Analysis_Report.xlsx")

print("\nExcel Report Contains:")

print("1. Cleaned_Data")
print("2. Item_Performance")
print("3. Category_Performance")
print("4. Overall_Summary")

print("\nThank you for using Restaurant Analysis System!")

