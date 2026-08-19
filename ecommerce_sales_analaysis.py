import pandas as pd
import matplotlib.pyplot as plt
import os

# ==========================================
# 1. PROJECT SETUP
# ==========================================

print("=" * 60)
print("       E-COMMERCE SALES DATA ANALYSIS")
print("=" * 60)

# Create output folders
os.makedirs("charts", exist_ok=True)
os.makedirs("output", exist_ok=True)


# ==========================================
# 2. LOAD DATA
# ==========================================

try:
    df = pd.read_excel("ecommerce_sales_data.xlsx")
    print("\nData loaded successfully!")
except FileNotFoundError:
    print("\nERROR: Excel file not found!")
    print("Make sure the Excel file is in the same folder.")
    exit()


# ==========================================
# 3. DATA UNDERSTANDING
# ==========================================

print("\n" + "=" * 60)
print("DATASET INFORMATION")
print("=" * 60)

print("\nDataset Shape:")
print(f"Rows: {df.shape[0]}")
print(f"Columns: {df.shape[1]}")

print("\nColumn Names:")
print(df.columns.tolist())

print("\nFirst 5 Rows:")
print(df.head())

print("\nData Types:")
print(df.dtypes)


# ==========================================
# 4. DATA CLEANING
# ==========================================

print("\n" + "=" * 60)
print("DATA CLEANING")
print("=" * 60)

# Missing values before cleaning
print("\nMissing Values Before Cleaning:")
print(df.isnull().sum())

# Remove missing values
df = df.dropna()

# Check duplicate rows
duplicate_count = df.duplicated().sum()

print(f"\nDuplicate Rows Found: {duplicate_count}")

# Remove duplicates
df = df.drop_duplicates()

print(f"\nFinal Dataset Shape: {df.shape}")


# ==========================================
# 5. DATE CONVERSION
# ==========================================

df["Order_Date"] = pd.to_datetime(
    df["Order_Date"],
    errors="coerce"
)

# Remove invalid dates
df = df.dropna(subset=["Order_Date"])

# Create new date columns
df["Year"] = df["Order_Date"].dt.year
df["Month"] = df["Order_Date"].dt.month
df["Month_Name"] = df["Order_Date"].dt.month_name()
df["Year_Month"] = df["Order_Date"].dt.to_period("M").astype(str)


# ==========================================
# 6. CREATE NEW BUSINESS METRICS
# ==========================================

# Profit Margin
df["Profit_Margin"] = (
    df["Profit"] / df["Sales"] * 100
)

# Handle division issues
df["Profit_Margin"] = df["Profit_Margin"].replace(
    [float("inf"), -float("inf")],
    0
)

# Discount Amount
df["Discount_Amount"] = (
    df["Unit_Price"]
    * df["Quantity"]
    * df["Discount_Percent"]
    / 100
)


# ==========================================
# 7. SAVE CLEANED DATA
# ==========================================

df.to_csv(
    "output/cleaned_ecommerce_data.csv",
    index=False
)

print("\nCleaned data saved successfully!")


# ==========================================
# 8. KPI CALCULATIONS
# ==========================================

total_sales = df["Sales"].sum()

total_profit = df["Profit"].sum()

total_orders = df["Order_ID"].nunique()

total_customers = df["Customer"].nunique()

total_quantity = df["Quantity"].sum()

average_order_value = total_sales / total_orders

profit_margin = (
    total_profit / total_sales
) * 100


print("\n" + "=" * 60)
print("BUSINESS KPI SUMMARY")
print("=" * 60)

print(f"\nTotal Sales: ₹{total_sales:,.2f}")

print(f"Total Profit: ₹{total_profit:,.2f}")

print(f"Total Orders: {total_orders:,}")

print(f"Total Customers: {total_customers:,}")

print(f"Total Quantity Sold: {total_quantity:,}")

print(f"Average Order Value: ₹{average_order_value:,.2f}")

print(f"Overall Profit Margin: {profit_margin:.2f}%")


# ==========================================
# 9. CATEGORY ANALYSIS
# ==========================================

category_analysis = (
    df.groupby("Category")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Total_Quantity=("Quantity", "sum"),
        Average_Profit_Margin=("Profit_Margin", "mean")
    )
    .sort_values(
        by="Total_Sales",
        ascending=False
    )
)

print("\n" + "=" * 60)
print("CATEGORY ANALYSIS")
print("=" * 60)

print(category_analysis)

category_analysis.to_csv(
    "output/category_analysis.csv"
)


# ==========================================
# 10. TOP 10 PRODUCTS
# ==========================================

top_products = (
    df.groupby("Product")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Quantity_Sold=("Quantity", "sum")
    )
    .sort_values(
        by="Total_Sales",
        ascending=False
    )
    .head(10)
)

print("\n" + "=" * 60)
print("TOP 10 PRODUCTS")
print("=" * 60)

print(top_products)

top_products.to_csv(
    "output/top_10_products.csv"
)


# ==========================================
# 11. TOP 10 CUSTOMERS
# ==========================================

top_customers = (
    df.groupby("Customer")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Orders=("Order_ID", "nunique"),
        Total_Profit=("Profit", "sum")
    )
    .sort_values(
        by="Total_Sales",
        ascending=False
    )
    .head(10)
)

print("\n" + "=" * 60)
print("TOP 10 CUSTOMERS")
print("=" * 60)

print(top_customers)

top_customers.to_csv(
    "output/top_10_customers.csv"
)


# ==========================================
# 12. STATE ANALYSIS
# ==========================================

state_analysis = (
    df.groupby("State")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Orders=("Order_ID", "nunique")
    )
    .sort_values(
        by="Total_Sales",
        ascending=False
    )
)

print("\n" + "=" * 60)
print("STATE ANALYSIS")
print("=" * 60)

print(state_analysis)

state_analysis.to_csv(
    "output/state_analysis.csv"
)


# ==========================================
# 13. PAYMENT METHOD ANALYSIS
# ==========================================

payment_analysis = (
    df.groupby("Payment_Method")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Orders=("Order_ID", "nunique"),
        Total_Profit=("Profit", "sum")
    )
    .sort_values(
        by="Total_Sales",
        ascending=False
    )
)

print("\n" + "=" * 60)
print("PAYMENT METHOD ANALYSIS")
print("=" * 60)

print(payment_analysis)

payment_analysis.to_csv(
    "output/payment_analysis.csv"
)


# ==========================================
# 14. DISCOUNT IMPACT ANALYSIS
# ==========================================

discount_analysis = (
    df.groupby("Discount_Percent")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Orders=("Order_ID", "nunique"),
        Average_Profit_Margin=("Profit_Margin", "mean")
    )
    .sort_index()
)

print("\n" + "=" * 60)
print("DISCOUNT IMPACT ANALYSIS")
print("=" * 60)

print(discount_analysis)

discount_analysis.to_csv(
    "output/discount_analysis.csv"
)


# ==========================================
# 15. MONTHLY SALES TREND
# ==========================================

monthly_analysis = (
    df.groupby("Year_Month")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Profit=("Profit", "sum"),
        Total_Orders=("Order_ID", "nunique")
    )
    .sort_index()
)

print("\n" + "=" * 60)
print("MONTHLY PERFORMANCE")
print("=" * 60)

print(monthly_analysis)

monthly_analysis.to_csv(
    "output/monthly_analysis.csv"
)


# ==========================================
# 16. REPEAT CUSTOMER ANALYSIS
# ==========================================

customer_orders = (
    df.groupby("Customer")["Order_ID"]
    .nunique()
)

repeat_customers = (
    customer_orders[customer_orders > 1]
)

repeat_customer_count = len(repeat_customers)

repeat_customer_percentage = (
    repeat_customer_count
    / total_customers
    * 100
)

print("\n" + "=" * 60)
print("REPEAT CUSTOMER ANALYSIS")
print("=" * 60)

print(f"\nRepeat Customers: {repeat_customer_count}")

print(
    f"Repeat Customer Percentage: "
    f"{repeat_customer_percentage:.2f}%"
)


# ==========================================
# 17. CHART 1 - SALES BY CATEGORY
# ==========================================

plt.figure(figsize=(10, 6))

plt.bar(
    category_analysis.index,
    category_analysis["Total_Sales"]
)

plt.title("Sales by Category")
plt.xlabel("Category")
plt.ylabel("Total Sales")

plt.xticks(rotation=45)

plt.tight_layout()

plt.savefig(
    "charts/01_sales_by_category.png"
)

plt.close()


# ==========================================
# 18. CHART 2 - MONTHLY SALES TREND
# ==========================================

plt.figure(figsize=(12, 6))

plt.plot(
    monthly_analysis.index,
    monthly_analysis["Total_Sales"],
    marker="o"
)

plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales")

plt.xticks(rotation=45)

plt.tight_layout()

plt.savefig(
    "charts/02_monthly_sales_trend.png"
)

plt.close()


# ==========================================
# 19. CHART 3 - TOP 10 PRODUCTS
# ==========================================

plt.figure(figsize=(12, 6))

plt.bar(
    top_products.index,
    top_products["Total_Sales"]
)

plt.title("Top 10 Products by Sales")
plt.xlabel("Product")
plt.ylabel("Total Sales")

plt.xticks(rotation=45)

plt.tight_layout()

plt.savefig(
    "charts/03_top_10_products.png"
)

plt.close()


# ==========================================
# 20. CHART 4 - SALES BY STATE
# ==========================================

plt.figure(figsize=(12, 6))

plt.bar(
    state_analysis.index,
    state_analysis["Total_Sales"]
)

plt.title("Sales by State")
plt.xlabel("State")
plt.ylabel("Total Sales")

plt.xticks(rotation=45)

plt.tight_layout()

plt.savefig(
    "charts/04_sales_by_state.png"
)

plt.close()


# ==========================================
# 21. CHART 5 - PAYMENT METHOD
# ==========================================

plt.figure(figsize=(10, 6))

plt.bar(
    payment_analysis.index,
    payment_analysis["Total_Sales"]
)

plt.title("Sales by Payment Method")
plt.xlabel("Payment Method")
plt.ylabel("Total Sales")

plt.xticks(rotation=45)

plt.tight_layout()

plt.savefig(
    "charts/05_payment_method_sales.png"
)

plt.close()


# ==========================================
# 22. CHART 6 - PROFIT BY CATEGORY
# ==========================================

plt.figure(figsize=(10, 6))

plt.bar(
    category_analysis.index,
    category_analysis["Total_Profit"]
)

plt.title("Profit by Category")
plt.xlabel("Category")
plt.ylabel("Total Profit")

plt.xticks(rotation=45)

plt.tight_layout()

plt.savefig(
    "charts/06_profit_by_category.png"
)

plt.close()


# ==========================================
# 23. CREATE FINAL SUMMARY
# ==========================================

summary = pd.DataFrame({
    "Metric": [
        "Total Sales",
        "Total Profit",
        "Total Orders",
        "Total Customers",
        "Total Quantity Sold",
        "Average Order Value",
        "Overall Profit Margin %",
        "Repeat Customers",
        "Repeat Customer Percentage"
    ],
    
    "Value": [
        total_sales,
        total_profit,
        total_orders,
        total_customers,
        total_quantity,
        average_order_value,
        profit_margin,
        repeat_customer_count,
        repeat_customer_percentage
    ]
})

summary.to_csv(
    "output/business_summary.csv",
    index=False
)


# ==========================================
# PROJECT COMPLETION
# ==========================================

print("\n" + "=" * 60)
print("PROJECT COMPLETED SUCCESSFULLY!")
print("=" * 60)

print("\nFiles Created:")

print("1. output/cleaned_ecommerce_data.csv")
print("2. output/category_analysis.csv")
print("3. output/top_10_products.csv")
print("4. output/top_10_customers.csv")
print("5. output/state_analysis.csv")
print("6. output/payment_analysis.csv")
print("7. output/discount_analysis.csv")
print("8. output/monthly_analysis.csv")
print("9. output/business_summary.csv")

print("\nCharts saved in charts folder!")

print("\nThank you!")