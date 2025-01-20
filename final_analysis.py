import math
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# File paths
file_path = "C:/Users/ASUS/Downloads/Command Centre Dataset.xlsx"
output_file_path = "C:/Users/ASUS/Downloads/Final_Command_Centre_Analysis_1.xlsx"

# === Data Preparation ===

# Load datasets
try:
    sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")
    marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")
    customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
    competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")
    print("Datasets loaded successfully!")
except FileNotFoundError:
    print(f"Error: File not found at {file_path}. Please check the path.")
    exit()

# Standardize Date Format for Datasets with a 'Date' Column
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
marketing_data['Date'] = pd.to_datetime(marketing_data['Date'])
customer_feedback_data['Date'] = pd.to_datetime(customer_feedback_data['Date'])

# Remove Duplicates
sales_data = sales_data.drop_duplicates()
marketing_data = marketing_data.drop_duplicates()
customer_feedback_data = customer_feedback_data.drop_duplicates()
competitor_data = competitor_data.drop_duplicates()

# Rename columns for consistency
if "Competitor_Product_Category" in competitor_data.columns:
    competitor_data.rename(columns={"Competitor_Product_Category": "Product_Category"}, inplace=True)

# Merge datasets
merged_data = pd.merge(sales_data, marketing_data, on=["Country", "Zone", "Date"], how="outer")
merged_data = pd.merge(merged_data, customer_feedback_data, on=["Country", "Zone", "Product_Category", "Date"], how="outer")
final_data = pd.merge(merged_data, competitor_data, on=["Country", "Zone", "Product_Category"], how="outer")

# Calculate Profit
final_data["Profit"] = final_data["Revenue"].fillna(0) - final_data["Cost_of_Goods_Sold"].fillna(0)

# Fill Missing Values
final_data["Customer_Rating"] = final_data["Customer_Rating"].fillna("No Rating")
final_data["Country"] = final_data["Country"].fillna("Unknown Country")
final_data["Zone"] = final_data["Zone"].fillna("Unknown Zone")
final_data["Product_Category"] = final_data["Product_Category"].fillna("Unknown Category")
final_data["Profit"] = final_data["Profit"].fillna(0)

# Reduce rows by sampling (optimal row count: 700,000)
reduced_final_data = final_data.sample(n=700000, random_state=42)

# Select required columns for the reduced data
filtered_final_data = reduced_final_data[["Country", "Zone", "Profit", "Product_Category", "Customer_Rating"]]

# === Save Results ===

try:
    with pd.ExcelWriter(output_file_path, mode="w", engine="openpyxl") as writer:
        filtered_final_data.to_excel(writer, sheet_name="Reduced_Merged_Data", index=False)
    print("\nINFO: Reduced dataset saved successfully!")
except Exception as e:
    print(f"\nERROR: {e}")
