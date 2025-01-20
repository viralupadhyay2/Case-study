# data_preparation.py
import pandas as pd

# File path
file_path = "C:/Users/ASUS/Downloads/Command Centre Dataset (1).xlsx"
output_path = "Cleaned_Command_Centre_Dataset.xlsx"

# Load datasets
sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")
marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")
customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")

# Clean Data
sales_data.fillna({"Country": "Unknown", "Zone": "Unknown", "Revenue": 0, "Cost_of_Goods_Sold": 0}, inplace=True)
marketing_data.fillna({"Country": "Unknown", "Zone": "Unknown", "Marketing_Spend": 0}, inplace=True)
customer_feedback_data.fillna({"Country": "Unknown", "Zone": "Unknown", "Customer_Rating": 0}, inplace=True)
competitor_data.fillna({"Country": "Unknown", "Zone": "Unknown", "Competitor_Price": 0}, inplace=True)

# Standardize Dates
for df in [sales_data, marketing_data, customer_feedback_data]:
    df['Date'] = pd.to_datetime(df['Date'])

# Save Cleaned Data
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    sales_data.to_excel(writer, sheet_name="Sales_Data", index=False)
    marketing_data.to_excel(writer, sheet_name="Marketing_Data", index=False)
    customer_feedback_data.to_excel(writer, sheet_name="Customer_Feedback_Data", index=False)
    competitor_data.to_excel(writer, sheet_name="Competitor_Data", index=False)

print("Cleaned data saved successfully!")
