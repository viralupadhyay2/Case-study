import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# File Path
file_path = "C:/Users/ASUS/Downloads/Command Centre Dataset.xlsx"

# Load datasets
sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")
marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")
customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")

# Standardize Date Format
sales_data['Date'] = pd.to_datetime(sales_data['Date'])
marketing_data['Date'] = pd.to_datetime(marketing_data['Date'])
customer_feedback_data['Date'] = pd.to_datetime(customer_feedback_data['Date'])

# Merge Sales and Marketing
merge_1 = pd.merge(sales_data, marketing_data, on=["Country", "Zone", "Date"], how="outer", indicator=True)
merge_1_summary = merge_1["_merge"].value_counts()

# Visualize Merge 1
plt.figure(figsize=(8, 6))
sns.barplot(x=merge_1_summary.index, y=merge_1_summary.values, palette="viridis")
plt.title("Sales and Marketing Data Merge Status")
plt.xlabel("Merge Status")
plt.ylabel("Number of Records")
plt.tight_layout()
plt.savefig("Sales_Marketing_Merge_Status.png")
plt.show()

# Merge Result with Customer Feedback
merge_2 = pd.merge(merge_1, customer_feedback_data, on=["Country", "Zone", "Product_Category", "Date"], how="outer", indicator=True)
merge_2_summary = merge_2["_merge"].value_counts()

# Visualize Merge 2
plt.figure(figsize=(8, 6))
sns.barplot(x=merge_2_summary.index, y=merge_2_summary.values, palette="coolwarm")
plt.title("Merged Data with Customer Feedback Data Status")
plt.xlabel("Merge Status")
plt.ylabel("Number of Records")
plt.tight_layout()
plt.savefig("Merge_With_Customer_Feedback_Status.png")
plt.show()

# Merge Final with Competitor Data
final_merge = pd.merge(merge_2, competitor_data, on=["Country", "Zone", "Product_Category"], how="outer", indicator=True)
final_merge_summary = final_merge["_merge"].value_counts()

# Visualize Final Merge
plt.figure(figsize=(8, 6))
sns.barplot(x=final_merge_summary.index, y=final_merge_summary.values, palette="magma")
plt.title("Final Merge with Competitor Data Status")
plt.xlabel("Merge Status")
plt.ylabel("Number of Records")
plt.tight_layout()
plt.savefig("Final_Merge_Status.png")
plt.show()

print("INFO: Merge visualizations saved successfully!")
