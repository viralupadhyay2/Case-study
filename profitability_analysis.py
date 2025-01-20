# profitability_analysis.py
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load cleaned data
file_path = "Cleaned_Command_Centre_Dataset.xlsx"
sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")

# Calculate Profit
sales_data["Profit"] = sales_data["Revenue"] - sales_data["Cost_of_Goods_Sold"]

# Group by Zone and Product Category
profitability = sales_data.groupby(["Zone", "Product_Category"])["Profit"].sum().reset_index()

# Visualize Profitability
plt.figure(figsize=(10, 6))
sns.barplot(data=profitability, x="Zone", y="Profit", hue="Product_Category", palette="viridis")
plt.title("Profitability by Zone and Product Category")
plt.xlabel("Zone")
plt.ylabel("Total Profit")
plt.tight_layout()
plt.savefig("Profitability_Analysis.png")
plt.show()
