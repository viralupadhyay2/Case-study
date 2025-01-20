import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load the cleaned dataset
file_path = "C:/Users/ASUS/Downloads/Cleaned_Command_Centre_Dataset.xlsx"
sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")

# Calculate Profit if not already calculated
sales_data["Profit"] = sales_data["Revenue"] - sales_data["Cost_of_Goods_Sold"]

# Convert 'Date' column to datetime format (if not already done)
sales_data['Date'] = pd.to_datetime(sales_data['Date'])

# Group by Date to calculate total Profit over time
profit_trend = sales_data.groupby('Date')["Profit"].sum().reset_index()

# Plot the trend of Profit over time
plt.figure(figsize=(12, 6))
sns.lineplot(data=profit_trend, x='Date', y='Profit', marker="o", color='blue')
plt.title("Trend of Profit Over Time", fontsize=14)
plt.xlabel("Date", fontsize=12)
plt.ylabel("Total Profit", fontsize=12)
plt.grid(True)
plt.tight_layout()

# Save the plot
plt.savefig("Profit_Trend_Over_Time.png")

# Display the plot
plt.show()

print("\nINFO: Trend of profit over time visualization saved as 'Profit_Trend_Over_Time.png'")
