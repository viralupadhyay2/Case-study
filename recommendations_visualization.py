import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load cleaned dataset
file_path = "C:/Users/ASUS/Downloads/Cleaned_Command_Centre_Dataset.xlsx"

# Load datasets
sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")
customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")

# === Profitability Insights ===
sales_data["Profit"] = sales_data["Revenue"] - sales_data["Cost_of_Goods_Sold"]
profitability = sales_data.groupby(["Zone", "Product_Category"])["Profit"].sum().reset_index()
low_profit_zones = profitability.sort_values(by="Profit", ascending=True).head(5)

# === Brand Popularity Insights ===
# Use customer_feedback_data for brand popularity metrics
brand_popularity = customer_feedback_data.groupby("Country")[
    ["Customer_Rating", "Volume_of_Feedback"]
].mean().reset_index()

low_brand_awareness = brand_popularity.sort_values(by="Customer_Rating", ascending=True).head(5)

# === Competitive Positioning Insights ===
competitor_data.rename(columns={"Competitor_Product_Category": "Product_Category"}, inplace=True)
competitive_analysis = competitor_data.groupby("Product_Category")[
    ["Competitor_Price", "Competitor_Market_Share"]
].mean().reset_index()

# === Visualization ===

# Plot 1: Low Profit Zones and Categories
plt.figure(figsize=(8, 5))
sns.barplot(data=low_profit_zones, x="Zone", y="Profit", hue="Product_Category", palette="viridis")
plt.title("Low Profit Zones and Product Categories")
plt.ylabel("Profit")
plt.tight_layout()
plt.savefig("Low_Profit_Analysis.png")
plt.show()

# Plot 2: Low Brand Awareness
plt.figure(figsize=(8, 5))
sns.barplot(data=low_brand_awareness, x="Customer_Rating", y="Country", palette="coolwarm")
plt.title("Countries with Low Customer Ratings")
plt.xlabel("Customer Rating")
plt.tight_layout()
plt.savefig("Low_Brand_Awareness.png")
plt.show()

# Plot 3: Competitive Positioning
plt.figure(figsize=(8, 5))
sns.scatterplot(
    data=competitive_analysis,
    x="Competitor_Price",
    y="Competitor_Market_Share",
    hue="Product_Category",
    palette="Set2",
    s=100
)
plt.title("Competitive Positioning: Price vs. Market Share")
plt.xlabel("Competitor Price")
plt.ylabel("Market Share (%)")
plt.tight_layout()
plt.savefig("Competitive_Positioning_Insights.png")
plt.show()

print("\nINFO: All recommendation visualizations saved successfully!")
