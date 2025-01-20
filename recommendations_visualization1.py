import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# File path to the Excel dataset
file_path = "C:/Users/ASUS/Downloads/Final_Command_Centre_Analysis1.xlsx"

# Load necessary sheets
try:
    profitability_data = pd.read_excel(file_path, sheet_name="Profitability_Analysis")
    brand_popularity_data = pd.read_excel(file_path, sheet_name="Brand_Popularity_Analysis")
    competitor_data = pd.read_excel(file_path, sheet_name="Competitive_Positioning_Analysis")
    print("Datasets loaded successfully!")
except FileNotFoundError:
    print(f"ERROR: File not found at {file_path}. Please check the path.")
    exit()

# === Visualization: Profitability Analysis ===
plt.figure(figsize=(12, 6))
sns.barplot(data=profitability_data, x="Zone", y="Profit", hue="Product_Category", palette="coolwarm")
plt.title("Profitability by Zone and Product Category", fontsize=14)
plt.xlabel("Zone", fontsize=12)
plt.ylabel("Profit", fontsize=12)
plt.xticks(rotation=45)
plt.legend(title="Product Category")
plt.tight_layout()
plt.savefig("Profitability_Recommendations.png")
plt.close()
print("Profitability visualization saved as 'Profitability_Recommendations.png'.")

# === Visualization: Brand Popularity Analysis ===
plt.figure(figsize=(8, 6))
sns.heatmap(brand_popularity_data, annot=True, cmap="coolwarm")
plt.title("Correlation Matrix: Brand Popularity Insights", fontsize=14)
plt.tight_layout()
plt.savefig("Brand_Popularity_Recommendations.png")
plt.close()
print("Brand popularity visualization saved as 'Brand_Popularity_Recommendations.png'.")

# === Visualization: Competitive Positioning Analysis ===
plt.figure(figsize=(12, 6))
sns.scatterplot(data=competitor_data, x="Competitor_Price", y="Competitor_Market_Share",
                hue="Product_Category", style="Competitor_Name", palette="viridis", s=100)
plt.title("Competitive Positioning: Price vs Market Share", fontsize=14)
plt.xlabel("Competitor Price", fontsize=12)
plt.ylabel("Market Share (%)", fontsize=12)
plt.legend(title="Product Category & Competitor")
plt.tight_layout()
plt.savefig("Competitive_Positioning_Recommendations.png")
plt.close()
print("Competitive positioning visualization saved as 'Competitive_Positioning_Recommendations.png'.")
