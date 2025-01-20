import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load cleaned dataset
file_path = "C:/Users/ASUS/Downloads/Cleaned_Command_Centre_Dataset.xlsx"

# Load the competitor dataset
competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")

# Rename Competitor_Product_Category to Product_Category for consistency
if "Competitor_Product_Category" in competitor_data.columns:
    competitor_data.rename(columns={"Competitor_Product_Category": "Product_Category"}, inplace=True)

# Group data for analysis
competitor_analysis = competitor_data.groupby(["Product_Category", "Competitor_Name"])[
    ["Competitor_Price", "Competitor_Market_Share"]
].mean().reset_index()

# Create the scatter plot
plt.figure(figsize=(10, 6))
sns.scatterplot(
    data=competitor_analysis,
    x="Competitor_Price",
    y="Competitor_Market_Share",
    hue="Product_Category",
    style="Competitor_Name",
    palette="coolwarm",
    s=100
)

# Add plot details
plt.title("Competitive Positioning: Price vs. Market Share", fontsize=14)
plt.xlabel("Competitor Price", fontsize=12)
plt.ylabel("Market Share (%)", fontsize=12)
plt.legend(title="Product Category", fontsize=10, loc="best")
plt.tight_layout()

# Save and show the plot
plt.savefig("Competitive_Positioning_Analysis.png")
plt.show()

print("\nINFO: Visualization saved as 'Competitive_Positioning_Analysis.png'")
