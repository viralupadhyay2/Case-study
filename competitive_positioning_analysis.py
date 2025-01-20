import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# File path to the Excel file
file_path = "C:/Users/ASUS/Downloads/Final_Command_Centre_Analysis.xlsx"

# Specify the sheet name for Competitive Positioning Analysis
sheet_name = "Competitive_Positioning_Analysis"

# Load the dataset
try:
    competitive_data = pd.read_excel(file_path, sheet_name=sheet_name)
    print(f"{sheet_name} loaded successfully.")
except FileNotFoundError:
    print(f"Error: File not found at {file_path}. Please check the path.")
    exit()

# Visualize Competitive Pricing vs. Market Share
plt.figure(figsize=(10, 6))
sns.scatterplot(
    data=competitive_data,
    x="Competitor_Price",
    y="Competitor_Market_Share",
    hue="Product_Category",
    style="Competitor_Name",
    palette="viridis",
    s=100
)
plt.title("Competitive Positioning: Price vs. Market Share", fontsize=14)
plt.xlabel("Competitor Price", fontsize=12)
plt.ylabel("Market Share (%)", fontsize=12)
plt.legend(title="Product Category", bbox_to_anchor=(1.05, 1), loc="upper left")
plt.tight_layout()

# Save and show the chart
output_path = "C:/Users/ASUS/Downloads/Competitive_Positioning_Scatterplot.png"
plt.savefig(output_path)
plt.show()

print(f"Scatterplot saved to: {output_path}")
