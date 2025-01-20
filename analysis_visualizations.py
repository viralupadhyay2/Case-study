import pandas as pd

# Load cleaned dataset
file_path = "C:/Users/ASUS/Downloads/Cleaned_Command_Centre_Dataset.xlsx"

# Load the competitor dataset
competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")

# Display column names to check the dataset structure
print("Columns in Competitor_Data:", competitor_data.columns)

# Proceed with analysis based on available columns
if "Product_Category" in competitor_data.columns:
    competitor_analysis = competitor_data.groupby(["Product_Category", "Competitor_Name"])[
        ["Competitor_Price", "Competitor_Market_Share"]
    ].mean().reset_index()
else:
    print("Error: 'Product_Category' column is missing in Competitor_Data.")
    # Use alternate columns if needed
    competitor_analysis = competitor_data.groupby(["Competitor_Name"])[
        ["Competitor_Price", "Competitor_Market_Share"]
    ].mean().reset_index()
    competitor_analysis["Product_Category"] = "Unknown"  # Add a placeholder

print("Competitor Analysis Data:\n", competitor_analysis.head())
