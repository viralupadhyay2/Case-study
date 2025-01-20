import pandas as pd

# File path for the Excel file
file_path = "C:/Users/ASUS/Downloads/Final_Command_Centre_Analysis.xlsx"

# Load datasets
try:
    # Load profitability analysis
    profitability = pd.read_excel(file_path, sheet_name="Profitability_Analysis")
    
    # Load brand popularity analysis
    brand_popularity = pd.read_excel(file_path, sheet_name="Brand_Popularity_Analysis", index_col=0)
    
    # Load competitive positioning analysis
    competitive_positioning = pd.read_excel(file_path, sheet_name="Competitive_Positioning_Analysis")
    
    print("Datasets loaded successfully!")
except FileNotFoundError:
    print(f"Error: File not found at {file_path}. Please check the path.")
    exit()

# Generate recommendations
recommendations = []

# Profitability Recommendations
most_profitable = profitability.sort_values(by="Profit", ascending=False).head(1)
least_profitable = profitability.sort_values(by="Profit", ascending=True).head(1)

recommendations.append("Profitability Insights:")
recommendations.append(f" - Focus on the most profitable zone and product category: {most_profitable.iloc[0].to_dict()}.")
recommendations.append(f" - Address the least profitable zone and product category: {least_profitable.iloc[0].to_dict()}.")

# Brand Popularity Recommendations
high_correlation = brand_popularity.abs().unstack().sort_values(ascending=False).iloc[1:2]
correlation_key = high_correlation.index[0]
recommendations.append("Brand Popularity Insights:")
recommendations.append(f" - Focus on improving customer ratings and feedback volume as they show a high correlation ({correlation_key[0]} vs {correlation_key[1]}).")

# Competitive Positioning Recommendations
top_competitor = competitive_positioning.sort_values(by="Competitor_Market_Share", ascending=False).head(1)
recommendations.append("Competitive Positioning Insights:")
recommendations.append(f" - Analyze and emulate the pricing strategy of the top competitor: {top_competitor.iloc[0].to_dict()}.")

# Save recommendations to a text file
recommendations_file_path = "C:/Users/ASUS/Downloads/Final_Recommendations_Report.txt"

with open(recommendations_file_path, "w") as file:
    file.write("\n=== Recommendations ===\n\n")
    file.write("\n".join(recommendations))

print("\nINFO: Recommendations saved successfully to Final_Recommendations_Report.txt")
