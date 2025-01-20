import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load cleaned dataset
file_path = "C:/Users/ASUS/Downloads/Cleaned_Command_Centre_Dataset.xlsx"

# Load relevant data for Brand Popularity Analysis
customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")

# Merge Customer Feedback Data and Marketing Data on relevant keys
brand_popularity_data = pd.merge(
    customer_feedback_data,
    marketing_data,
    on=["Country", "Zone", "Date"],
    how="inner"
)[["Brand_Awareness_Score", "Customer_Rating", "Volume_of_Feedback"]]

# Drop rows with missing values for the analysis
brand_popularity_data = brand_popularity_data.dropna()

# Generate a heatmap for the correlation matrix
plt.figure(figsize=(8, 6))
correlation_matrix = brand_popularity_data.corr()
sns.heatmap(correlation_matrix, annot=True, cmap="coolwarm", fmt=".2f")
plt.title("Correlation Matrix: Brand Popularity Analysis", fontsize=14)
plt.tight_layout()
plt.savefig("Brand_Popularity_Correlation_Heatmap.png")
plt.show()

# Generate individual scatter plots for key relationships
plt.figure(figsize=(8, 6))
sns.scatterplot(
    data=brand_popularity_data, x="Brand_Awareness_Score", y="Customer_Rating", hue="Volume_of_Feedback", palette="viridis"
)
plt.title("Brand Awareness Score vs Customer Rating", fontsize=14)
plt.xlabel("Brand Awareness Score", fontsize=12)
plt.ylabel("Customer Rating", fontsize=12)
plt.colorbar(label="Volume of Feedback")
plt.tight_layout()
plt.savefig("Brand_Awareness_vs_Customer_Rating.png")
plt.show()
