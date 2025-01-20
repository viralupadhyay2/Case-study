# brand_popularity_analysis.py
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load cleaned data
file_path = "Cleaned_Command_Centre_Dataset.xlsx"
customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")

# Merge relevant columns
brand_data = pd.merge(
    marketing_data[["Country", "Brand_Awareness_Score", "Customer_Reach"]],
    customer_feedback_data[["Country", "Customer_Rating", "Volume_of_Feedback"]],
    on="Country"
)

# Visualize Correlation Matrix
plt.figure(figsize=(8, 6))
sns.pairplot(brand_data, diag_kind="kde")
plt.title("Brand Popularity Analysis")
plt.savefig("Brand_Popularity_Analysis.png")
plt.show()
