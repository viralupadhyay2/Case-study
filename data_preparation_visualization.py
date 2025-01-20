import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load Excel file
file_path = "C:/Users/ASUS/Downloads/Command Centre Dataset.xlsx"

# Load datasets
sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")
marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")
customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
competitor_data = pd.read_excel(file_path, sheet_name="Competitor_Data")

# Function to visualize missing values
def visualize_missing_values(data, dataset_name):
    missing_summary = data.isnull().sum()
    print(f"--- Missing Values in {dataset_name} ---")
    print(missing_summary)
    plt.figure(figsize=(8, 6))
    sns.heatmap(data.isnull(), cbar=False, cmap="viridis")
    plt.title(f"Missing Values Heatmap: {dataset_name}", fontsize=14)
    plt.tight_layout()
    plt.savefig(f"{dataset_name}_Missing_Values_Heatmap.png")
    plt.show()

# Visualize missing values for each dataset
visualize_missing_values(sales_data, "Sales_Data")
visualize_missing_values(marketing_data, "Marketing_Data")
visualize_missing_values(customer_feedback_data, "Customer_Feedback_Data")
visualize_missing_values(competitor_data, "Competitor_Data")

# Address missing values
sales_data.fillna({'Country': 'Unknown', 'Zone': 'Unknown', 'Revenue': 0, 'Cost_of_Goods_Sold': 0}, inplace=True)
marketing_data.fillna({'Country': 'Unknown', 'Zone': 'Unknown', 'Marketing_Spend': 0}, inplace=True)
customer_feedback_data.fillna({'Country': 'Unknown', 'Zone': 'Unknown', 'Customer_Rating': 'No Rating'}, inplace=True)
competitor_data.fillna({'Country': 'Unknown', 'Zone': 'Unknown', 'Competitor_Price': 0}, inplace=True)

# Function to visualize duplicates
def visualize_duplicates(data, dataset_name):
    duplicates = data.duplicated().sum()
    print(f"--- Duplicates in {dataset_name}: {duplicates} ---")
    if duplicates > 0:
        plt.figure(figsize=(8, 4))
        sns.barplot(x=["Unique Rows", "Duplicates"], y=[len(data) - duplicates, duplicates], palette="coolwarm")
        plt.title(f"Duplicate Summary: {dataset_name}", fontsize=14)
        plt.ylabel("Count", fontsize=12)
        plt.tight_layout()
        plt.savefig(f"{dataset_name}_Duplicate_Summary.png")
        plt.show()

# Visualize duplicates for each dataset
visualize_duplicates(sales_data, "Sales_Data")
visualize_duplicates(marketing_data, "Marketing_Data")
visualize_duplicates(customer_feedback_data, "Customer_Feedback_Data")
visualize_duplicates(competitor_data, "Competitor_Data")

# Remove duplicates
sales_data = sales_data.drop_duplicates()
marketing_data = marketing_data.drop_duplicates()
customer_feedback_data = customer_feedback_data.drop_duplicates()
competitor_data = competitor_data.drop_duplicates()

print("\nINFO: Missing values handled and duplicates removed.")
