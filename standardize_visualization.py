# Import necessary libraries
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load Excel file
file_path = "C:/Users/ASUS/Downloads/Command Centre Dataset.xlsx"

# Load each dataset
try:
    sales_data = pd.read_excel(file_path, sheet_name="Sales_Data")
    marketing_data = pd.read_excel(file_path, sheet_name="Marketing_Data")
    customer_feedback_data = pd.read_excel(file_path, sheet_name="Customer_Feedback_Data")
    print("Datasets loaded successfully!")
except FileNotFoundError:
    print("File not found. Please check the file path.")
    exit()

# === Standardize Date Formats ===
datasets = {
    "Sales_Data": sales_data,
    "Marketing_Data": marketing_data,
    "Customer_Feedback_Data": customer_feedback_data,
}

# Standardize date formats
for name, dataset in datasets.items():
    dataset['Date'] = pd.to_datetime(dataset['Date'], errors='coerce')
    print(f"{name} - Date Format Standardized")

# Visualize the date distributions
for name, dataset in datasets.items():
    plt.figure(figsize=(8, 5))
    dataset['Year'] = dataset['Date'].dt.year
    sns.countplot(data=dataset, x="Year", palette="viridis")
    plt.title(f"Year Distribution in {name}")
    plt.xlabel("Year")
    plt.ylabel("Frequency")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{name}_Date_Distribution.png")
    plt.close()

# === Standardize Country Names ===
# Define country name mappings for standardization
country_mapping = {
    "US": "USA",
    "United States": "USA",
    "UK": "United Kingdom",
    "England": "United Kingdom",
    # Add additional mappings if needed
}

# Standardize country names
for name, dataset in datasets.items():
    dataset['Country'] = dataset['Country'].replace(country_mapping)
    print(f"{name} - Country Names Standardized")

# Visualize the standardized country names
for name, dataset in datasets.items():
    plt.figure(figsize=(8, 5))
    sns.countplot(data=dataset, y="Country", order=dataset['Country'].value_counts().index, palette="coolwarm")
    plt.title(f"Country Distribution in {name} (Standardized)")
    plt.xlabel("Frequency")
    plt.ylabel("Country")
    plt.tight_layout()
    plt.savefig(f"{name}_Country_Distribution.png")
    plt.close()

print("INFO: Date and country standardization completed. Visualizations saved as PNG files.")
