# Import necessary libraries
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# File path for the dataset
file_path = "C:/Users/ASUS/Downloads/Command Centre Dataset (1).xlsx"

# Load datasets
datasets = {
    "Sales_Data": pd.read_excel(file_path, sheet_name="Sales_Data"),
    "Marketing_Data": pd.read_excel(file_path, sheet_name="Marketing_Data"),
    "Customer_Feedback_Data": pd.read_excel(file_path, sheet_name="Customer_Feedback_Data"),
    "Competitor_Data": pd.read_excel(file_path, sheet_name="Competitor_Data"),
}

# Function to handle missing, inconsistent, or duplicate data
def clean_data(df, dataset_name):
    print(f"\n--- Cleaning {dataset_name} ---")

    # Handle missing values
    missing_values_before = df.isnull().sum()
    df.fillna("Unknown", inplace=True)
    print(f"Missing values before: \n{missing_values_before}")
    print(f"Missing values after: \n{df.isnull().sum()}")

    # Handle duplicate data
    duplicates_before = df.duplicated().sum()
    df.drop_duplicates(inplace=True)
    print(f"Duplicates removed: {duplicates_before}")

    return df

# Clean each dataset
for name, dataset in datasets.items():
    datasets[name] = clean_data(dataset, name)

# Standardize date formats and country names
def standardize_formats(df, columns, dataset_name):
    print(f"\n--- Standardizing {dataset_name} ---")

    if "Date" in columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        print(f"Date format standardized for {dataset_name}.")

    if "Country" in columns:
        country_mapping = {
            "US": "USA",
            "United States": "USA",
            "UK": "United Kingdom",
            "England": "United Kingdom",
            # Add other mappings as necessary
        }
        df["Country"] = df["Country"].replace(country_mapping)
        print(f"Country names standardized for {dataset_name}.")

    return df

# Standardize each dataset
for name, dataset in datasets.items():
    datasets[name] = standardize_formats(dataset, dataset.columns, name)

print("\nINFO: Data preparation completed successfully.")

# Save cleaned datasets
output_path = "C:/Users/ASUS/Downloads/Cleaned_Command_Centre_Dataset.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for name, dataset in datasets.items():
        dataset.to_excel(writer, sheet_name=name, index=False)

print(f"\nCleaned datasets saved to {output_path}")
