import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = "C:\\Users\\I753931\\OneDrive - SAP SE\\Desktop\\Python\\Test\\projects_summary.xlsx"
xls = pd.ExcelFile(file_path)

# Reload the data with the correct header row (row index 1)
df_corrected = pd.read_excel(xls, sheet_name='SecuritySummary', header=1)

# Select relevant columns and drop empty rows
df_corrected = df_corrected[['Product Group', 'Component', 'Critical', 'High', 'Medium', 'Low']].dropna()

# Convert severity columns to numeric values
severity_cols = ['Critical', 'High', 'Medium', 'Low']
df_corrected[severity_cols] = df_corrected[severity_cols].apply(pd.to_numeric, errors='coerce')

# Summarize the total counts for each severity level
severity_summary_corrected = df_corrected[severity_cols].sum()

# Plot the corrected summary in a bar chart
plt.figure(figsize=(8, 5))
severity_summary_corrected.plot(kind='bar', color=['red', 'orange', 'yellow', 'green'])

plt.title("Summary of Security Vulnerabilities by Severity Level")
plt.xlabel("Severity Level")
plt.ylabel("Count of Issues")
plt.xticks(rotation=0)
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Show the corrected plot
plt.show()
