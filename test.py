
import pandas as pd
import matplotlib.pyplot as plt

# File path to CSV
file_path = "C:\\Users\\I753931\\OneDrive - SAP SE\\Desktop\\Python\\Test\\sdol005.csv"

# Loading the CSV file into a DataFrame
df = pd.read_csv(file_path)

# Converting the "Date" column to a proper date format so Python can understand # Review later***
df["Date"] = pd.to_datetime(df["Date"])

# Grouping the data by "Date" and adding up the "Violation Count" for each date
timeline_summary = df.groupby("Date")["Violation Count"].sum().reset_index()

# Createing a line chart to show how violations change over time
plt.figure(figsize=(12, 6))  # Set the size of the chart
plt.plot(timeline_summary["Date"], timeline_summary["Violation Count"], marker='o', linestyle='-', linewidth=2)

# Adding labels and title to the chart
plt.xlabel("Date")
plt.ylabel("Total Violation Count")
plt.title("Timeline Improvement Measure Chart")

# Rotating the x-axis labels 
plt.xticks(rotation=0)

# Add grid lines 
plt.grid(True)

# Display the chart
plt.show()
