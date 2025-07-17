import pandas as pd
from pathlib import PureWindowsPath

# List of Excel files to combine
input_files = [
    PureWindowsPath(r"\\DEWDFG212.wdf.global.corp.sap\vul-scan\Reports\SPM\Scan_Results\priv\SPM-PRIV-20250416.xlsx"),
    PureWindowsPath(r"\\DEWDFG212.wdf.global.corp.sap\vul-scan\Reports\SPM-COM\Scan_Results\priv\SPM-COM-PRIV-20250416.xlsx"),
    PureWindowsPath(r"\\DEWDFG212.wdf.global.corp.sap\vul-scan\Reports\SPM-ICM\Scan_Results\priv\SPM-ICM-PRIV-20250416.xlsx"),
    #"\\DEWDFG212.wdf.global.corp.sap\vul-scan\Reports\SPM-COM\Scan_Results\priv\SPM-COM-PRIV-20250416.xlsx"
    #"\\DEWDFG212.wdf.global.corp.sap\vul-scan\Reports\SPM-ICM\Scan_Results\priv\SPM-ICM-PRIV-20250416.xlsx"
    #"C:\\Users\\I753931\\OneDrive - SAP SE\\Desktop\\Python\\Test\\SPM-COM-PRIV-20250402.xlsx",
    #"C:\\Users\\I753931\\OneDrive - SAP SE\\Desktop\\Python\\Test\\SPM-ICM-PRIV-20250402.xlsx",
    #"C:\\Users\\I753931\\OneDrive - SAP SE\\Desktop\\Python\\Test\\SPM-PRIV-20250402.xlsx"
]

# List to store DataFrames
combined_data = []

# Read each file (columns A to BJ)
for file in input_files:
    try:
        # Use usecols to limit to columns A (1) through BJ (62)
        df = pd.read_excel(file, engine='openpyxl', usecols="A:BJ")
        #df['Source File'] = file  # optional: mark the origin
        combined_data.append(df)
    except Exception as e:
        print(f"Error reading {file}: {e}")

# Combine all data
final_df = pd.concat(combined_data, ignore_index=True)

# Save to new Excel file
output_file = "SPM-Combined-EVM 8.xlsx"
try:
    final_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Combined file saved as: {output_file}")
except Exception as e:
    print(f"Error writing combined file: {e}")
