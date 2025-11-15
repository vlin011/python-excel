import pandas as pd
import glob
import os

def process_file(filepath):
    # Read the sheet as text (since it's not a standard table)
    df_raw = pd.read_excel(filepath, sheet_name="Sheet1", header=None)
    lines = df_raw[0].astype(str).tolist()
    
    records = []
    i = 0
    while i < len(lines):
        # Detect company line
        if "COMPANY" in lines[i]:
            company = lines[i].split('-')[-1].strip()
            i += 1
        else:
            company = None
        
        # Detect contract
        if i < len(lines) and "Contract" in lines[i]:
            i += 1
            contract = lines[i].strip()
            i += 1
        else:
            contract = None
        
        # Detect fuel type
        if i < len(lines) and "Fuel Type" in lines[i]:
            i += 1
            fuel_type = lines[i].strip()
            i += 1
        else:
            fuel_type = None
        
        # Detect month block
        if i < len(lines) and "MONTH" in lines[i]:
            i += 1  # skip "MONTH"
            # Next lines: GALLONS, GROSS
            if "GALLONS" in lines[i]:
                i += 1
            if "GROSS" in lines[i]:
                i += 1
            # Now, read month blocks until next contract/company or end
            while i + 2 < len(lines) and not any(x in lines[i] for x in ["Contract", "COMPANY"]):
                month = lines[i].strip()
                gallons = lines[i+1].strip()
                gross = lines[i+2].strip()
                records.append({
                    "Company": company,
                    "Contract": contract,
                    "Fuel Type": fuel_type,
                    "Month": month,
                    "Gallons": gallons,
                    "Gross": gross
                })
                i += 3
        else:
            i += 1  # skip unknown line
    
    return pd.DataFrame(records)

# Main script
data_folder = os.path.join(os.getcwd(), "data")
excel_files = glob.glob(os.path.join(data_folder, "*.xlsx"))

all_data = []
for file in excel_files:
    print(f"Processing {file}")
    try:
        df = process_file(file)
        all_data.append(df)
    except Exception as e:
        print(f"Error processing {file}: {e}")

if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    output_file = os.path.abspath("combined_columnar_output.xlsx")
    final_df.to_excel(output_file, index=False)
    print(f"✅ Data transformed and saved to {output_file}")
else:
    print("❌ No data processed. Check your 'data' folder and file structure.")

