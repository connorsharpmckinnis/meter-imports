import pandas as pd
import os

# Map Excel "Meter Size" values to asset system codes
SIZE_TO_CODE = {
    '3"': "300",
    '5/8"': "058",
    '1"': "100",
    '1-1/2"': "150",
    '2"': "200",
    '4"': "400",
    '6"': "600"
}

def map_meter_size(size_str):
    if pd.isna(size_str):
        return ""
    for key, val in SIZE_TO_CODE.items():
        if key in size_str:
            return val
    return ""

def transform_excel_to_csv(input_path: str, output_path: str):
    # Load Excel data
    df = pd.read_excel(input_path)

    # Create transformed DataFrame
    output = pd.DataFrame()
    output["MISER"] = df["Meter SN"].fillna(df["Register SN"].astype(str))
    output["MISERV"] = 3
    output["MITSTDT"] = df["Meter Test Date"].fillna(df["Ship Date"])
    output["MIMTYPE"] = "Sensus"
    output["MIPDTE"] = "01/01/0001"
    output["MIMULT"] = 1
    output["MIKIND"] = "W"
    output["MIDIAL"] = 9
    output["MICDTE"] = "01/01/0001"
    output["MIMTRS"] = "IPK"
    output["MIHANDHELD"] = "N"
    output["MIOPEN2"] = df["Meter Size"].apply(map_meter_size)

    # Fill in missing required columns with blanks
    TARGET_COLUMNS = [
        'MISER', 'MISERV', 'MITSTDT', 'MITSTR', 'MIASFKWH', 'MIASLKWH', 'MIASFFL', 'MIASFLL',
        'MIASLLL', 'MIASLFL', 'MIWHY', 'MIMOD#', 'MIMTYPE', 'MIVOU', 'MIPDTE', 'MIPCST',
        'MIMULT', 'MIMNO', 'MIMFG', 'MITYPE', 'MIKIND', 'MIPHSE', 'MIVOLT', 'MIAMPS',
        'MIWIRE', 'MIRACTION', 'MIDIAL', 'MIFORM', 'MIRR', 'MICLAS', 'MIPF', 'MIOSER',
        'MIOREAD', 'MIOKWRD', 'MICDTE', 'MIWHSE', 'MIMCT', 'MIMTRS', 'MIOPEN2', 'MIOPEN3',
        'MIHANDHELD'
    ]

    for col in TARGET_COLUMNS:
        if col not in output.columns:
            output[col] = ""

    # Reorder the columns to match the required format
    output = output[TARGET_COLUMNS]

    # Save to CSV
    output.to_csv(output_path, index=False, encoding="utf-8")
    print(f"✅ Output saved to: {output_path}")

def batch_process_folder(folder_path: str):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not files: 
        print("❌ No Excel files found in the specified folder.")
        return
    
    for file in files: 
        input_path = os.path.join(folder_path, file)
        output_base = os.path.splitext(file)[0]
        output_path = f"test_out/{output_base}_transformed.csv"
        transform_excel_to_csv(input_path, output_path)


if __name__ == "__main__":

    input_xlsx = "test_in"  # Change as needed
    #output_csv = "test_out/transformed_output.csv"

    batch_process_folder(input_xlsx)
    
    '''if not os.path.exists(input_xlsx):
        print(f"❌ Input file not found: {input_xlsx}")
    else:
        transform_excel_to_csv(input_xlsx, output_csv)'''