import os
import pandas as pd

def combine_excel_files(path):
    all_dataframes = []

    for filename in os.listdir(path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            df = pd.read_excel(os.path.join(path, filename))
            all_dataframes.append(df)

    combined_df = pd.concat(all_dataframes, ignore_index=True)

    combined_df.to_excel("combined.xlsx", index=False)

if __name__ == "__main__":
    path = input("Enter the path to the directory containing the Excel files: ")
    combine_excel_files(path)
