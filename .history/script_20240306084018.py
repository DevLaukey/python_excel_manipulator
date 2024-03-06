import os
import pandas as pd
import argparse

def combine_excel_files(path):
    all_dataframes = []

    for filename in os.listdir(path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            df = pd.read_excel(os.path.join(path, filename))
            all_dataframes.append(df)

    combined_df = pd.concat(all_dataframes, ignore_index=True)

    combined_df.to_excel("combined.xlsx", index=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Combine Excel files in a given directory')
    parser.add_argument('path', type=str, help='Path to the directory containing the Excel files')

    args = parser.parse_args()

    combine_excel_files(args.path)s