import os
import pandas as pd

def extract_scenario_name(filename):
    # Extract the scenario name from the input filename
    base_name = os.path.basename(filename)
    if "_operations_results_logs.xlsx" in base_name:
        return base_name.replace("_operations_results_logs.xlsx", "")
    elif "_overall_results.xlsx" in base_name:
        return base_name.replace("_overall_results.xlsx", "")
    else:
        return None

def calculate_revenues(df):
    # Calculate the sum of columns for each revenue type
    revenues = df.groupby('Scenario').sum()[['Day-ahead revenues (€)',
                                             'primary up band revenues (€)',
                                             'primary down band revenues (€)',
                                             'secondary up band revenues (€)',
                                             'secondary down band revenues (€)',
                                             'secondary up reserve energy revenues (€)',
                                             'secondary down reserve energy revenues (€)',
                                             'Overall balancing revenues (€)',
                                             'Energy lack balancing revenues (€)',
                                             'Energy surplus balancing revenues (€)',
                                             'secondary up band balancing cost (€)',
                                             'secondary down band balancing cost (€)']]
    return revenues

def combine_excel_files(path):
    all_dataframes = []

    for filename in os.listdir(path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            df = pd.read_excel(os.path.join(path, filename))

            # Extract scenario name from the filename
            scenario_name = extract_scenario_name(filename)

            if scenario_name is not None:
                # Add a new column with the scenario name
                df['Scenario'] = scenario_name
                all_dataframes.append(df)

    combined_df = pd.concat(all_dataframes, ignore_index=True)

    # Calculate revenues
    revenues = calculate_revenues(combined_df)

    # Create a new dataframe with the calculated revenues
    calculations_df = pd.DataFrame({'CALCULATIONS': list(revenues.index)})

    # Add the calculated values to the calculations_df
    for col in revenues.columns:
        calculations_df[col] = revenues[col].values

    # Use the scenario name as the output file name
    output_filename = f"{scenario_name}_comparison.xlsx"
    
    # Write to Excel with two sheets: scenario_name and scenario_name_comparison
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, sheet_name=scenario_name, index=False)
        calculations_df.to_excel(writer, sheet_name=f"{scenario_name}_comparison", index=False)

if __name__ == "__main__":
    path = input("Enter the path to the directory containing the Excel files: ")
    combine_excel_files(path)
