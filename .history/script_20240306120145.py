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
    # Define separate formulas for each revenue type using actual column names
    revenue_formulas = {
        'Day-ahead revenues (€)': ['Day-ahead revenues (€) + Day-ahead revenues (€)']
        # Add formulas for other revenue types using actual column names
    }

    # Create a copy of the dataframe to avoid modifying the original
    calculated_df = df.copy()

    # Apply the formulas for each revenue type
    for revenue, formula in revenue_formulas.items():
        calculated_df[revenue] = df[formula[0]] + df[formula[1]]

    # Return only the calculated revenue columns
    calculated_revenues = calculated_df[['Scenario'] + list(revenue_formulas.keys())]

    return calculated_revenues

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

    # Use the scenario name as the output file name
    output_filename = f"{scenario_name}_comparison.xlsx"
    
    # Write to Excel with two sheets: scenario_name and scenario_name_comparison
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, sheet_name=scenario_name, index=False)
        revenues.to_excel(writer, sheet_name=f"{scenario_name}_comparison", index=False)

        # Add original columns from overall_results.xlsx to the scenario_name_comparison sheet
        overall_columns = combined_df.filter(regex='_overall_results.xlsx$', axis=1)
        overall_columns['Scenario'] = scenario_name
        overall_columns.to_excel(writer, sheet_name=f"{scenario_name}_comparison", startrow=len(revenues)+3, index=False)

if __name__ == "__main__":
    path = input("Enter the path to the directory containing the Excel files: ")
    combine_excel_files(path)
