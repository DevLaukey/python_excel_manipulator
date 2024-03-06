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
        'Day-ahead revenues (€)': ['p_annonce_day_ahead_MW * real_prices_day_ahead_euros_per_MWh'],
        'primary up band revenues (€)': ['p_annonce_reserve_MW_primary_up * real_prices_band_euros_per_MW_primary_up'],
        'primary down band revenues (€)': ['p_annonce_reserve_MW_primary_down * real_prices_band_euros_per_MW_primary_down'],
        'secondary up band revenues (€)': ['p_annonce_reserve_MW_secondary_up * real_prices_band_euros_per_MW_secondary_up'],
        'secondary down band revenues (€)': ['p_annonce_reserve_MW_secondary_down * real_prices_band_euros_per_MW_secondary_down'],
        'secondary up reserve energy revenues (€)':['e_injection_reserve_MWh_secondary_up * real_prices_energy_reserve_euros_per_MWh_secondary_up'],
        'secondary down reserve energy revenues (€)':
            ['e_injection_reserve_MWh_secondary_down * real_prices_energy_reserve_euros_per_MWh_secondary_down'],
            'Energy lack balancing revenues (€)':['e_balancing_overall_lack_MWh * real_prices_energy_lack_balancing_euros_per_MWh'],
            'Energy surplus balancing revenues (€)':['e_balancing_overall_surplus_MWh * real_prices_energy_surplus_balancing_euros_per_MWh'],
            'secondary up band balancing cost (€)': ['p_balancing_reserve_MW_secondary_up_band_na * real_prices_band_euros_per_MW_secondary_up'],
            'secondary down band balancing cost (€)': ['p_balancing_reserve_MW_secondary_down_band_na * real_prices_band_euros_per_MW_secondary_down'],
    }

    # Create a copy of the dataframe to avoid modifying the original
    calculated_df = df.copy()

    # Apply the formulas for each revenue type
    for revenue, formula in revenue_formulas.items():
        calculated_df[revenue] = eval(formula[0], {'p_annonce_day_ahead_MW': df['p_annonce_day_ahead_MW'],
                                                   'real_prices_day_ahead_euros_per_MWh': df['real_prices_day_ahead_euros_per_MWh'],
                                                 })

    # Return only the calculated revenue columns
    calculated_revenues = calculated_df[['Scenario'] + list(revenue_formulas.keys())]

    return calculated_revenues

def create_calculations_table(calculated_revenues):
    # Create a table for calculations in the specified format
    calculations_table = pd.DataFrame({
        'CALCULATIONS': list(calculated_revenues.columns)[1:],  # Exclude the 'Scenario' column
        'Value': calculated_revenues.sum(axis=0)[1:]  # Sum values for each revenue type
    })

    return calculations_table

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

    # Create a table for calculations
    calculations_table = create_calculations_table(revenues)

    # Use the scenario name as the output file name
    output_filename = f"{scenario_name}_comparison.xlsx"

    # Write to Excel with two sheets: scenario_name and scenario_name_comparison
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        # Add the first sheet with the values
        combined_df.to_excel(writer, sheet_name=scenario_name, index=False)
        
        # Add the second sheet with the calculations table
        calculations_table.to_excel(writer, sheet_name=f"{scenario_name}_comparison", startrow=1, index=False)

if __name__ == "__main__":
    path = input("Enter the path to the directory containing the Excel files: ")
    combine_excel_files(path)
