import pandas as pd
import numpy as np
import re
import argparse

def load_data(file_path):
    """
    Load the Excel file and preprocess the data.
    
    :param file_path: Path to the Excel file.
    :return: Tuple containing choice values and strategy DataFrame.
    """
    df = pd.read_excel(file_path, header=None)
    val = df[0].isna().idxmax()
    choice_vals = df.head(val).set_index(0)[1]
    strat_df = df[val+1:]
    strat_df.columns = strat_df.iloc[0]
    strat_df = strat_df[1:]
    strat_df = strat_df.set_index('Strategy Name')
    return choice_vals, strat_df

def calculate_losses(choice_vals):
    """
    Extract loss names, find numeric values, and map to corresponding keys.
    
    :param choice_vals: DataFrame with choice values.
    :return: Dictionary of loss names to numeric values.
    """
    sd_loss_names = [item for item in choice_vals.index if 'sd' in item.lower()]
    sd_losses = [float(re.findall(r"\d+\.\d+|\d+", s)[0]) for s in sd_loss_names]
    return dict(zip(sd_loss_names, sd_losses))

def compute_financial_metrics(choice_vals, strat_df, loss_dict):
    """
    Compute the various financial metrics based on strategy and choice values, organized by calculation type.
    
    :param choice_vals: DataFrame with choice values.
    :param strat_df: DataFrame with strategy data.
    :param loss_dict: Dictionary mapping loss types to values.
    :return: DataFrame with computed financial metrics.
    """
    strat_score = strat_df[strat_df.columns[1:5]].mean(axis=1)
    risk_adj_allocation = choice_vals['AUM'] * strat_df['Allocation%'] * strat_score

    calc_data = {
        'strat_score': strat_score,
        'risk adj allocation': risk_adj_allocation
    }
    
    # Initialize dictionaries for each calculation type
    limit_loss_data = {}
    sd_move_data = {}
    sd_vol_move_data = {}
    
    for loss, value in loss_dict.items():
        # Populate each calculation type with respective 'sd' values
        limit_loss_data[f"{value} SD limit loss"] = choice_vals[loss] * risk_adj_allocation
        sd_move_data[f"{value} SD move"] = value * strat_df['30d Vol'] / np.sqrt(strat_df['Day Convention'].astype(int))
        sd_vol_move_data[f"{value} SD vol move"] = value * choice_vals['Vol Factor'] * strat_df['30d Vol']

    # Merge calculation groups into main dictionary
    calc_data.update(limit_loss_data)
    calc_data.update(sd_move_data)
    calc_data.update(sd_vol_move_data)

    calc_data['Max Daily Theta'] = (risk_adj_allocation * choice_vals['Weekly Decay Loss'] / (strat_df['Day Convention'] / 52)).astype(float)
    return pd.DataFrame.from_dict(calc_data, orient='index').T.astype(float).round(2)

def main(file_path):
    """
    Main function to run the financial data processing.
    
    :param file_path: Path to the Excel file input.
    """
    choice_vals, strat_df = load_data(file_path)
    loss_dict = calculate_losses(choice_vals)
    calc_df = compute_financial_metrics(choice_vals, strat_df, loss_dict)
    calc_df.to_excel('output.xlsx', index=True)
    print('Data has been written to output.xlsx')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process the Excel file for financial calculations.')
    parser.add_argument('file_path', type=str, help='Path to the Excel file')
    args = parser.parse_args()
    main(args.file_path)
