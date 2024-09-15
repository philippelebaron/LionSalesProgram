import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import ticker
import statistics


def set_df(file, sheet):
    df = pd.read_excel(file, sheet_name=sheet)
    return df


def get_sheets(file):
    return pd.ExcelFile(file).sheet_names


# get totals
def get_total_production(df):
    column = df.iloc[:, 3]
    total_production = 0
    for sum in column:
        # Calculate total production
        total_production += sum
    return total_production
    

def calculate_8020_production_stats(df, name):
    # Returns a tuple containing the row number of the last A player, 
    # the fraction of the producers, and the fractio of the production 
    # they produce according to the 80/20 rule
    cumulative_sum = 0.0
    row_num = 0.0
    global leg_title
    leg_title = df.columns[0]
    order_dict = df.iloc[:, 3]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)
    df_total_prod = get_total_production(df)
    df_len = len(df.index)
    for sale in sorted_values:
        row_num += 1
        cumulative_sum += sale[1]
        if ((row_num / len(df.index)) + (cumulative_sum / df_total_prod)) >= 1:
            list_indices = []
            for num in range(0, int(row_num)):
                list_indices.append(sorted_values[num][0])
            new_df = df.iloc[list_indices,:]
            new_df_total_prod = get_total_production(new_df)
            data = {"Name": [name],
                    "A Producers": [int(row_num)], 
                    "B Producers": [int(df_len - row_num)],
                    "Total Producers": [int(df_len)],
                    "Proportion of Producers %": [str(round(100 * (row_num / df_len)))], 
                    "Proportion of Production %": [str(round(100 * (cumulative_sum / df_total_prod)))],
                    "A Production": [round(new_df_total_prod)],
                    "B Production": [round(df_total_prod - new_df_total_prod)],
                    "Total Production": [round(df_total_prod)],
                    "Average Production (A)": [round(new_df_total_prod / len(new_df.index))],
                    "Average Production (B)": [round((df_total_prod - new_df_total_prod) / (df_len - len(new_df.index)))],
                    "Average Production (All)": [round(df_total_prod / df_len)]}
            dataframe = pd.DataFrame(data)
            prod_df = dataframe.transpose().reset_index()
            prod_df.columns = ['New Column'] + list(prod_df.columns[1:])
            prod_df.columns = prod_df.iloc[0]
            prod_df = prod_df.iloc[1:].reset_index(drop=True)
            return prod_df

