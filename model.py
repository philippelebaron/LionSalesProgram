import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib import ticker
import statistics
import openpyxl

summary_1 = None
summary_2 = None
last_prod_frame = None
leg_title = None

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
    

def get_total_transactions(df):
    total_transactions = 0
    for sum in df.iloc[:, 2]:
        # Calculate total sales
        total_transactions += sum
    return total_transactions


def calculate_8020_sales(df):
    # Returns dataframe containing A-players calculated via values in the 
    # "Distinct Count of Siebel Order Number" column
    cumulative_sum = 0.0
    row_num = 0.0
    order_dict = df.iloc[:, 2]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)
    new_df = {}
    list_indices = []
    total_transactions = get_total_transactions(df)
    df_len = len(df.index)
    for sale in sorted_values:
        row_num += 1
        cumulative_sum += sale[1]
        if ((row_num / df_len) + (cumulative_sum / total_transactions)) >= 1:
            for num in range(0, int(row_num)):
                list_indices.append(sorted_values[num][0])
            new_df = df.iloc[list_indices,:]
            new_df = new_df.drop(new_df.columns[[1, 3]], axis=1)
            new_df.iloc[:, 1] = new_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
            new_df.to_clipboard()
            global last_prod_frame
            last_prod_frame = new_df
            return new_df


def calculate_sales_stats(df, name):
    # Returns tuple containing: (Number of sellers, proportion of salesmen, proportion of sales)
    cumulative_sum = 0.0
    row_num = 0.0
    order_dict = df.iloc[:, 2]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)
    df_total_transactions = get_total_transactions(df)
    df_len = len(df.index)
    for sale in sorted_values:
        row_num += 1
        cumulative_sum += sale[1]
        if ((row_num / df_len) + (cumulative_sum / df_total_transactions)) >= 1:
            list_indices = []
            for num in range(0, int(row_num)):
                list_indices.append(sorted_values[num][0])
            new_df = df.iloc[list_indices,:]
            new_df_total_transactions = get_total_transactions(new_df)
            data = {"Name": [name],
                    "A Sellers": [f"{int(row_num):,d}"], 
                    "B Sellers": [f"{int(df_len - row_num):,d}"],
                    "Total Sellers": [f"{int(df_len):,d}"],
                    "Proportion of Sellers %": [str(round(100 * (row_num / df_len)))], 
                    "Proportion of Transactions %": [str(round(100 * (cumulative_sum / df_total_transactions)))],
                    "A Transactions": [f"{round(new_df_total_transactions):,d}"],
                    "B Transactions": [f"{round(df_total_transactions - new_df_total_transactions):,d}"],
                    "Total Transactions": [f"{round(df_total_transactions):,d}"],
                    "Average Transactions (A)": [f"{round(new_df_total_transactions / len(new_df.index)):,d}"],
                    "Average Transactions (B)": [f"{round((df_total_transactions - new_df_total_transactions) / (df_len - len(new_df.index))):,d}"],
                    "Average Transactions (All)": [f"{round(df_total_transactions / df_len):,d}"]}
            dataframe = pd.DataFrame(data)
            return dataframe


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
            return dataframe


def calculate_8020_production_df(df):
    cumulative_sum = 0.0
    row_num = 0.0
    order_dict = df.iloc[:, 3]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)
    new_df = {}
    list_indices = []
    df_total_prod = get_total_production(df)
    df_len = len(df.index)
    for sale in sorted_values:
        row_num += 1
        cumulative_sum += sale[1]
        if ((row_num / df_len) + (cumulative_sum / df_total_prod)) >= 1:
            for num in range(0, int(row_num)):
                list_indices.append(sorted_values[num][0])
            new_df = df.iloc[list_indices,:]
            new_df = new_df.drop(new_df.columns[[1, 2]], axis=1)
            new_df.iloc[:, 1] = new_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
            new_df.to_clipboard()
            global last_prod_frame
            last_prod_frame = new_df
            return new_df


def bottom_20(df):
    cumulative_sum = 0.0
    row_num = 0.0
    order_dict = df.iloc[:, 3]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)
    new_df = {}
    list_indices = []
    df_len = len(df.index)
    eighty_pct = int(round(0.8 * df_len)) 
    for sale in sorted_values[eighty_pct:]:
        row_num += 1
        cumulative_sum += sale[1]
    for num in range(eighty_pct, df_len):
        list_indices.append(sorted_values[num][0])
    new_df = df.iloc[list_indices,:]
    new_df = new_df.drop(new_df.columns[[1, 2]], axis=1)
    new_df.iloc[:, 1] = new_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
    new_df.to_clipboard()
    global last_prod_frame
    last_prod_frame = new_df
    return new_df


def bottom_20_stats(df, name):
    cumulative_sum = 0.0
    row_num = 0.0
    order_dict = df.iloc[:, 3]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)
    new_df = {}
    list_indices = []
    df_len = len(df.index)
    eighty_pct = int(round(0.8 * df_len)) 
    for sale in sorted_values[eighty_pct:]:
        row_num += 1
        cumulative_sum += sale[1]
    for num in range(eighty_pct, df_len):
        list_indices.append(sorted_values[num][0])
    new_df = df.iloc[list_indices,:]
    df_len = len(new_df)
    bottom_prod = get_total_production(new_df)
    data = {
        "Name": [name],
        "Bottom Producers": [f"{int(df_len):,d}"],
        "Proportion of Producers %": [str(round(100 * (df_len / len(df))))],
        "Proportion of Production %": [str(round(100 * (bottom_prod / get_total_production(df))))],
        "Bottom Production": [f"{int(round(bottom_prod)):,d}"],
        "Average Production": [f"{int(round(bottom_prod / df_len)):,d}"]
    }
    dataframe = pd.DataFrame(data)
    return dataframe


def calculate_static_production(df) -> dict:
    # Returns a dictionary containing the bottom row numbers of 
    # ABCD player ranges according to the given key 
    # (A = top 15%, B+ = next 10%, B = 30%, B- = 10%, C = 25%, D = bottom 10%)
    size = len(df.index)
    players = ["A", "B+", "B", "B-", "C", "D", "Total"]
    proportions = ["15%", "10%", "30%", "10%", "25%", "10%", "100%"]
    members = [np.ceil(0.15 * size), np.ceil(0.25 * size) - np.ceil(0.15 * size), np.ceil(0.55 * size) - np.ceil(0.25 * size), np.ceil(0.65 * size) - np.ceil(0.55 * size), np.ceil(0.9 * size) - np.ceil(0.65 * size), np.ceil(size) - np.ceil(0.9 * size), np.ceil(size)]
    member_cuttoff = [np.ceil(0.15 * size), np.ceil(0.25 * size), np.ceil(0.55 * size), np.ceil(0.65 * size), np.ceil(0.9 * size), np.ceil(size), np.ceil(size)]
    order_dict = df.iloc[:, 3]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)

    prod_loop_count = 0
    production = 0
    total_production = [0, 0, 0, 0, 0, 0, 0]
    cut = [0, 0, 0, 0, 0, 0, 0]
    for prod in sorted_values:
        # sum transactions
        production += prod[1]
        if prod_loop_count < member_cuttoff[0]:
            total_production[0] += prod[1]
        elif prod_loop_count < member_cuttoff[1]:
            total_production[1] += prod[1]
        elif prod_loop_count < member_cuttoff[2]:
            total_production[2] += prod[1]
        elif prod_loop_count < member_cuttoff[3]:
            total_production[3] += prod[1]
        elif prod_loop_count < member_cuttoff[4]:
            total_production[4] += prod[1]
        elif prod_loop_count < member_cuttoff[5]:
            total_production[5] += prod[1]
        else:
            total_production[6] += prod[1]

        # find cut
        if prod_loop_count == member_cuttoff[0] - 1:
            cut[0] = prod[1]
        elif prod_loop_count == member_cuttoff[1] - 1:
            cut[1] = prod[1]
        elif prod_loop_count == member_cuttoff[2] - 1:
            cut[2] = prod[1]
        elif prod_loop_count == member_cuttoff[3] - 1:
            cut[3] = prod[1]
        elif prod_loop_count == member_cuttoff[4] - 1:
            cut[4] = prod[1]
        elif prod_loop_count == member_cuttoff[5] - 1:
            cut[5] = prod[1]
        elif prod_loop_count == member_cuttoff[6] - 1:
            cut[6] = prod[1]

        prod_loop_count += 1
    


    pct_production = []
    average_production = []
    xy_ratio_total = []
    xy_ratio_avg = []
    
    total_production[6] += production
    for total in total_production:
        pct_production.append(round(100 * (total / production)))

    for x in range(0,7):
        average_production.append(round(total_production[x] / members[x]))

    for z in range(0, 5):
        try:
            xy_ratio_total.append(round(100 * (total_production[z] / total_production[z + 1])))
        except ZeroDivisionError:
            xy_ratio_total.append(999999)
    for w in range(0,5):
        try:
            xy_ratio_avg.append(round(100 * (average_production[w] / average_production[w + 1])))
        except ZeroDivisionError:
            xy_ratio_avg.append(999999)
    xy_ratio_avg.append(0)
    xy_ratio_total.append(0)
    xy_ratio_avg.append(0)
    xy_ratio_total.append(0)
    
    data = {"Tier": players,
            "Breakdown": proportions,
            "Members": np.array(members).astype(int),
            "Production Cut": np.array(cut).astype(int),
            "Production": np.array(total_production).astype(int),
            "% of Total Production": np.array(pct_production).astype(int),
            "Average Production": np.array(average_production).astype(int),
            "Production AvsBvsCvsD %": np.array(xy_ratio_total).astype(int),
            "Average Production AvsBvsCvsD %": np.array(xy_ratio_avg).astype(int)
            }
    dataframe = pd.DataFrame(data)
    dataframe['Production Cut'] = dataframe['Production Cut'].apply(lambda x : "{:,}".format(x))
    dataframe['Members'] = dataframe['Members'].apply(lambda x : "{:,}".format(x))
    dataframe['Production'] = dataframe['Production'].apply(lambda x : "{:,}".format(x))
    dataframe['Average Production'] = dataframe['Average Production'].apply(lambda x : "{:,}".format(x))
    dataframe['Production AvsBvsCvsD %'] = dataframe['Production AvsBvsCvsD %'].apply(lambda x : "{:,}".format(x))
    dataframe['Average Production AvsBvsCvsD %'] = dataframe['Average Production AvsBvsCvsD %'].apply(lambda x : "{:,}".format(x))
    return dataframe


def calculate_static_transactions(df):
    size = len(df.index)
    players = ["A", "B+", "B", "B-", "C", "D", "Total"]
    proportions = ["15%", "10%", "30%", "10%", "25%", "10%", "100%"]
    members = [np.ceil(0.15 * size), np.ceil(0.25 * size) - np.ceil(0.15 * size), np.ceil(0.55 * size) - np.ceil(0.25 * size), np.ceil(0.65 * size) - np.ceil(0.55 * size), np.ceil(0.9 * size) - np.ceil(0.65 * size), np.ceil(size) - np.ceil(0.9 * size), np.ceil(size)]
    member_cuttoff = [np.ceil(0.15 * size), np.ceil(0.25 * size), np.ceil(0.55 * size), np.ceil(0.65 * size), np.ceil(0.9 * size), np.ceil(size), np.ceil(size)]
    order_dict = df.iloc[:, 2]
    sorted_values = sorted(order_dict.items(), key=lambda x:x[1], reverse=True)

    transaction_loop_count = 0
    transactions = 0
    total_transactions = [0, 0, 0, 0, 0, 0, 0]
    cut = [0, 0, 0, 0, 0, 0, 0]
    for transaction in sorted_values:
        # sum transactions
        transactions += transaction[1]
        if transaction_loop_count < member_cuttoff[0]:
            total_transactions[0] += transaction[1]
        elif transaction_loop_count < member_cuttoff[1]:
            total_transactions[1] += transaction[1]
        elif transaction_loop_count < member_cuttoff[2]:
            total_transactions[2] += transaction[1]
        elif transaction_loop_count < member_cuttoff[3]:
            total_transactions[3] += transaction[1]
        elif transaction_loop_count < member_cuttoff[4]:
            total_transactions[4] += transaction[1]
        elif transaction_loop_count < member_cuttoff[5]:
            total_transactions[5] += transaction[1]
        else:
            total_transactions[6] += transaction[1]

        # find cut
        if transaction_loop_count == member_cuttoff[0] - 1:
            cut[0] = transaction[1]
        elif transaction_loop_count == member_cuttoff[1] - 1:
            cut[1] = transaction[1]
        elif transaction_loop_count == member_cuttoff[2] - 1:
            cut[2] = transaction[1]
        elif transaction_loop_count == member_cuttoff[3] - 1:
            cut[3] = transaction[1]
        elif transaction_loop_count == member_cuttoff[4] - 1:
            cut[4] = transaction[1]
        elif transaction_loop_count == member_cuttoff[5] - 1:
            cut[5] = transaction[1]
        elif transaction_loop_count == member_cuttoff[6] - 1:
            cut[6] = transaction[1]

        transaction_loop_count += 1
    


    pct_transaction = []
    average_transactions = []
    xy_ratio_total = []
    xy_ratio_avg = []
    
    total_transactions[6] += transactions
    for total in total_transactions:
        pct_transaction.append(round(100 * (total / transactions)))

    for x in range(0,7):
        average_transactions.append(round(total_transactions[x] / members[x]))

    for z in range(0, 5):
        try:
            xy_ratio_total.append(round(100 * (total_transactions[z] / total_transactions[z + 1])))
        except ZeroDivisionError:
            xy_ratio_total.append(999999)
    for w in range(0,5):
        try:
            xy_ratio_avg.append(round(100 * (average_transactions[w] / average_transactions[w + 1])))
        except ZeroDivisionError:
            xy_ratio_avg.append(999999)
    xy_ratio_avg.append(0)
    xy_ratio_total.append(0)
    xy_ratio_avg.append(0)
    xy_ratio_total.append(0)

    data = {"Transactions Cut": np.array(cut).astype(int),
            "Transactions": np.array(total_transactions).astype(int),
            "% of Total Transactions": np.array(pct_transaction).astype(int),
            "Average Transaction": np.array(average_transactions).astype(int),
            "Transactions AvsBvsCvsD %": np.array(xy_ratio_total).astype(int),
            "Average Transaction AvsBvsCvsD %": np.array(xy_ratio_avg).astype(int)
            }
    dataframe = pd.DataFrame(data)
    dataframe['Transactions Cut'] = dataframe['Transactions Cut'].apply(lambda x : "{:,}".format(x))
    dataframe['Transactions'] = dataframe['Transactions'].apply(lambda x : "{:,}".format(x))
    dataframe['Average Transaction'] = dataframe['Average Transaction'].apply(lambda x : "{:,}".format(x))
    dataframe['Transactions AvsBvsCvsD %'] = dataframe['Transactions AvsBvsCvsD %'].apply(lambda x : "{:,}".format(x))
    dataframe['Average Transaction AvsBvsCvsD %'] = dataframe['Average Transaction AvsBvsCvsD %'].apply(lambda x : "{:,}".format(x))
    return dataframe


def get_abcd_all(df):
    dfs = [calculate_static_production(df), calculate_static_transactions(df)]
    frame = pd.concat(dfs, axis=1)
    frame.to_clipboard(index=False)
    return frame


def summarize_data_prod(file, sheet_names):
    dfs = []
    for sheet in sheet_names:
        dfs.append(calculate_8020_production_stats(set_df(file, sheet), sheet))
    df = pd.concat(dfs, names=sheet_names, ignore_index=True)
    return df


def summarize_data_tran(file, sheet_names):
    dfs = []
    for sheet in sheet_names:
        dfs.append(calculate_sales_stats(set_df(file, sheet), sheet))
    df = pd.concat(dfs, names=sheet_names, ignore_index=True)
    return df


def summarize_data_all(file, sheet_names):
    
    prod = summarize_data_prod(file, sheet_names)
    global last_prod_frame
    last_prod_frame = prod
    tran = summarize_data_tran(file, sheet_names)
    complete_months = len(prod)
    
    plt.close("all")
    plt.style.use('seaborn-darkgrid')
    month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    figure, axis = plt.subplots(3, 3)
    figure.tight_layout(pad=0.25)

    if len(prod) <= 12:
        if summary_1 is not None:
            complete_months_1 = len(summary_1)
            axis[0, 0].plot(month_labels[:complete_months_1], summary_1['A Producers'].values, ".-", label="Year 1")
            axis[1, 0].plot(month_labels[:complete_months_1], summary_1['B Producers'].values, ".-")
            axis[2, 0].plot(month_labels[:complete_months_1], summary_1['Total Producers'].values, ".-")
            axis[0, 1].plot(month_labels[:complete_months_1], summary_1['A Production'].values, ".-")
            axis[1, 1].plot(month_labels[:complete_months_1], summary_1['B Production'].values, ".-")
            axis[2, 1].plot(month_labels[:complete_months_1], summary_1['Total Production'].values, ".-")
            axis[0, 2].plot(month_labels[:complete_months_1], summary_1['Average Production (A)'].values, ".-")
            axis[1, 2].plot(month_labels[:complete_months_1], summary_1['Average Production (B)'].values, ".-")
            axis[2, 2].plot(month_labels[:complete_months_1], summary_1['Average Production (All)'].values, ".-")

        if summary_2 is not None:
            complete_months_2 = len(summary_2)
            axis[0, 0].plot(month_labels[:complete_months_2], summary_2['A Producers'].values, ".-", label="Year 2")
            axis[1, 0].plot(month_labels[:complete_months_2], summary_2['B Producers'].values, ".-")
            axis[2, 0].plot(month_labels[:complete_months_2], summary_2['Total Producers'].values, ".-")
            axis[0, 1].plot(month_labels[:complete_months_2], summary_2['A Production'].values, ".-")
            axis[1, 1].plot(month_labels[:complete_months_2], summary_2['B Production'].values, ".-")
            axis[2, 1].plot(month_labels[:complete_months_2], summary_2['Total Production'].values, ".-")
            axis[0, 2].plot(month_labels[:complete_months_2], summary_2['Average Production (A)'].values, ".-")
            axis[1, 2].plot(month_labels[:complete_months_2], summary_2['Average Production (B)'].values, ".-")
            axis[2, 2].plot(month_labels[:complete_months_2], summary_2['Average Production (All)'].values, ".-")
        
        axis[0, 0].plot(month_labels[:complete_months], prod['A Producers'].values, ".-", label="Current")
        axis[0, 0].set_ylabel('A Producers')
        axis[0, 0].set_xticklabels(month_labels, rotation = 45)
        axis[0, 0].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[0, 0].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[1, 0].plot(month_labels[:complete_months], prod['B Producers'].values, ".-")
        axis[1, 0].set_ylabel('B Producers')
        axis[1, 0].set_xticklabels(month_labels, rotation = 45)
        axis[1, 0].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[1, 0].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[2, 0].plot(month_labels[:complete_months], prod['Total Producers'].values, ".-")
        axis[2, 0].set_ylabel('Total Producers')
        axis[2, 0].set_xticklabels(month_labels, rotation = 45)
        axis[2, 0].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[2, 0].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[0, 1].plot(month_labels[:complete_months], prod['A Production'].values, ".-")
        axis[0, 1].set_ylabel('A Production')
        axis[0, 1].set_xticklabels(month_labels, rotation = 45)
        axis[0, 1].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[0, 1].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[1, 1].plot(month_labels[:complete_months], prod['B Production'].values, ".-")
        axis[1, 1].set_ylabel('B Production')
        axis[1, 1].set_xticklabels(month_labels, rotation = 45)
        axis[1, 1].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[1, 1].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[2, 1].plot(month_labels[:complete_months], prod['Total Production'].values, ".-")
        axis[2, 1].set_ylabel('Total Production')
        axis[2, 1].set_xticklabels(month_labels, rotation = 45)
        axis[2, 1].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[2, 1].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[0, 2].plot(month_labels[:complete_months], prod['Average Production (A)'].values, ".-")
        axis[0, 2].set_ylabel('Average Production (A)')
        axis[0, 2].set_xticklabels(month_labels, rotation = 45)
        axis[0, 2].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[0, 2].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[1, 2].plot(month_labels[:complete_months], prod['Average Production (B)'].values, ".-")
        axis[1, 2].set_ylabel('Average Production (B)')
        axis[1, 2].set_xticklabels(month_labels, rotation = 45)
        axis[1, 2].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[1, 2].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        axis[2, 2].plot(month_labels[:complete_months], prod['Average Production (All)'].values, ".-")
        axis[2, 2].set_ylabel('Average Production (All)')
        axis[2, 2].set_xticklabels(month_labels, rotation = 45)
        axis[2, 2].set_yticklabels(month_labels, rotation = 45, fontsize=8)
        axis[2, 2].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, p: format(int(x), ',')))

        
        global leg_title
        axis[0, 0].legend(loc='upper left', fontsize=8, title="" + leg_title + " Production", title_fontsize=8)


    prod_df = prod.transpose().reset_index()
    prod_df.columns = ['New_Column'] + list(prod_df.columns[1:])
    prod_df.columns = prod_df.iloc[0]
    prod_df = prod_df.iloc[1:].reset_index(drop=True)

    tran_df = tran.transpose().reset_index()
    tran_df.columns = ['New_Column'] + list(tran_df.columns[1:])
    tran_df.columns = tran_df.iloc[0]
    tran_df = tran_df.iloc[1:].reset_index(drop=True)
    
    dfs = [prod_df, tran_df]
    frame = pd.concat(dfs)
    frame = frame.applymap(format_int_with_commas)
    frame.to_clipboard(index=False)
    
    return frame


def show_plot():
    plt.show()


def set_stor1():
    global summary_1
    summary_1 = last_prod_frame


def set_stor2():
    global summary_2
    summary_2 = last_prod_frame


def format_int_with_commas(x):
    if isinstance(x, int):
        return f"{x:,}"
    else:
        return x


def get_current_target():
    global last_prod_frame
    return (round(last_prod_frame['Total Production'].iloc[2] * 1.15))


def get_average_a_proportion():
    global last_prod_frame
    return round(statistics.mean(last_prod_frame['Proportion of Producers %'].astype(float)), 1)


def validate_simu_input():
    return (len(last_prod_frame) == 5)


def simulate(target, prop):
    # need to grab old sheet, summary is not valid input
    global last_prod_frame

    name = []
    a_producers = []
    b_producers = []
    tot_producers = []
    a_prop = []
    b_prop = []
    a_prod = []
    b_prod = []
    tot_prod = []
    avg_a = []
    avg_b = []
    avg_all = []

    # Simulation 1
    name.append("SIMU 1 (iso #Apl)")
    a_producers.append(last_prod_frame['A Producers'].iloc[2])
    b_producers.append(last_prod_frame['B Producers'].iloc[2])
    tot_producers.append(last_prod_frame['Total Producers'].iloc[2])
    tot_prod.append(target)
    a_prop.append(prop)
    b_prop.append(100 - prop)
    a_prod.append(round((b_prop[0] / 100) * tot_prod[0]))
    b_prod.append(round(target - a_prod[0]))
    avg_a.append(round(a_prod[0] / a_producers[0]))
    avg_b.append(round(b_prod[0] / b_producers[0]))
    avg_all.append(round(tot_prod[0] / tot_producers[0]))

    # Simulation 2
    name.append("SIMU 2 (H2 vs H1 #Apl)")
    tot_producers.append(round(last_prod_frame['Total Producers'].iloc[0] *
                         ((last_prod_frame['Total Producers'].iloc[2] / last_prod_frame['Total Producers'].iloc[1])
                         + (last_prod_frame['Total Producers'].iloc[4] / last_prod_frame['Total Producers'].iloc[3]))
                         / 2))
    tot_prod.append(target)
    b_producers.append(round(sum(last_prod_frame['Proportion of Production %'].astype(int)) * tot_producers[1] / 500))
    a_producers.append(round(tot_producers[1] - b_producers[1]))
    a_prop.append(round(100.0 * a_producers[1] / tot_producers[1], 1))
    b_prop.append(100 - a_prop[1])
    a_prod.append(a_prod[0])
    b_prod.append(tot_prod[1] - a_prod[1])
    avg_a.append(round(a_prod[1] / a_producers[1]))
    avg_b.append(round(b_prod[1] / b_producers[1]))
    avg_all.append(round(tot_prod[1] / tot_producers[1]))


    # Simulation 3
    name.append("SIMU 3 (iso Aprod avg)")
    a_prop.append(prop)
    b_prop.append(100 - a_prop[2])
    tot_prod.append(target)
    a_prod.append(round((b_prop[0] / 100) * tot_prod[0]))
    b_prod.append(round(target - a_prod[0]))
    avg_a.append(last_prod_frame['Average Production (A)'].iloc[2])
    avg_b.append(last_prod_frame['Average Production (B)'].iloc[2])
    avg_all.append(last_prod_frame['Average Production (All)'].iloc[2])
    a_producers.append(round(a_prod[2] / avg_a[2]))
    b_producers.append(round(b_prod[2] / avg_b[2]))
    tot_producers.append(round(tot_prod[2] / avg_all[2]))


    # Simulation 4
    name.append("SIMU 4 (H2 vs H1 Aprod)")
    avg_a.append(round(last_prod_frame['Average Production (A)'].iloc[0] *
                         ((last_prod_frame['Average Production (A)'].iloc[2] / last_prod_frame['Average Production (A)'].iloc[1])
                         + (last_prod_frame['Average Production (A)'].iloc[4] / last_prod_frame['Average Production (A)'].iloc[3]))
                         / 2))
    avg_b.append(round(last_prod_frame['Average Production (B)'].iloc[0] *
                         ((last_prod_frame['Average Production (B)'].iloc[2] / last_prod_frame['Average Production (B)'].iloc[1])
                         + (last_prod_frame['Average Production (B)'].iloc[4] / last_prod_frame['Average Production (B)'].iloc[3]))
                         / 2))
    avg_all.append(round(last_prod_frame['Average Production (All)'].iloc[0] *
                         ((last_prod_frame['Average Production (All)'].iloc[2] / last_prod_frame['Average Production (All)'].iloc[1])
                         + (last_prod_frame['Average Production (All)'].iloc[4] / last_prod_frame['Average Production (All)'].iloc[3]))
                         / 2))
    tot_prod.append(target)
    a_prod.append(round((b_prop[0] / 100) * tot_prod[0]))
    b_prod.append(round(target - a_prod[0]))
    a_producers.append(round(a_prod[3] / avg_a[3]))
    b_producers.append(round(b_prod[3] / avg_b[3]))
    tot_producers.append(round(tot_prod[3] / avg_all[3]))
    a_prop.append(prop)
    b_prop.append(100 - a_prop[3])
    


    data = {
        "": name,
        "A Producers": a_producers, 
        "B Producers": b_producers,
        "Total Producers": tot_producers,
        "Proportion of Producers %": a_prop, 
        "Proportion of Production %": b_prop,
        "A Production": a_prod,
        "B Production": b_prod,
        "Total Production": tot_prod,
        "Average Production (A)": avg_a,
        "Average Production (B)": avg_b,
        "Average Production (All)": avg_all
    }
    frame = pd.DataFrame(data)

    #frame.at[0, "Average Production (A)"] = '\x1B[1m' + frame["Average Production (A)"].iloc[0].astype(str)

    prod_df = frame.transpose().reset_index()
    prod_df.columns = ['New_Column'] + list(prod_df.columns[1:])
    prod_df.columns = prod_df.iloc[0]
    prod_df = prod_df.iloc[1:].reset_index(drop=True)
    prod_df = prod_df.applymap(format_int_with_commas)

    prod_df.to_clipboard(index=False)

    return prod_df

# copy frame 1 to new dataframe, bool for year 1 is true
# for frame 2: if name does not exist add to frame, if name does exist, bool for year 2 is true,
# for frame 3: same thing, 
# add counter for amount of years
# tim sort by years
# 
def get_repeats():
    global summary_1, summary_2, last_prod_frame
    if summary_1 is None or summary_2 is None or last_prod_frame is None:
        return "Cannot sort. Please refer to User Manual."
    if len(summary_1.columns) != 2 or len(summary_2.columns) != 2 or len(last_prod_frame.columns) != 2:
        return "Cannot sort. Please refer to User Manual."
    data = {
        "Name": [],
        "Year 1": [],
        "Year 2": [],
        "Current": [],
        "Count": []
    }
    sort_frame = pd.DataFrame(data)
    for og_index, name in enumerate(summary_1.iloc[:, 0].values):
        sort_frame.loc[len(sort_frame.index)] = [name, summary_1.iloc[og_index, 1], "", "", 1]
    for og_index, name in enumerate(summary_2.iloc[:, 0]):
        if name in sort_frame["Name"].values:
            index = sort_frame.index[sort_frame["Name"] == name]
            sort_frame.iloc[[index], [2]] = summary_2.iloc[og_index, 1]
            sort_frame.iloc[[index], [4]] = 2
        else:
            sort_frame.loc[len(sort_frame.index)] = [name, "", summary_2.iloc[og_index, 1], "", 1]
    for og_index, name in enumerate(last_prod_frame.iloc[:, 0]):
        if name in sort_frame["Name"].values:
            index = sort_frame.index[sort_frame["Name"] == name]
            sort_frame.iloc[[index], [3]] = last_prod_frame.iloc[og_index, 1]
            sort_frame.iloc[index, 4] = sort_frame.iloc[index, 4] + 1
        else:
            sort_frame.loc[len(sort_frame.index)] = [name, "", "", last_prod_frame.iloc[og_index, 1], 1]
    sort_frame = sort_frame.sort_values("Count", ascending=False)
    return sort_frame


def sort_repeats(pre_sorted_frame):
    if isinstance(pre_sorted_frame, str):
        return pd.DataFrame({"Error:": [pre_sorted_frame]})
    count_three = pre_sorted_frame.loc[pre_sorted_frame["Count"] == 3]
    count_two = pre_sorted_frame.loc[pre_sorted_frame["Count"] == 2]
    count_one = pre_sorted_frame.loc[pre_sorted_frame["Count"] == 1]

    count_two_yr_1_2 = count_two[count_two["Current"] == ""]
    count_two_yr_1_cur = count_two[count_two["Year 2"] == ""]
    count_two_yr_2_cur = count_two[count_two["Year 1"] == ""]

    count_one_yr_1 = count_one[count_one["Year 1"] != ""]
    count_one_yr_2 = count_one[count_one["Year 2"] != ""]
    count_one_yr_cur = count_one[count_one["Current"] != ""]

    dfs = [count_three, count_two_yr_1_2, count_two_yr_1_cur, count_two_yr_2_cur, count_one_yr_1, count_one_yr_2, count_one_yr_cur]
    sorted_frame = pd.concat(dfs)
    sorted_frame.to_clipboard(index=False)
    return sorted_frame

def clear_summary_1():
    global summary_1
    summary_1 = None
    return True

def clear_summary_2():
    global summary_2
    summary_2 = None
    return True

def get_b_production(df):
    # exclusion with a prod and full list
    full_df = df.drop(df.columns[[1, 2]], axis=1)
    full_df.iloc[:, 1] = full_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
    a_players = calculate_8020_production_df(df)
    b_players = full_df[~full_df.iloc[:, 0].isin(a_players.iloc[:, 0])]
    b_players.to_clipboard()
    global last_prod_frame
    last_prod_frame = b_players
    return b_players

def get_all_production(df):
    # return full list
    full_df = df.drop(df.columns[[1, 2]], axis=1)
    full_df.iloc[:, 1] = full_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
    global last_prod_frame
    last_prod_frame = full_df
    return full_df

def get_b_transaction(df):
    full_df = df.drop(df.columns[[1, 3]], axis=1)
    full_df.iloc[:, 1] = full_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
    a_players = calculate_8020_sales(df)
    b_players = full_df[~full_df.iloc[:, 0].isin(a_players.iloc[:, 0])]
    b_players = b_players.sort_values(by=b_players.columns[1], ascending=False)
    b_players.to_clipboard()
    global last_prod_frame
    last_prod_frame = b_players
    return b_players

def get_all_transaction(df):
    full_df = df.drop(df.columns[[1, 3]], axis=1)
    full_df.iloc[:, 1] = full_df.iloc[:, 1].astype(int).apply(lambda x : "{:,}".format(x))
    full_df = full_df.sort_values(by=full_df.columns[1], ascending=False)
    global last_prod_frame
    last_prod_frame = full_df
    return full_df