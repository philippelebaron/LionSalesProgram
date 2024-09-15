import model

def set_file_name(file) -> str:
    if file == "default":
        return "Excel File Name"
    return file

def get_filename():
    return model.filename

def get_sheets(file):
    return model.get_sheets(file)

def set_model_df(file, sheet):
    return model.set_df(file, sheet)
    
def get_8020_p(file, sheet):
    return model.calculate_8020_production_df(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def get_8020_p_b(file, sheet):
    return model.get_b_production(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def get_all_p(file, sheet):
    return model.get_all_production(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def get_8020p_sales(file, sheet):
    return model.calculate_8020_production_stats(set_model_df(file, sheet), sheet).applymap(format_int_with_commas).to_string(index=False, col_space=25)

def get_8020_t(file, sheet):
    return model.calculate_8020_sales(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def get_8020_t_b(file, sheet):
    return model.get_b_transaction(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def get_all_t(file, sheet):
    return model.get_all_transaction(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def get_8020_tsales(file, sheet):
    return model.calculate_sales_stats(set_model_df(file, sheet), sheet).to_string(index=False, col_space=25)

def get_abcd_stats(file, sheet):
    return model.get_abcd_all(set_model_df(file, sheet)).to_string(index=False)

def summarize_prod(file, sheets):
    return model.summarize_data_all(file, sheets).to_string(index=False)

def summarize_tran(file, sheets):
    return model.summarize_data_tran(file, sheets).to_string(index=False)

def show():
    model.show_plot()

def get_bottom_20_stats(file, sheet):
    return model.bottom_20_stats(set_model_df(file, sheet), sheet).to_string(index=False, col_space=25)

def get_bottom_20(file, sheet):
    return model.bottom_20(set_model_df(file, sheet)).to_string(justify="center", col_space=[50, 35])

def set_store_1():
    model.set_stor1()

def set_store_2():
    model.set_stor2()

def format_int_with_commas(x):
    if isinstance(x, int):
        return f"{x:,}"
    else:
        return x
    
def get_current_target():
    return model.get_current_target()

def get_avg_proportion():
    return model.get_average_a_proportion()

def validate_simu():
    return model.validate_simu_input()

def simulate(targ, prop):
    simulation = model.simulate(int(targ), float(prop))


    simulation.iloc[7, 1] = "clrGr*" + str(simulation.iloc[7, 1]) + "clrGr**"
    simulation.iloc[7, 2] = "clrGr*" + str(simulation.iloc[7, 2]) + "clrGr**"
    simulation.iloc[7, 3] = "clrGr*" + str(simulation.iloc[7, 3]) + "clrGr**"
    simulation.iloc[7, 4] = "clrGr*" + str(simulation.iloc[7, 4]) + "clrGr**"
    simulation.iloc[3, 1] = "clrGr*" + str(simulation.iloc[3, 1]) + "clrGr**"
    simulation.iloc[3, 2] = "clrGr*" + str(simulation.iloc[3, 2]) + "clrGr**"
    simulation.iloc[3, 3] = "clrGr*" + str(simulation.iloc[3, 3]) + "clrGr**"
    simulation.iloc[3, 4] = "clrGr*" + str(simulation.iloc[3, 4]) + "clrGr**"


    simulation.iloc[8, 1] = "bread*" + str(simulation.iloc[8, 1]) + "bread**"
    simulation.iloc[9, 1] = "bread*" + str(simulation.iloc[9, 1]) + "bread**"
    simulation.iloc[10, 1] = "bread*" + str(simulation.iloc[10, 1]) + "bread**"

    simulation.iloc[8, 2] = "bread*" + str(simulation.iloc[8, 2]) + "bread**"
    simulation.iloc[9, 2] = "bread*" + str(simulation.iloc[9, 2]) + "bread**"
    simulation.iloc[10, 2] = "bread*" + str(simulation.iloc[10, 2]) + "bread**"

    simulation.iloc[0, 3] = "bread*" + str(simulation.iloc[0, 3]) + "bread**"
    simulation.iloc[1, 3] = "bread*" + str(simulation.iloc[1, 3]) + "bread**"
    simulation.iloc[2, 3] = "bread*" + str(simulation.iloc[2, 3]) + "bread**"

    simulation.iloc[0, 4] = "bread*" + str(simulation.iloc[0, 4]) + "bread**"
    simulation.iloc[1, 4] = "bread*" + str(simulation.iloc[1, 4]) + "bread**"
    simulation.iloc[2, 4] = "bread*" + str(simulation.iloc[2, 4]) + "bread**"

    
    return simulation.to_string(index=False)

def sort():
    return model.sort_repeats(model.get_repeats()).to_string(index=False)

def clear_store_1():
    return model.clear_summary_1()

def clear_store_2():
    return model.clear_summary_2()