import model_lite_mac

def set_file_name(file) -> str:
    if file == "default":
        return "Excel File Name"
    return file

def get_filename():
    return model_lite_mac.filename

def get_sheets(file):
    return model_lite_mac.get_sheets(file)

def set_model_df(file, sheet):
    return model_lite_mac.set_df(file, sheet)
    
def get_8020p_sales(file, sheet):
    return model_lite_mac.calculate_8020_production_stats(set_model_df(file, sheet), sheet).applymap(format_int_with_commas).to_string(index=False, col_space=25)

def format_int_with_commas(x):
    if isinstance(x, int):
        return f"{x:,}"
    else:
        return x