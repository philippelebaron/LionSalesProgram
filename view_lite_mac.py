import tkinter as tk
from tkinter import filedialog
from tkmacosx import Button
import controller_lite_mac

# Create new tk window with minimum size 600x400
window = tk.Tk()
window.minsize(750, 540)
window.title("Sales Data Lite")


# Create text to display name of excel file selected
filetext = tk.Label(text=controller_lite_mac.set_file_name("default"))
filetext.grid(row=0, column=1, columnspan=4, sticky='w', padx=10, pady=10)


# Handles logic behind browse button 
# (sets filetext to excel file location and passes it to controller_lite_mac)
# Also sets listbox sheets upon selection
def browse_button() -> None:
    filetext["text"] = controller_lite_mac.set_file_name(filedialog.askopenfilename())
    listbox.delete(0, tk.END)
    new_choices = controller_lite_mac.get_sheets(filetext["text"])
    for choice in new_choices:
        listbox.insert(tk.END, choice)
    

# Creates listbox to keep track of selections of sheets
sheet_names = []
listbox = tk.Listbox(window, selectmode="multiple", exportselection=False)
listbox.grid(row=1, column=1, sticky='ew', columnspan=3, pady=10, rowspan=2)

# Creates scrollbar for listbox
scroll_list = tk.Scrollbar(window, orient="vertical")
scroll_list.grid(row=1, column=3, sticky='nse', pady=10, rowspan=2)
listbox.config(yscrollcommand=scroll_list.set)
scroll_list.config(command=listbox.yview)


# Button to browse/select data file
button = Button(window, text="Browse", command=browse_button)
button.grid(row=0, column=4)

# Vertical scroll for display
scroll_vert = tk.Scrollbar(window, orient="vertical")
scroll_vert.grid(row=4, column=5, sticky='ns')

# Horizontal scroll for display
scroll_hor = tk.Scrollbar(window, orient = 'horizontal')
scroll_hor.grid(row=5, column=0, columnspan=5, sticky='ew')

# Display area for showing data
display_label = tk.Text(window, width=115, font=("Courier New", 10), yscrollcommand=scroll_vert.set, wrap=tk.NONE, height=20)
display_label.grid(row=4, column=0, columnspan=5, sticky='ew')

# Configures x and y scrollbars
display_label.config(xscrollcommand=scroll_hor.set)
display_label.config(yscrollcommand=scroll_vert.set)
scroll_vert.config(command=display_label.yview)
scroll_hor.config(command=display_label.xview)

production_frame = tk.Frame(window)
production_frame.grid(row=3, column=2, columnspan=1, pady=10)

# Retrieves and displays A players based on production
def get_a_players_p():
    # communicate to controller_lite_mac
    display_label.delete(1.0, tk.END)
    for sheet in sheet_names:
        display_label.insert(tk.END, sheet)
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller_lite_mac.get_8020p_sales(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
    return

# Button for 80/20 rule production
button_8020p = Button(production_frame, text="8020 Rule Production", command=get_a_players_p)
button_8020p.grid(row=0, column=0, pady=2, padx=2)

# Sends selected sheets to sheet_names array to be stored
def select():
    sheet_names.clear()
    for i in listbox.curselection():
        sheet_names.append(listbox.get(i))

def select_all():
    sheet_names.clear()
    listbox.select_set(0, tk.END)
    for i in listbox.curselection():
        sheet_names.append(listbox.get(i))

def reset_selection():
    sheet_names.clear()
    listbox.selection_clear(0, tk.END)


selection_frame = tk.Frame(window)
selection_frame.grid(row=1, column=4, rowspan=2)
# Button for selecting listed selections for sheets
button_pick_multi = Button(selection_frame, text="Select File(s)", command=select)
button_pick_multi.grid(row=0, column=0, pady=2)

button_select_all = Button(selection_frame, text = "Select All", command=select_all)
button_select_all.grid(row=1, column=0, pady=2)

button_reset_selections = Button(selection_frame, text = "Reset Selection", command=reset_selection)
button_reset_selections.grid(row=2, column=0, pady=2)

window.mainloop()