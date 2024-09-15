import tkinter as tk
from tkinter import filedialog
from tkmacosx import Button
import controller_mac

# Create new tk window with minimum size 600x400
window = tk.Tk()
window.minsize(960, 640)
window.title("Sales Data")


# Create text to display name of excel file selected
filetext = tk.Label(text=controller.set_file_name("default"))
filetext.grid(row=0, column=1, columnspan=4, sticky='w', padx=10, pady=10)


# Handles logic behind browse button 
# (sets filetext to excel file location and passes it to controller)
# Also sets listbox sheets upon selection
def browse_button() -> None:
    filetext["text"] = controller.set_file_name(filedialog.askopenfilename())
    listbox.delete(0, tk.END)
    new_choices = controller.get_sheets(filetext["text"])
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
production_frame.grid(row=3, column=0, columnspan=1, pady=10)

# Retrieves and displays A players based on production
def get_a_players_p():
    # communicate to controller
    display_label.delete(1.0, tk.END)
    for sheet in sheet_names:
        display_label.insert(tk.END, sheet)
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_8020p_sales(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_8020_p(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
    return

def get_bottom_twenty():
    display_label.delete(1.0, tk.END)
    for sheet in sheet_names:
        display_label.insert(tk.END, sheet)
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_bottom_20_stats(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_bottom_20(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
    return

def get_sorted():
    display_label.delete(1.0, tk.END)
    display_label.insert(tk.END, '\n')
    display_label.insert(tk.END, controller.sort())
    display_label.insert(tk.END, '\n')
    display_label.insert(tk.END, '\n')
    return

# Button for 80/20 rule production
button_8020p = Button(production_frame, text="8020 Rule Production", command=get_a_players_p)
button_8020p.grid(row=0, column=0, pady=2, padx=2)

button_bottom_twenty = Button(production_frame, text="Bottom 20 Production", command=get_bottom_twenty)
button_bottom_twenty.grid(row=1, column=0, pady=2, padx=2)

button_sort = Button(window, text="Sort", command=get_sorted)
button_sort.grid(row=3, column=1, pady=2, padx=2)

# Displays/retrieves data for 80/20 rule based on transactions
def get_a_players_t():
    display_label.delete(1.0, tk.END)
    for sheet in sheet_names:
        display_label.insert(tk.END, sheet)
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_8020_tsales(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_8020_t(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
    return

# Button for 80/20 transactions
button_8020t = Button(production_frame, text="8020 Rule Transactions", command=get_a_players_t)
button_8020t.grid(row=2, column=0, pady=2)

# Retrives/displays abcd players and stats for given sheet
def get_abcd_players():
    # communicate to controller
    display_label.delete(1.0, tk.END)
    for sheet in sheet_names:
        display_label.insert(tk.END, sheet)
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, controller.get_abcd_stats(filetext["text"], sheet))
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
        display_label.insert(tk.END, '\n')
    return

# button for abcd players
button_abcd = Button(window, text="ABCD Players", command=get_abcd_players)
button_abcd.grid(row=3, column=2, pady=10)

def get_graph():
    # communicate to controller
    display_label.delete(1.0, tk.END)
    display_label.insert(tk.END, controller.summarize_prod(filetext["text"], sheet_names))
    display_label.insert(tk.END, '\n')
    display_label.insert(tk.END, '\n')
    return

sum_frame = tk.Frame(window)
sum_frame.grid(row=3, column=4, columnspan=1, rowspan=1, pady=10)

button_sum = Button(sum_frame, text="Summarize", command=get_graph)
button_sum.grid(row=0, column=0, pady=2, padx=2)

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

button_show_graphs = Button(sum_frame, text = "Graph", command = controller.show)
button_show_graphs.grid(row=1, column=0, pady=2, padx=2, sticky='e')

def store_1():
    controller.set_store_1()
    button_store_1.configure(background='#969696')


button_store_1 = Button(sum_frame, text="Store 1", command = store_1)
button_store_1.grid(row=0, column=1, pady=2, padx=2)

def store_2():
    controller.set_store_2()
    button_store_2.configure(background='#969696')


button_store_2 = Button(sum_frame, text="Store 2", command = store_2)
button_store_2.grid(row=1, column=1, pady=2, padx=2)

def reset_store_1():
    controller.clear_store_1()
    button_store_1.configure(background='#ffffff')


button_reset_store_1 = Button(sum_frame, text="Reset Store 1", command = reset_store_1)
button_reset_store_1.grid(row=0, column=2, pady=2, padx=2)

def reset_store_2():
    controller.clear_store_2()
    button_store_2.configure(background='#ffffff')

button_reset_store_2 = Button(sum_frame, text="Reset Store 2", command = reset_store_2)
button_reset_store_2.grid(row=1, column=2, pady=2, padx=2)


def simulate(target, a_prop):
    display_label.delete(1.0, tk.END)
    display_label.insert(tk.END, controller.simulate(target, a_prop))
    display_label.insert(tk.END, '\n')
    display_label.insert(tk.END, '\n')
    display_label.insert(tk.END, "clrGr*- inputclrGr**")
    display_label.insert(tk.END, '\n')
    display_label.insert(tk.END, "bread*- outputbread**")
    

    still_green_indicators = True
    still_red_indicators = True
    
    while(still_green_indicators or still_red_indicators):
        display_label.tag_configure("highlight", foreground="green")
        display_label.tag_configure("lowlight", foreground="red")
        display_label.tag_configure("hidden", elide=True)
        char_count = tk.IntVar()
        index = display_label.search(r'(?:clrGr\*).*(?:clrGr\*\*)', "1.0", "end", count=char_count, regexp=True)

        # we have to adjust the character indexes to skip over the identifiers
        if index != "":
            start = "%s + 6 chars" % index
            end = "%s + %d chars" % (index, char_count.get()-7)
            display_label.tag_add("highlight", start, end)
        
        display_label.tag_configure("hidden", elide=True)

        index = display_label.search(r'(?:clrGr\*).*(?:clrGr\*\*)', "1.0", "end", count=char_count, regexp=True)
        if index != "":
            start = index
            end = "%s + 6 chars" % index
            display_label.tag_add("hidden", start, end)
            display_label.insert(index, "             ")
        else:
            still_green_indicators = False

        
        index = display_label.search(r'(?:clrGr\*\*)', "1.0", "end", count=char_count, regexp=True)
        if index != "":
            start = index
            end = "%s + 7 chars" % index
            display_label.tag_add("hidden", start, end)


        index = display_label.search(r'(?:bread\*).*(?:bread\*\*)', "1.0", "end", count=char_count, regexp=True)        
        if index != "":
            start = "%s + 6 chars" % index
            end = "%s + %d chars" % (index, char_count.get()-7)
            display_label.tag_add("lowlight", start, end)
        

        index = display_label.search(r'(?:bread\*).*(?:bread\*\*)', "1.0", "end", count=char_count, regexp=True)
        if index != "":
            start = index
            end = "%s + 6 chars" % index
            display_label.tag_add("hidden", start, end)
            display_label.insert(index, "             ")
        else:
            still_red_indicators = False

        
        index = display_label.search(r'(?:bread\*\*)', "1.0", "end", count=char_count, regexp=True)
        if index != "":
            start = index
            end = "%s + 7 chars" % index
            display_label.tag_add("hidden", start, end)
    return
    


def close_win(top, entry1, entry2):
    simulate(entry1, entry2)
    top.destroy()


#Define a function to open the Popup Dialogue
def popupwin():
    if controller.validate_simu():
        top = tk.Toplevel(window)
        top.geometry("250x100")
        top.title("Simulate")

        label_1 = tk.Label(top, text="Target:")
        label_1.grid(row = 0, column = 0, sticky='e')

        label_2 = tk.Label(top, text="A Proportion:")
        label_2.grid(row = 1, column = 0, sticky='e')

        entry_1 = tk.Entry(top, width = 25)
        entry_1.grid(row=0, column=1, sticky='w')

        entry_1.insert(0, str(controller.get_current_target()))

        entry_2 = tk.Entry(top, width = 25)
        entry_2.grid(row=1, column=1, sticky='w')

        entry_2.insert(0, str(controller.get_avg_proportion()))
    
        #Create a Button Widget in the Toplevel Window
        button= Button(top, text="Run", command=lambda:close_win(top, entry_1.get(), entry_2.get()))
        button.grid(row=2, column=0, pady=5, columnspan=2)





button_simulate = Button(window, text='Simulate', command=popupwin)
button_simulate.grid(row=3, column=3, pady=10)


window.mainloop()