import tkinter as tk
from tkinter import *

root = tk.Tk()
root.title("Query UI Testing")
root.geometry("850x1000")

combined_selections = set()

# Define the function to filter the tables listbox
def apply_filters(listbox, text_var, items_list):
   # selected_items = {listbox.get(i) for i in listbox.curselection()}
    search_text = text_var.get().strip()
    filtered = items_list.copy()
    filtered = [item for item in filtered if search_text in item]
    listbox.delete(0, tk.END)
    for item in filtered:
        listbox.insert(tk.END, item)
    for i in range(listbox.size()):
        if listbox.get(i) in combined_selections:
            listbox.selection_set(i)

# define the function to run when the search text changes
def on_change_search_text(listbox, text_var, items_list):
    apply_filters(listbox, text_var, items_list)

# define a function when the add button is clicked
def on_click_add_selection():
    selected_items = [listbox.get(i) for i in listbox.curselection()] 
    combined_selections.update(selected_items)
    selections_label.config(text=combined_selections)


# define the search text string variable
text_var = tk.StringVar()

# define the list of items for the listbox
items_list = ['Element1', 'Element2', 'Element3', 'Element4', 'Item1', 'Item2', 'Item3', 'Item4', 'Choice1', 'Choice2', 'Choice3', 'Choice4']

# add the search box widget
search_box = tk.Entry(
    root, textvariable=text_var
)
search_box.grid(
    row=0, column=0, padx=10, pady=10, sticky='ew'
)

# add the listbox and populate it
listbox = tk.Listbox(
    root, selectmode="multiple"
)
listbox.grid(
    row=1, column=0, padx=10, pady=10, sticky='ew'
)
for item in items_list:
    listbox.insert(
        tk.END, item
    )

# add the button to add selections
add_selection = tk.Button(root, text="Add to selection", command=lambda: on_click_add_selection())
add_selection.grid(row=3, column=0, padx=10, pady=10, sticky='ew')

# add the label to show the selections made
selections_label = tk.Label(root, text="Selections Made:")
selections_label.grid(row =4, column=0, padx=10, pady=10, sticky='ew')

# assign a function to the text variable
text_var.trace_add(
    "write",
    lambda *args: on_change_search_text(
        listbox, text_var, items_list)
    )










if __name__ == "__main__":
    root.mainloop()