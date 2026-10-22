# Import the required libraries
import tkinter as tk
from tkinter import ttk


def on_tables_listbox_select(event, tables_listbox, tables_listbox_result):
    """
    This function handles the tables selection from the listbox
    """
    global selected_indices, selected_tables
    counter = 1
    try:
        current_selection = tables_listbox.curselection()
        selected_indices = list(current_selection)
        selected_tables = [tables_listbox.get(i) for i in selected_indices]
        label_result_txt = ""
        for tbl in selected_tables:
            label_result_txt += f"{counter}- {tbl}\n"
            counter += 1
        tables_listbox_result.config(
            text=f"Selected Tables:\n{label_result_txt}",
            font=("Consolas", 10),
            justify="left",
            anchor="w",
        )
        return selected_tables
    except IndexError:
        pass


def on_columns_listbox_select(event, columns_listbox, columns_listbox_result):
    """
    This function handles the columns selection from the listbox
    """
    global selected_col_indices, selected_cols
    counter = 1
    try:
        current_selection_columns = columns_listbox.curselection()
        selected_col_indices = list(current_selection_columns)
        selected_cols = [columns_listbox.get(i) for i in selected_col_indices]
        label_result_txt = ""
        for col in selected_cols:
            label_result_txt += f"{counter}- {col}\n"
            counter += 1
        columns_listbox_result.config(
            text=f"Selected Columns:\n{label_result_txt}",
            font=("Consolas", 10),
            justify="left",
            anchor="w",
        )
        return selected_cols
    except IndexError:
        pass


def on_select_reset_widgets(root):
    """
    This function resets all the widgets on the interface
    """

    def _reset_container(container):
        """
        This function resets the widgets inside a specific container
        """

        # loop through the widgets
        for widget in container.winfo_children():
            # if the widget has childrens inside it, reset them
            if widget.winfo_children():
                _reset_container(widget)

            # reset entry, listbox widgets
            elif isinstance(widget, (tk.Entry, ttk.Entry)):
                widget.delete(0, tk.END)

            # reset text widgets
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)

            # reset the checkbuttons
            elif isinstance(widget, (tk.Checkbutton, ttk.Checkbutton, tk.Radiobutton, ttk.Radiobutton)):
                widget.deselect()

            # reset the listbox widgets
            elif isinstance(widget, (tk.Listbox)):
                widget.selection_clear(0, tk.END)

    # reset root-level controls
    _reset_container(root)


def on_select_reset_tables(tables_listbox, tables_listbox_result, columns_listbox):
    # Clear the tables listbox
    tables_listbox.selection_clear(0, tk.END)
    # Clear the tables selected textbox
    tables_listbox_result.config(text="")
    # Clear the columns listbox
    columns_listbox.selection_clear(0, tk.END)


def on_select_deselect_all_columns(mode, columns_listbox, columns_listbox_result):
    """
    This function runs when the tickbox for selecting and deselecting all the columns is ticked/unticked
    """
    if mode.get() == 1:
        columns_listbox.selection_set(0, tk.END)
        columns_listbox_result.config(text="All columns selected")
    else:
        columns_listbox.selection_clear(0, tk.END)


def on_select_deselect_prelim_columns(
    mode, columns_listbox, check_var_select_prelim_cols
):
    """
    This function runs when the tickbox for selecting and deselecting the prelim columns is ticked/unticked
    """
    # get the listbox items
    items = columns_listbox.get(0, "end")

    # loop through the selected tables to determine if its supplementary data being queried
    for table in selected_tables:
        if "Supplementary" in table:
            check_var_select_prelim_cols.set(0)
            return False

    # get the index of the AnyNeoadjuvantTherapy column
    idx = items.index("AnyNeoadjuvantTherapy")
    if mode.get() == 1:
        columns_listbox.select_set(0, idx)
    else:
        columns_listbox.selection_clear(0, tk.END)


def on_select_deselect_all_tables(mode, tables_listbox):
    """
    This function runs when the tickbox for selecting and deselecting all the table names is ticked/unticked
    """
    if mode.get() == 1:
        tables_listbox.selection_set(0, tk.END)
    else:
        tables_listbox.selection_clear(0, tk.END)


def on_select_searchbox_confirmation(tables_listbox, columns_listbox):
    """
    This function runs the apply_table_filters function when the search box confirmation button is clicked
    """
    # Import the required functions
    from code_files.utilities.ui.callbacks import callback_table_selection

    callback_table_selection(tables_listbox, columns_listbox)


def on_click_add_to_selection(tables_listbox, tables_listbox_result, combined_selection):
    selected_items = [tables_listbox.get(i) for i in tables_listbox.curselection()]
    combined_selection.update(selected_items)
    tables_listbox_result.config(text=combined_selection)
    

