# Import the required libraries
import tkinter as tk


def select_distinct_output(mode):
    """
    This function selects the distinct output from the query
    """
    if mode.get() == 1:
        querying_master_list = True
    else:
        querying_master_list = False
    return querying_master_list


def apply_table_filters(
    tables_listbox,
    tables_list,
    text_var,
    check_var_supp,
    check_var_anon_followup,
    check_var_unanon_followup,
    tables_listbox_result,
    combined_selections
):
    """
    This function runs when the tick boxes for dataset types are ticked/unticked
    """
    # Import the required functions
    from code_files.utilities.ui.on_selections import on_tables_listbox_select

    search_text = text_var.get().strip()
    filtered = tables_list.copy()

    # Apply category filter if any checkbox is active
    if check_var_supp.get() == 1:
        filtered = filtered[
            filtered["table_name"].str.contains(
                "Supplementary_Anon", case=False, na=False
            )
        ]
    elif check_var_anon_followup.get() == 1:
        pattern = "FollowUp_Anon|Prelim_Anon"
        filtered = filtered[
            filtered["table_name"].str.contains(pattern, case=False, na=False)
        ]
    elif check_var_unanon_followup.get() == 1:
        pattern = "FollowUp_Unanon|Prelim_Unanon"
        filtered = filtered[
            filtered["table_name"].str.contains(pattern, case=False, na=False)
        ]

    # Apply text search filter
    if search_text:
        filtered = filtered[
            filtered["table_name"].str.contains(search_text, case=False, na=False)
        ]

    # Re-populate the Listbox
    tables_listbox.delete(0, tk.END)
    sorted_filtered = filtered.sort_values(by=["table_schema", "table_name"])
    for idx, row in sorted_filtered.reset_index(drop=True).iterrows():
        item_str = f"{row['table_schema']}.{row['table_name']}"
        tables_listbox.insert(tk.END, item_str)

    # Restore previous selections
    for i  in range(tables_listbox.size()):
        if tables_listbox.get(i) in combined_selections:
            tables_listbox.selection_set(i)


    # Trigger selection label update manually to match restored items
    on_tables_listbox_select(None, tables_listbox, tables_listbox_result)


# By default, supplementary data is not selected until the checkbox is ticked
is_Supplementary = False
