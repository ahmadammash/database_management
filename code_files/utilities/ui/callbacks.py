# Import the required libraries
import tkinter as tk

import pandas as pd

# Import the required functions
from code_files.querying.fetch_information import fetch_col_names, query_data


def callback_table_selection(tables_listbox, columns_listbox):
    """
    This function runs when the "Confirm Tables Selection" button is clicked
    """
    from code_files.utilities.ui.column_configs import (
        col_order_Anon,
        col_order_master_list_long,
        col_order_PatientList,
        col_order_Unanon,
    )

    selected_indices = tables_listbox.curselection()
    selected_tables_list = [tables_listbox.get(i) for i in selected_indices]
    columns_listbox.delete(0, tk.END)
    if selected_tables_list:
        try:
            cols_list = fetch_col_names(selected_tables_list)

            # determine if prelim/follow up data is selected
            # get the first item from the selected tables
            first_table = tables_listbox.get(selected_indices[0])
            # if prelim/follow up data/master list is selected, sort the columns per template
            if "FollowUp_Anon" in first_table or "Prelim_Anon" in first_table:
                col_order = col_order_Anon
            # if unanonymised data is selected, sort the columns per template
            elif "Prelim_Unanon" in first_table or "FollowUp_Unanon" in first_table:
                col_order = col_order_Unanon
            # if the patient list is selected, sort the columns per template
            elif "PatientList" in first_table:
                col_order = col_order_PatientList
            # if the master list is selected, sort the columns per template
            elif "master_list_long" in first_table:
                col_order = col_order_master_list_long
            # if supplementary data is selected, sort the columns alphabetically
            elif "Supplementary" in first_table:
                cols_list = sorted(cols_list, key=str.lower)

            # for all except supplementary tables, sort the table by col_order
            if "Supplementary" not in first_table and "col_order" in locals():
                ordered = [col for col in col_order if col in cols_list]
                remaining = [col for col in cols_list if col not in ordered]
                cols_list = ordered + remaining

            for col in cols_list:
                columns_listbox.insert(tk.END, col)
        except (TypeError, ValueError, KeyError, OSError) as e:
            print(f"Error fetching columns: {e}")


def callback_run_query(
    tables_listbox,
    columns_listbox,
    id_selected,
    id_list_in_clause_entry,
    id_exclusions,
    id_list_exclusions_in_clause_entry,
    order_by_clause_entry,
    where_clause_entry,
    check_var_is_supp_query,
    top_entry,
    check_var_master_query,
    show_is_var_master_query,
    check_var_returnFollowUpOnly_query,
    show_is_var_returnFollowUpOnly_query,
    check_var_returnFirstAvailableRecord,
    show_is_var_returnFirstAvailableRecord,
    check_var_returnFirstBlock_query,
    show_is_supp_query_checkbox,
):
    """
    This function runs when the "Run Query" button is clicked
    """
    # Import the column order lists
    from code_files.utilities.ui.get_widget_values import (
        get_checkbox_isMasterList,
        get_checkbox_returnFirstAvailableRecord,
        get_checkbox_returnFirstBlock,
        get_checkbox_returnFollowUp_Only,
    )

    global df_result
    global id_column_output, id_list_output
    df_result = pd.DataFrame()

    # Get selected tables as a clean list
    current_selection_tables = tables_listbox.curselection()
    selected_tables_list = [tables_listbox.get(i) for i in current_selection_tables]

    # Get selected columns as a clean list
    selected_indices_columns = columns_listbox.curselection()
    selected_cols_list = [columns_listbox.get(i) for i in selected_indices_columns]

    # Join tables with a pipe '|' or pass as a list so spaces don't break them
    # Join columns with a comma ',' as standard SQL formatting requires
    tables_payload = "|".join(selected_tables_list)
    selected_cols_str = ",".join(selected_cols_list)

    id_column_output = id_selected.get()
    id_column_exclusion_output = id_exclusions.get()
    id_list_output = id_list_in_clause_entry.get("1.0", "end-1c")
    id_list_exclusions_output = id_list_exclusions_in_clause_entry.get("1.0", "end-1c")
    order_by_clause = order_by_clause_entry.get("1.0", "end-1c")
    where_clause = where_clause_entry.get("1.0", "end-1c")
    where_clause_conditions = [
        line.strip() for line in where_clause.splitlines() if line.strip()
    ]
    is_Supplementary = check_var_is_supp_query.get() == 1

    df_result, query_text = query_data(
        tables=tables_payload,
        columns=selected_cols_str,
        conditions=where_clause_conditions,
        id_in_clause=id_column_output,
        in_clause=id_list_output,
        id_exclusions_clause=id_column_exclusion_output,
        exclusions_clause=id_list_exclusions_output,
        order_by=order_by_clause,
        duplicate_results=True,
        isSupplementary=is_Supplementary,
        isMasterList=get_checkbox_isMasterList(
            check_var_master_query,
            show_is_supp_query_checkbox,
            show_is_var_master_query,
        ),
        applyReplacements=True,
        top_rows=top_entry.get(),
        return_firstBlockOnly=get_checkbox_returnFirstBlock(
            check_var_returnFirstBlock_query, show_is_supp_query_checkbox
        ),
        return_FollowUpOnly=get_checkbox_returnFollowUp_Only(
            check_var_returnFollowUpOnly_query,
            show_is_var_returnFirstAvailableRecord,
            show_is_supp_query_checkbox,
        ),
        return_firstAvailableRecord=get_checkbox_returnFirstAvailableRecord(
            check_var_returnFirstAvailableRecord,
            show_is_var_returnFollowUpOnly_query,
            show_is_supp_query_checkbox,
        ),
    )
    print("\n--- Final Merged Output DataFrame ---")
    print(
        "=========================================================================================================="
    )
    print(
        "=========================================================================================================="
    )
    print(df_result.reset_index(drop=True))
    df_result.to_clipboard()
    print(query_text)
    print(
        "=========================================================================================================="
    )
    print(
        "=========================================================================================================="
    )
