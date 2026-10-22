def update_prelim_checkbox_state(
    tables_listbox, select_prelim_cols_checkbox, check_var_select_prelim_cols
):
    # Enable or disable the checkbox based on the listbox selection
    # get the selected tables
    selected_tables = {tables_listbox.get(i) for i in tables_listbox.curselection()}
    # determine whether the selected tables have supplementary data
    has_supplementary = any("Supplementary" in table for table in selected_tables)
    # if Supplementary data is selected, disable the select prelim columns checkbox
    if has_supplementary:
        select_prelim_cols_checkbox.config(state="disabled")
        check_var_select_prelim_cols.set(0)
    else:
        select_prelim_cols_checkbox.config(state="normal")
