# Import the required functions
from code_files.utilities.ui.configure_outputs import apply_table_filters
from code_files.utilities.ui.on_selections import on_select_deselect_prelim_columns
from code_files.utilities.ui.update_widgets import update_prelim_checkbox_state


def on_change_supp_checkbox(
    check_var_supp,
    check_var_unanon_supp,
    check_var_anon_followup,
    check_var_unanon_followup,
    tables_listbox,
    tables_list,
    text_var,
    tables_listbox_result,
    combined_selections
):
    """
    This function sets the value for the check_var_supp when the tick box for supplementary tables is used
    """
    if check_var_supp.get() == 1:
        check_var_anon_followup.set(0)
        check_var_unanon_followup.set(0)
    apply_table_filters(
        tables_listbox,
        tables_list,
        text_var,
        check_var_supp,
        check_var_unanon_supp,
        check_var_anon_followup,
        check_var_unanon_followup,
        tables_listbox_result,
        combined_selections
    )


def on_change_anon_followup_checkbox(
    check_var_anon_followup,
    check_var_supp,
    check_var_unanon_supp,
    check_var_unanon_followup,
    tables_listbox,
    tables_list,
    text_var,
    tables_listbox_result,
    combined_selections
):
    """
    This function sets the value for the check_var_supp when the tick box for anonymised follow up tables is used
    """
    if check_var_anon_followup.get() == 1:
        check_var_supp.set(0)
        check_var_unanon_followup.set(0)
    apply_table_filters(
        tables_listbox,
        tables_list,
        text_var,
        check_var_supp,
        check_var_unanon_supp,
        check_var_anon_followup,
        check_var_unanon_followup,
        tables_listbox_result,
        combined_selections
    )


def on_change_unanon_followup_checkbox(
    check_var_unanon_followup,
    check_var_supp,
    check_var_unanon_supp,
    check_var_anon_followup,
    tables_listbox,
    tables_list,
    text_var,
    tables_listbox_result,
    combined_selections
):
    """
    This function sets the value for the check_var_supp when the tick box for unanonymised follow up tables is used
    """
    if check_var_unanon_followup.get() == 1:
        check_var_supp.set(0)
        check_var_anon_followup.set(0)
    apply_table_filters(
        tables_listbox,
        tables_list,
        text_var,
        check_var_supp,
        check_var_unanon_supp,
        check_var_anon_followup,
        check_var_unanon_followup,
        tables_listbox_result,
        combined_selections
    )


def on_change_prelim_columns_checkbox(
    mode,
    select_prelim_cols_checkbox,
    columns_listbox,
    check_var_select_prelim_cols,
    tables_listbox
):
    # Handle when the checkbox is toggled by the user
    update_prelim_checkbox_state(
        tables_listbox, select_prelim_cols_checkbox, check_var_select_prelim_cols
    )
    # only execute the logic if the checkbox is active
    if select_prelim_cols_checkbox["state"] == "normal":
        on_select_deselect_prelim_columns(
            mode, columns_listbox, check_var_select_prelim_cols
        )


def on_change_search_text(
    tables_listbox,
    tables_list,
    text_var,
    check_var_supp,
    check_var_unanon_supp,
    check_var_anon_followup,
    check_var_unanon_followup,
    tables_listbox_result,
    combined_selections
):
    """
    This function re-runs the apply_table_filters() function everytime there is a change in the text typed into
    the search box
    """
    apply_table_filters(
        tables_listbox,
        tables_list,
        text_var,
        check_var_supp,
        check_var_unanon_supp,
        check_var_anon_followup,
        check_var_unanon_followup,
        tables_listbox_result,
        combined_selections
    )
