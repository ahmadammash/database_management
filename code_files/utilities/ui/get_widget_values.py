# Import the required functions
from code_files.utilities.ui.configure_outputs import select_distinct_output


def get_checkbox_returnFirstBlock(
    check_var_returnFirstBlock_query, show_is_supp_query_checkbox
):
    """
    This function runs when the checkbox to return the first block only is ticked/unticked
    """
    if check_var_returnFirstBlock_query.get() == 1:
        # disable the checkbox for the supplementary data
        show_is_supp_query_checkbox.config(state="disabled")
        return_firstBlockOnly = True
    else:
        # ensure the checkbox for the supplementary data is enabled again
        show_is_supp_query_checkbox.config(state="active")
        return_firstBlockOnly = False
    return return_firstBlockOnly


def get_checkbox_returnFollowUp_Only(
    check_var_returnFollowUpOnly_query,
    show_is_var_returnFirstAvailableRecord,
    show_is_supp_query_checkbox,
):
    """
    This function runs when the checkbox for returning master list data from follow up tables only
    """
    if check_var_returnFollowUpOnly_query.get() == 1:
        # disable the checkbox to return the first available record as both cannot be logically selected together
        show_is_var_returnFirstAvailableRecord.config(state="disabled")
        # disable the checkbox for the supplementary data
        show_is_supp_query_checkbox.config(state="disabled")
        return_FollowUpOnly = True
    else:
        # ensure the checkbox to return the first available record is enabled again
        show_is_var_returnFirstAvailableRecord.config(state="active")
        # ensure the checkbox for the supplementary data is enabled again
        show_is_supp_query_checkbox.config(state="active")
        return_FollowUpOnly = False
    return return_FollowUpOnly


def get_checkbox_returnFirstAvailableRecord(
    check_var_returnFirstAvailableRecord,
    show_is_var_returnFollowUpOnly_query,
    show_is_supp_query_checkbox,
):
    """
    This function runs when the checkbox for returning the first available record is ticked
    """
    if check_var_returnFirstAvailableRecord.get() == 1:
        # disable the checkbox to return the follow up data only as both cannot be logically selected together
        show_is_var_returnFollowUpOnly_query.config(state="disabled")
        return_FirstAvailableRecord = True
        # disable the checkbox for the supplementary data
        show_is_supp_query_checkbox.config(state="disabled")
    else:
        # ensure the checkbox to return the follow up data only is enabled again
        show_is_var_returnFollowUpOnly_query.config(state="active")
        # ensure the checkbox for the supplementary data is enabled again
        show_is_supp_query_checkbox.config(state="active")
        return_FirstAvailableRecord = False
    return return_FirstAvailableRecord


def get_checkbox_isSuppQuery(
    check_var_is_supp_query,
    show_is_var_returnFirstAvailableRecord,
    show_is_var_returnFollowUpOnly_query,
    show_is_var_master_query,
    show_is_var_returnFirstBlock_query,
):
    """
    This function runs when the checkbox for querying supplementary data is ticked
    """
    if check_var_is_supp_query.get() == 1:
        # disable the checkboxes for the master list
        show_is_var_returnFirstAvailableRecord.config(state="disabled")
        show_is_var_returnFollowUpOnly_query.config(state="disabled")
        show_is_var_master_query.config(state="disabled")
        show_is_var_returnFirstBlock_query.config(state="disabled")
    else:
        # ensure the checkboxes for the master list are enabled again
        show_is_var_returnFirstAvailableRecord.config(state="active")
        show_is_var_returnFollowUpOnly_query.config(state="active")
        show_is_var_master_query.config(state="active")
        show_is_var_returnFirstBlock_query.config(state="active")


def get_checkbox_isMasterList(
    check_var_master_query, show_is_supp_query_checkbox, show_is_var_master_query
):
    """
    This function runs when the checkbox for querying the master list is ticked
    """
    if check_var_master_query.get() == 1:
        # disable the checkbox for the supplementary data
        show_is_supp_query_checkbox.config(state="disabled")
        show_is_var_master_query.config(
            command=select_distinct_output(check_var_master_query)
        )
        return_masterList = True
    else:
        # ensure the checkbox for the supplementary data is enabled again
        show_is_supp_query_checkbox.config(state="active")
        return_masterList = False
    return return_masterList


def get_checkbox_combine_followup_supp_data(
    check_var_combine_followup_supp_data, selected_tables
):
    """
    This function runs when the checkbox for combing the follow up and supplementary data is ticked
    """
    if check_var_combine_followup_supp_data.get() == 1:
        # divide the selected tables into prelim/follow up and supplementary
        prelim_followup_tables = [
            tbl
            for tbl in selected_tables
            if any(x in tbl for x in ["FollowUp_Anon", "Prelim_Anon"])
        ]
        [
            tbl
            for tbl in selected_tables
            if any(x in tbl for x in ["Supplementary_Anon"])
        ]
        print(prelim_followup_tables)
