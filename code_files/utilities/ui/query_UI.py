# ================================================================================================================
# IMPORTS AND SYSTEMS SETUP
# ================================================================================================================

# Import the required libraries
import tkinter as tk
import warnings
from tkinter import *
from tkinter import scrolledtext, ttk

# Define a set variable to store the combined selections
combined_selections = set()

# Suppress warnings
warnings.simplefilter("ignore", UserWarning)
warnings.simplefilter("ignore", FutureWarning)


# Import functions from the project directory
from code_files.querying.fetch_information import get_table_names
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.ui.callbacks import callback_run_query
from code_files.utilities.ui.get_widget_values import (
    get_checkbox_combine_followup_supp_data,
    get_checkbox_isMasterList,
    get_checkbox_isSuppQuery,
    get_checkbox_returnFirstAvailableRecord,
    get_checkbox_returnFirstBlock,
    get_checkbox_returnFollowUp_Only,
)
from code_files.utilities.ui.on_changes import (
    on_change_anon_followup_checkbox,
    on_change_prelim_columns_checkbox,
    on_change_search_text,
    on_change_supp_checkbox,
    on_change_unanon_followup_checkbox,
)
from code_files.utilities.ui.on_selections import (
    on_click_add_to_selection,
    on_columns_listbox_select,
    on_select_deselect_all_columns,
    on_select_deselect_all_tables,
    on_select_reset_tables,
    on_select_reset_widgets,
    on_select_searchbox_confirmation,
    on_tables_listbox_select,
)
from code_files.utilities.ui.update_widgets import update_prelim_checkbox_state

# ================================================================================================================
# DATABASE CONNECTION & INITIALISATION
# ================================================================================================================

project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# define the empty lists for the selected indices and selected tables
selected_indices = []
selected_tables = []


# ================================================================================================================
# CORE UI Building
# ================================================================================================================

# Build the GUI application
root = tk.Tk()
root.title("SQL Server Table Viewer")
root.geometry("850x1000")

def build_ui():

    # Create a main frame to hold the canvas and the scrollbar
    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=1)

    # Create a canvas inside the main frame
    canvas = tk.Canvas(main_frame)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

    # Add a scrollbar to the main frame and link it to the canvas
    main_frame_scrollbar = tk.Scrollbar(
        main_frame, orient=tk.VERTICAL, command=canvas.yview
    )
    main_frame_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.config(yscrollcommand=main_frame_scrollbar.set)

    # Create another frame inside the canvas for all the widgets
    widgets_frame = tk.Frame(canvas)

    # Create a window inside the canvas to host the widgets_frame
    canvas_window = canvas.create_window((0, 0), window=widgets_frame, anchor="nw")

    # Configure the scrolling behaviour via event bindigs
    def configure_scroll_region(event):
        """
        This function configures the scroll region for a frame
        """
        canvas.configure(scrollregion=canvas.bbox("all"))

    def configure_frame_width(event):
        """
        This function configures the frame width for the canvas window
        """
        canvas.itemconfig(canvas_window, width=event.width)

    # Update the scrollbar whenever the size of the inner frame changes
    widgets_frame.bind("<Configure>", configure_scroll_region)

    # Update the inner frame width when the canvas size changes
    canvas.bind("<Configure>", configure_frame_width)

    # Fetch the tables
    tables_list = get_table_names()

    # Define a variable for the row number used when placing items on the grid
    grid_row = 0

    # -------------------------------------------------------------------------------------------------------------------------------
    """SECTION FOR TABLES SEARCH BOX"""
    # Tables search box
    tables_searchbox_title = tk.Label(
        widgets_frame,
        text="Enter the Table Name to Search For:",
        font=("Arial", 18, "bold underline"),
    )
    tables_searchbox_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Define the string variable to hold the text entered into the search box
    text_var = tk.StringVar()
    # Record the changes in the text_var
    text_var.trace_add(
        "write",
        lambda *args: on_change_search_text(
            tables_listbox,
            tables_list,
            text_var,
            check_var_supp,
            check_var_unanon_supp,
            check_var_anon_followup,
            check_var_unanon_followup,
            tables_listbox_result,
            combined_selections
        ),
    )
    # Add the input box for the table search
    tables_searchbox_entry = tk.Entry(
        widgets_frame, font=("Arial", 12), textvariable=text_var
    )
    tables_searchbox_entry.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR TABLES LISTBOX"""
    # Listbox table selection
    tables_listbox_title = tk.Label(
        widgets_frame,
        text="Select the Table(s) to Query:",
        font=("Arial", 18, "bold underline"),
    )
    tables_listbox_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Add a frame for the listbox for table selection
    tables_frame = tk.Frame(widgets_frame)
    tables_frame.grid(row=grid_row, column=0, padx=10, pady=10)
    # Add a scrollbar to the listbox
    tables_listbox_scrollbar = tk.Scrollbar(tables_frame, orient=tk.VERTICAL)
    # Create the listbox and configure the scrollbar
    tables_listbox = tk.Listbox(
        tables_frame,
        yscrollcommand=tables_listbox_scrollbar.set,
        width=80,
        height=40,
        selectmode="multiple",
        exportselection=False,
        font=("Arial", 12),
    )
    tables_listbox.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    tables_listbox_scrollbar.config(command=tables_listbox.yview)
    tables_listbox_scrollbar.grid(row=grid_row, column=1, padx=10, pady=10, sticky="ns")
    # Sort the tables_list by the table name
    tables_list = tables_list.sort_values(by=["table_schema", "table_name"])
    # Populate the Listbox with table names
    for idx, row in tables_list.iterrows():
        schema = row["table_schema"]
        tbl = row["table_name"]
        tables_listbox.insert(tk.END, f"{schema}.{tbl}")
    # Bind the listbox to the table selector helper function
    tables_listbox.bind(
        "<<ListboxSelect>>",
        lambda event: on_tables_listbox_select(
            event, tables_listbox, tables_listbox_result
        ),
    )
    # Bind the listbox to the Prelim checkbox
    tables_listbox.bind(
        "<<ListboxSelect>>",
        lambda event: update_prelim_checkbox_state(
            tables_listbox, select_prelim_cols_checkbox, check_var_select_prelim_cols
        ),
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR FILTERING FOR ANONYMISED SUPPLEMENTARY TABLES"""
    # Add a tick box to filter for just the supplementary tables
    check_var_supp = IntVar()
    show_supp_tbls_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Show Anonymised Supplementary and Other Tables Only",
        variable=check_var_supp,
        command=lambda: on_change_supp_checkbox(
            check_var_supp,
            check_var_unanon_supp,
            check_var_anon_followup,
            check_var_unanon_followup,
            tables_listbox,
            tables_list,
            text_var,
            tables_listbox_result,
            combined_selections
        ),
        font=("Arial", 12),
    )
    show_supp_tbls_checkbox.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR FILTERING FOR UNANONYMISED SUPPLEMENTARY TABLES"""
    # Add a tick box to filter for just the supplementary tables
    check_var_unanon_supp = IntVar()
    show_unanon_supp_tbls_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Show Unanonymised Supplementary and Tables Only",
        variable=check_var_unanon_supp,
        command=lambda: on_change_supp_checkbox(
            check_var_supp,
            check_var_unanon_supp,
            check_var_anon_followup,
            check_var_unanon_followup,
            tables_listbox,
            tables_list,
            text_var,
            tables_listbox_result,
            combined_selections
        ),
        font=("Arial", 12),
    )
    show_unanon_supp_tbls_checkbox.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR FILTERING FOR ANONYMISED FOLLOW UP TABLES"""
    # Add a tick box to filter for just the anonymised follow up tables
    check_var_anon_followup = IntVar()
    show_anon_followup_tbls_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Show Anonymised Follow Up and Prelim Tables Only",
        variable=check_var_anon_followup,
        command=lambda: on_change_anon_followup_checkbox(
            check_var_anon_followup,
            check_var_supp,
            check_var_unanon_supp,
            check_var_unanon_followup,
            tables_listbox,
            tables_list,
            text_var,
            tables_listbox_result,
            combined_selections
        ),
        font=("Arial", 12),
    )
    show_anon_followup_tbls_checkbox.grid(
        row=grid_row, column=0, padx=15, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR FILTERING FOR UNANONYMISED FOLLOW UP TABLES"""
    # Add a tick box to filter for just the unanonymised follow up tables
    check_var_unanon_followup = IntVar()
    show_unanon_followup_tbls_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Show Unanonymised Follow Up and Prelim Tables Only",
        variable=check_var_unanon_followup,
        command=lambda: on_change_unanon_followup_checkbox(
            check_var_unanon_followup,
            check_var_supp,
            check_var_unanon_supp,
            check_var_anon_followup,
            tables_listbox,
            tables_list,
            text_var,
            tables_listbox_result,
            combined_selections
        ),
        font=("Arial", 12),
    )
    show_unanon_followup_tbls_checkbox.grid(
        row=grid_row, column=0, padx=15, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1
    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR SELECTING ALL TABLES"""
    # Add a "Select All" Checkbox
    check_var_select_all_tables = IntVar()
    select_all_tables_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Select All Tables",
        variable=check_var_select_all_tables,
        command=lambda: on_select_deselect_all_tables(
            check_var_select_all_tables, tables_listbox
        ),
        font=("Arial", 12),
    )
    select_all_tables_checkbox.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1
    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR TABLES SEARCH CONFIRMATION BUTTON"""
    # Tables search box confirmation button
    tables_searchbox_button = tk.Button(
        widgets_frame,
        command=lambda: on_select_searchbox_confirmation(
            tables_listbox, columns_listbox
        ),
        text="Confirm Table Search",
        font=("Arial", 12, "bold"),
    )
    tables_searchbox_button.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR A BUTTON TO ADD THE SELECTIONS FROM THE SEARCH RESULT TO THE SELECTED TABLES"""
    # Add to selected tables button
    add_tables_to_selection_button = tk.Button(widgets_frame, text="Add Tables to Selection", 
                                               command=lambda: on_click_add_to_selection(tables_listbox, tables_listbox_result, combined_selections), 
                                               font=("Arial", 12, "bold"))
    add_tables_to_selection_button.grid(row=grid_row, column=0, padx=10, pady=10, sticky='ew')
    # Increment the row
    grid_row+=1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR RESETTING THE TABLES SELECTION LISTBOX"""
    # Add a button for resetting the tables selection listbox
    reset_tables_selection_button = tk.Button(
        widgets_frame,
        command=lambda: on_select_reset_tables(
            tables_listbox, tables_listbox_result, columns_listbox
        ),
        text="Reset Tables Selection",
        font=("Arial", 12, "bold"),
    )
    reset_tables_selection_button.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE RESULTS LABEL FROM THE TABLES LISTBOX"""
    # Add the Result Label
    tables_listbox_result = tk.Label(
        widgets_frame,
        text="No table selected yet.",
        fg="blue",
        font=("Arial", 12, "italic"),
    )
    tables_listbox_result.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR COLUMNS LISTBOX"""
    # Listbox column selection
    columns_listbox_title = tk.Label(
        widgets_frame, text="Select Column(s) to Query:", font=("Arial", 18, "bold")
    )
    columns_listbox_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Add a frame for the listbox for column selection
    columns_frame = tk.Frame(widgets_frame)
    columns_frame.grid(row=grid_row, column=0, padx=10, pady=10)
    # Add a scrollbar to the listbox given the list of columns is long
    columns_listbox_scrollbar = tk.Scrollbar(columns_frame, orient=tk.VERTICAL)
    # Create the listbox for the column names and configure the scrollbar
    columns_listbox = tk.Listbox(
        columns_frame,
        yscrollcommand=columns_listbox_scrollbar.set,
        width=80,
        height=40,
        selectmode="multiple",
        exportselection=False,
        font=("Arial", 12),
    )
    columns_listbox.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    columns_listbox_scrollbar.config(command=columns_listbox.yview)
    columns_listbox_scrollbar.grid(
        row=grid_row, column=1, padx=10, pady=10, sticky="ns"
    )
    # Bind the listbox to the column selector helper function
    columns_listbox.bind(
        "<<ListboxSelect>>",
        lambda event: on_columns_listbox_select(
            event, columns_listbox, columns_listbox_result
        ),
    )
    # Increment the row
    grid_row += 1
    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE RESULT LABEL FROM THE COLUMNS LISTBOX"""
    # Add the result label
    columns_listbox_result = tk.Label(
        widgets_frame,
        text="No columns selected yet.",
        fg="blue",
        font=("Arial", 12, "italic"),
    )
    columns_listbox_result.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR SELECTING ALL THE COLUMNS FROM THE LISTBOX"""
    # Add a "Select All" Checkbox
    check_var_select_all_cols = IntVar()
    select_all_cols_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Select All Columns",
        variable=check_var_select_all_cols,
        command=lambda: on_select_deselect_all_columns(
            check_var_select_all_cols, columns_listbox, columns_listbox_result
        ),
        font=("Arial", 12),
    )
    select_all_cols_checkbox.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1
    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR SELECTING THE PRELIM DATA COLUMNS FROM THE LISTBOX (WHEN APPLICABLE)"""
    # Checkbox for selecting the prelim data columns only
    check_var_select_prelim_cols = IntVar()
    select_prelim_cols_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Select Prelim Columns",
        variable=check_var_select_prelim_cols,
        command=lambda: on_change_prelim_columns_checkbox(
            check_var_select_prelim_cols,
            select_prelim_cols_checkbox,
            columns_listbox,
            check_var_select_prelim_cols,
            tables_listbox,
            combined_selections
        ),
        font=("Arial", 12),
    )
    select_prelim_cols_checkbox.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE IN CLAUSE"""
    # Define string variables for the entries
    id_selected = StringVar()

    # Entry widget title for the IN Clause
    id_in_clause_title = ttk.Label(
        widgets_frame,
        text="Enter the ID Column Name for the IN Clause:",
        font=("Arial", 18, "bold"),
    )
    id_in_clause_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Entry widget for ID IN Clause
    id_in_clause_entry = ttk.Entry(
        widgets_frame, width=30, textvariable=id_selected, font=("Arial", 12)
    )
    id_in_clause_entry.grid(row=grid_row, column=0, sticky="ew")
    # Increment the row
    grid_row += 1
    # Entry widget title for the list of IDs for the IN Clause
    id_list_in_clause_title = ttk.Label(
        widgets_frame,
        text="Paste the ID List for the IN Clause:",
        font=("Arial", 18, "bold"),
    )
    id_list_in_clause_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Entry widget for the list of IDs for the IN Clause
    id_list_in_clause_entry = scrolledtext.ScrolledText(
        widgets_frame, width=30, height=40, undo=True, font=("Arial", 12)
    )
    id_list_in_clause_entry.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE IN CLAUSE (EXCLUSION)"""
    # Define string variables for the entries
    id_exclusions = StringVar()

    # Entry widget title for the IN Clause
    id_exclusions_in_clause_title = ttk.Label(
        widgets_frame,
        text="Enter the ID Column Name for the Exclusion IN Clause:",
        font=("Arial", 18, "bold"),
    )
    id_exclusions_in_clause_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Entry widget for ID exclusion IN Clause
    id_exclusions_in_clause_entry = ttk.Entry(
        widgets_frame, width=30, textvariable=id_exclusions, font=("Arial", 12)
    )
    id_exclusions_in_clause_entry.grid(row=grid_row, column=0, sticky="ew")
    # Increment the row
    grid_row += 1
    # Entry widget title for the list of IDs for the IN Clause
    id_exclusions_list_in_clause_title = ttk.Label(
        widgets_frame,
        text="Paste the ID List for the Exclusion IN Clause:",
        font=("Arial", 18, "bold"),
    )
    id_exclusions_list_in_clause_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Entry widget for the list of IDs for the Exclusion IN Clause
    id_exclusions_in_clause_entry = scrolledtext.ScrolledText(
        widgets_frame, width=30, height=40, undo=True, font=("Arial", 12)
    )
    id_exclusions_in_clause_entry.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE CONDITIONS"""
    # Entry widget title for the conditions (WHERE Clause)
    where_clause_title = ttk.Label(
        widgets_frame, text="Enter the WHERE Clause", font=("Arial", 18, "bold")
    )
    where_clause_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1
    # Entry widget for the conditions (WHERE Clause)
    where_clause_entry = scrolledtext.ScrolledText(
        widgets_frame, width=30, height=5, undo=True, font=("Arial", 12)
    )
    where_clause_entry.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE ORDER BY CLAUSE"""
    # Entry widget title for the ORDER BY Clause
    order_by_clause_title = ttk.Label(
        widgets_frame, text="Enter the ORDER BY Clause", font=("Arial", 18, "bold")
    )
    order_by_clause_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Entry widget for the ORDER BY Clause
    order_by_clause_entry = scrolledtext.ScrolledText(
        widgets_frame, width=30, height=5, undo=True, font=("Arial", 12)
    )
    order_by_clause_entry.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR ADDITIONAL SELECTIONS TITLE"""
    # Label for selecting any options that apply
    additional_options_title = ttk.Label(
        widgets_frame,
        text="Please Select any Applicable Option(s)",
        font=("Arial", 18, "bold"),
    )
    additional_options_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1
    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE SUBTITLE FOR WHEN SUPPLEMENTARY DATA IS BEING QUERIED"""
    # Label for querying supplementary data
    querying_supp_data_title = ttk.Label(
        widgets_frame,
        text="When Querying Supplementary Data",
        font=("Arial", 12, "italic bold"),
    )
    querying_supp_data_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR WHETHER SUPPLEMENTARY DATA IS BEING QUERIED"""
    # Checkbutton to tick when supplementary data is being queried
    check_var_is_supp_query = IntVar()
    show_is_supp_query_checkbox = tk.Checkbutton(
        widgets_frame,
        text="Please Tick if Querying Supplementary Data Tables",
        variable=check_var_is_supp_query,
        font=("Arial", 12),
        command=lambda: get_checkbox_isSuppQuery(
            check_var_is_supp_query,
            show_is_var_returnFirstAvailableRecord,
            show_is_var_returnFollowUpOnly_query,
            show_is_var_master_query,
            show_is_var_returnFirstBlock_query,
        ),
    )
    show_is_supp_query_checkbox.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="w"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE SUBTITLE FOR SELECTING BOTH THE FOLLOW UP AND SUPPLEMENTARY DATA"""
    # Label for querying both follow up and supplementary data
    querying_followup_supp_data_title = ttk.Label(
        widgets_frame,
        text="When Querying Both Follow Up and Supplementary Data",
        font=("Arial", 12, "italic bold"),
    )
    querying_followup_supp_data_title.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR TOGGLING COMBINING THE FOLLOW UP AND SUPPLEMENTARY DATA"""
    # Checkbutton to tick when wanting to combine the follow up and supplementary data
    check_var_combine_followup_supp_data = IntVar()
    show_is_var_combine_followup_supp_data = tk.Checkbutton(
        widgets_frame,
        text="Please tick to Combine Anonymised Follow Up and Supplementary data",
        variable=check_var_combine_followup_supp_data,
        font=("Arial", 12),
        command=lambda: get_checkbox_combine_followup_supp_data(
            check_var_combine_followup_supp_data, selected_tables
        ),
    )
    show_is_var_combine_followup_supp_data.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="w"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE SUBTITLE FOR WHEN MASTER LIST IS BEING QUERIED"""
    # Label for querying the master list
    querying_master_list_title = ttk.Label(
        widgets_frame,
        text="When Querying the Master List",
        font=("Arial", 12, "italic bold"),
    )
    querying_master_list_title.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="ew"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR TOGGLING QUERYING FROM THE MASTER LIST (LONG)"""
    # Checkbutton to tick when the master list (long version) is being queried
    check_var_master_query = IntVar()
    show_is_var_master_query = tk.Checkbutton(
        widgets_frame,
        text="Please tick if Querying the Master List",
        variable=check_var_master_query,
        font=("Arial", 12),
        command=lambda: get_checkbox_isMasterList(
            check_var_master_query,
            show_is_supp_query_checkbox,
            show_is_var_master_query,
        ),
    )
    show_is_var_master_query.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE OPTION TO RETURN THE FIRST BLOCK OF A SPECIFIC TYPE FROM EACH DONOR (MASTER LIST)"""
    # Checkbutton to tick when wanting to return the first block of a specific type from each donor
    check_var_returnFirstBlock_query = IntVar()
    show_is_var_returnFirstBlock_query = tk.Checkbutton(
        widgets_frame,
        text="Please tick to Return the First Primary or Met Block From Each Donor",
        variable=check_var_returnFirstBlock_query,
        font=("Arial", 12),
        command=lambda: get_checkbox_returnFirstBlock(
            check_var_returnFirstBlock_query, show_is_supp_query_checkbox
        ),
    )
    show_is_var_returnFirstBlock_query.grid(row=grid_row, column=0, padx=10, sticky="w")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE OPTION TO RETURN DATA FROM FOLLOW UP TABLES ONLY (MASTER LIST)"""
    # Checkbutton to tick when wanting to return data from the follow up tables only - when querying the master list
    check_var_returnFollowUpOnly_query = IntVar()
    show_is_var_returnFollowUpOnly_query = tk.Checkbutton(
        widgets_frame,
        text="Please tick to Return data from the Follow Up tables only",
        variable=check_var_returnFollowUpOnly_query,
        font=("Arial", 12),
        command=lambda: get_checkbox_returnFollowUp_Only(
            check_var_returnFollowUpOnly_query,
            show_is_var_returnFirstAvailableRecord,
            show_is_supp_query_checkbox,
        ),
    )
    show_is_var_returnFollowUpOnly_query.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="w"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE OPTION TO RETURN DATA FROM THE FIRST AVAILABLE RECORD (PRELIM OR FOLLOW UP) (MASTER LIST)"""
    # Checkbutton to tick when wanting to reutn the first available record - when querying the master list
    check_var_returnFirstAvailableRecord = IntVar()
    show_is_var_returnFirstAvailableRecord = tk.Checkbutton(
        widgets_frame,
        text="Please tick to Return the first avaialble record (Prelim only or Follow Up) from each Donor",
        variable=check_var_returnFirstAvailableRecord,
        font=("Arial", 12),
        command=lambda: get_checkbox_returnFirstAvailableRecord(
            check_var_returnFirstAvailableRecord,
            show_is_var_returnFollowUpOnly_query,
            show_is_supp_query_checkbox,
        ),
    )
    show_is_var_returnFirstAvailableRecord.grid(
        row=grid_row, column=0, padx=10, pady=10, sticky="w"
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR LIMITING THE OUTPUT TO A SPECIFIC NUMBER OF ROWS"""
    # Entry widget title for limiting the number of rows in the output
    top_entry_title = ttk.Label(
        widgets_frame,
        text="Enter the Number of Rows to Show in the Output",
        font=("Arial", 18, "bold"),
    )
    top_entry_title.grid(row=grid_row, column=0, padx=10, pady=10, sticky="w")
    # Increment the row
    grid_row += 1
    # Entry widget for limiting the number of rows in the output
    top_entry = tk.Entry(widgets_frame, width=30, font=("Arial", 14))
    top_entry.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR THE CONFIRMATION BUTTON"""
    # Confirmation button to run the query
    run_query = tk.Button(widgets_frame, text="Run Query", font=("Arial", 12, "bold"))
    run_query.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    run_query.config(
        command=lambda: callback_run_query(
            tables_listbox,
            columns_listbox,
            id_selected,
            id_list_in_clause_entry,
            id_exclusions,
            id_exclusions_in_clause_entry,
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
        )
    )
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------

    """SECTION FOR BUTTON TO RESET EVERYTHING"""
    # Button to reset all widgets
    reset_button = tk.Button(
        widgets_frame, text="Reset All Widgets", font=("Arial", 12, "bold")
    )
    reset_button.grid(row=grid_row, column=0, padx=10, pady=10, sticky="ew")
    reset_button.config(command=lambda: on_select_reset_widgets(root))
    # Increment the row
    grid_row += 1

    # -------------------------------------------------------------------------------------------------------------------------------


# Start the GUI loop
if __name__ == "__main__":
    build_ui()
    root.mainloop()
