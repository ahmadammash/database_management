import tkinter as tk
from tkinter import ttk

import pandas as pd

# Import functions from the project directory
from code_files.utilities.io.setup_db_cnxn import setup_cnxn
from code_files.utilities.ui.callbacks import callback_run_query
from code_files.utilities.ui.fetch_information import get_table_names
from code_files.utilities.ui.get_widget_values import (
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
    on_change_unanon_supp_checkbox,
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


class QueryBuilderApp:
    def __init__(self, root):
        self.root = tk.Tk()
        self.root.title("SQL Server Table Viewer")
        self.root.geometry("1200x850")

        # Shared state (replaces all global variables)
        self.project_root, self.cnxn, self.cursor, self.engine, _ = setup_cnxn()
        self.tables_list = get_table_names().sort_values(by=["table_schema", "table_name"])
        self.combined_selections = set()
        self.df_result = pd.DataFrame()

        # Build a 2-column layout
        self._build_layout()

    def _build_layout(self):
        # Left pane: tables and columns | right pane: IDs, filters, and options
        left_frame = ttk.Frame(self.root, padding=10)
        left_frame.grid(row=0, column=0, sticky="nsew")

        right_frame = ttk.Frame(self.root, padding=10)
        right_frame.grid(row=0, column=1, sticky="nsew")

        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        self._build_left_pane(left_frame)
        self._build_right_pane(right_frame)
    
    # def _build_left_pane(self):





if __name__ == "__main__":
    root = tk.Tk()
    app = QueryBuilderApp(root)
    root.mainloop()

