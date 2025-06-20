import tkinter as tk
from tkinter import ttk, filedialog, messagebox,simpledialog,  font as tkFont
import xml.etree.ElementTree as ET
import csv
import threading
from collections import Counter, deque, defaultdict
import traceback
import os
import json
import re
from datetime import datetime
import uuid
from lxml import etree
# from tkinter import


# --- Global Constants ---
CHUNK_SIZE = 1024 * 1024
MIN_ROWS_FOR_TABLE = 3
MIN_PERCENT_SIMILAR = 0.6
UNDO_STACK_SIZE = 20
VIRTUAL_TABLE_ROW_COUNT = 100  # Number of rows to display at once



class HelpWindow(tk.Toplevel):
    """
    A Toplevel window that displays a markdown-formatted help file
    with a clickable table of contents.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Help")
        self.geometry("800x600")
        self.transient(parent)

        self.toc_map = {}

        self._setup_widgets()
        self._load_and_parse_help_doc()

    def _setup_widgets(self):
        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)

        # --- TOC Frame ---
        toc_frame = ttk.Frame(main_pane, width=200)
        toc_frame.pack_propagate(False)
        ttk.Label(toc_frame, text="Table of Contents", font=("Segoe UI", 10, "bold")).pack(pady=5)
        self.toc_listbox = tk.Listbox(toc_frame, selectbackground="#0078D7", selectforeground="white")
        self.toc_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.toc_listbox.bind("<<ListboxSelect>>", self._on_toc_select)
        main_pane.add(toc_frame, weight=1)

        # --- Content Frame ---
        content_frame = ttk.Frame(main_pane)
        self.content_text = tk.Text(content_frame, wrap=tk.WORD, padx=10, pady=10, font=("Segoe UI", 10))
        content_vsb = ttk.Scrollbar(content_frame, orient="vertical", command=self.content_text.yview)
        self.content_text.configure(yscrollcommand=content_vsb.set)
        content_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        main_pane.add(content_frame, weight=3)

        # --- Configure Text Tags ---
        self.content_text.tag_configure("h1", font=("Segoe UI", 18, "bold"), spacing3=10)
        self.content_text.tag_configure("h2", font=("Segoe UI", 14, "bold"), spacing3=8, lmargin1=10)
        self.content_text.tag_configure("h3", font=("Segoe UI", 11, "bold"), spacing3=5, lmargin1=20)
        self.content_text.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        self.content_text.tag_configure("code", font=("Courier New", 10), background="#f0f0f0")
        self.content_text.tag_configure("li", lmargin1=30, lmargin2=45)

    def _load_and_parse_help_doc(self):
        try:
            with open("help.md", "r", encoding="utf-8") as f:
                content = f.readlines()
        except FileNotFoundError:
            self.content_text.insert(tk.END,
                                     "Error: help.md not found.\n\nPlease create this file in the same directory as the application.",
                                     "h1")
            return

        is_toc_section = False
        for line in content:
            stripped_line = line.strip()

            if stripped_line == "## Table of Contents":
                is_toc_section = True
                continue
            elif is_toc_section and stripped_line.startswith("*"):
                match = re.match(r'\*\s*\[(.*)\]\(#.*\)', stripped_line)
                if match:
                    toc_entry = match.group(1)
                    self.toc_listbox.insert(tk.END, toc_entry)
                continue
            elif is_toc_section and not stripped_line:
                is_toc_section = False
                continue

            if is_toc_section:
                continue

            if stripped_line.startswith("# "):
                tag, text = "h1", stripped_line[2:]
                self._add_text(text, tag)
            elif stripped_line.startswith("## "):
                tag, text = "h2", stripped_line[3:]
                self._add_text(text, tag, add_to_toc=True)
            elif stripped_line.startswith("### "):
                tag, text = "h3", stripped_line[4:]
                self._add_text(text, tag)
            elif stripped_line.startswith("* "):
                tag, text = "li", "• " + stripped_line[2:]
                self._add_text(text + "\n", tag)
            elif not stripped_line:
                self.content_text.insert(tk.END, "\n")
            else:
                self._add_text(line)

        self.content_text.config(state=tk.DISABLED)

    def _add_text(self, text, tag=None, add_to_toc=False):
        start_index = self.content_text.index(tk.END)

        parts = re.split(r'(`[^`]+`)', text)
        for i, part in enumerate(parts):
            if i % 2 == 1:
                self.content_text.insert(tk.END, part[1:-1], "code")
            else:
                bold_parts = re.split(r'(\*\*[^*]+\*\*)', part)
                for j, bold_part in enumerate(bold_parts):
                    if j % 2 == 1:
                        self.content_text.insert(tk.END, bold_part[2:-2], "bold")
                    else:
                        self.content_text.insert(tk.END, bold_part)

        if tag:
            self.content_text.tag_add(tag, start_index, tk.END)

        if tag not in ["li"]:
            self.content_text.insert(tk.END, "\n")

        if add_to_toc:
            self.toc_map[text] = start_index

    def _on_toc_select(self, event):
        selection_indices = self.toc_listbox.curselection()
        if not selection_indices:
            return

        selected_text = self.toc_listbox.get(selection_indices[0])
        if selected_text in self.toc_map:
            self.content_text.see(self.toc_map[selected_text])



class TabReorderDialog(tk.Toplevel):
    """
    A dialog to reorder the open tabs in the main notebook.
    """

    def __init__(self, parent, notebook):
        super().__init__(parent)
        self.notebook = notebook
        self.transient(parent)
        self.title("Reorder Tabs")
        self.geometry("350x300")
        self.grab_set()

        self._setup_widgets()
        self._populate_tabs()

    def _setup_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=0, column=0, sticky="nsew", rowspan=3)

        self.tab_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE)
        self.tab_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tab_listbox.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tab_listbox.configure(yscrollcommand=vsb.set)

        up_button = ttk.Button(main_frame, text="Move Up", command=self._move_up)
        up_button.grid(row=0, column=1, sticky="ew", padx=(10, 0))

        down_button = ttk.Button(main_frame, text="Move Down", command=self._move_down)
        down_button.grid(row=1, column=1, sticky="ew", padx=(10, 0))

        apply_button = ttk.Button(main_frame, text="Apply & Close", command=self.destroy)
        apply_button.grid(row=3, column=0, columnspan=2, sticky="e", pady=(10, 0))

    def _populate_tabs(self):
        self.tab_listbox.delete(0, tk.END)
        # We skip tab 0, which is the "Node Details" tab
        for i in range(1, len(self.notebook.tabs())):
            tab_text = self.notebook.tab(i, "text")
            self.tab_listbox.insert(tk.END, tab_text)

    def _move_up(self):
        selection = self.tab_listbox.curselection()
        if not selection:
            return

        pos = selection[0]
        if pos == 0:
            return

        # Move in the listbox
        text = self.tab_listbox.get(pos)
        self.tab_listbox.delete(pos)
        self.tab_listbox.insert(pos - 1, text)
        self.tab_listbox.selection_set(pos - 1)

        # Move in the actual notebook
        notebook_pos = pos + 1
        tab_widget_id = self.notebook.tabs()[notebook_pos]
        self.notebook.insert(notebook_pos - 1, tab_widget_id)

    def _move_down(self):
        selection = self.tab_listbox.curselection()
        if not selection:
            return

        pos = selection[0]
        if pos == self.tab_listbox.size() - 1:
            return

        # Move in the listbox
        text = self.tab_listbox.get(pos)
        self.tab_listbox.delete(pos)
        self.tab_listbox.insert(pos + 1, text)
        self.tab_listbox.selection_set(pos + 1)

        # Move in the actual notebook
        notebook_pos = pos + 1
        tab_widget_id = self.notebook.tabs()[notebook_pos]
        self.notebook.insert(notebook_pos + 1, tab_widget_id)


class BatchOperationsDialog(tk.Toplevel):
    """
    A dialog for performing batch update or delete operations on a table.
    """

    def __init__(self, parent, app, current_tab):
        super().__init__(parent)
        self.app = app
        self.current_tab = current_tab
        self.internal_key = current_tab.internal_key
        self.columns = current_tab.table_info["columns"]

        self.title("Batch Operations")
        self.transient(parent)
        self.grab_set()

        # --- UI Variables ---
        self.filter_col_var = tk.StringVar()
        self.filter_op_var = tk.StringVar(value="CONTAINS")
        self.filter_val_var = tk.StringVar()
        self.action_var = tk.StringVar(value="update")
        self.update_col_var = tk.StringVar()
        self.update_val_var = tk.StringVar()

        self._setup_widgets()

    def _setup_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Filter Frame ---
        filter_frame = ttk.Labelframe(main_frame, text="1. Find Rows Where", padding=10)
        filter_frame.pack(fill=tk.X, expand=True, pady=5)
        filter_frame.columnconfigure(1, weight=1)

        ttk.Label(filter_frame, text="Column:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Combobox(filter_frame, textvariable=self.filter_col_var, values=self.columns, state="readonly").grid(row=0,
                                                                                                                 column=1,
                                                                                                                 sticky="ew",
                                                                                                                 padx=5,
                                                                                                                 pady=2)

        op_values = ["CONTAINS", "EQUALS", "STARTS WITH", "ENDS WITH", "DOES NOT CONTAIN", "IS NOT EQUAL TO"]
        ttk.Combobox(filter_frame, textvariable=self.filter_op_var, values=op_values, state="readonly").grid(row=1,
                                                                                                             column=0,
                                                                                                             sticky="ew",
                                                                                                             padx=5,
                                                                                                             pady=2)
        ttk.Entry(filter_frame, textvariable=self.filter_val_var).grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # --- Action Frame ---
        action_frame = ttk.Labelframe(main_frame, text="2. Perform Action", padding=10)
        action_frame.pack(fill=tk.X, expand=True, pady=5)

        ttk.Radiobutton(action_frame, text="Update Rows", variable=self.action_var, value="update",
                        command=self._on_action_change).pack(anchor="w")
        ttk.Radiobutton(action_frame, text="Delete Rows", variable=self.action_var, value="delete",
                        command=self._on_action_change).pack(anchor="w")

        # --- Update Settings Frame ---
        self.update_settings_frame = ttk.Labelframe(main_frame, text="3. Update Settings", padding=10)
        self.update_settings_frame.pack(fill=tk.X, expand=True, pady=5)
        self.update_settings_frame.columnconfigure(1, weight=1)

        ttk.Label(self.update_settings_frame, text="Set Column:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Combobox(self.update_settings_frame, textvariable=self.update_col_var, values=self.columns,
                     state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(self.update_settings_frame, text="To Value:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(self.update_settings_frame, textvariable=self.update_val_var).grid(row=1, column=1, sticky="ew",
                                                                                     padx=5, pady=2)

        # --- Execution Frame ---
        exec_frame = ttk.Frame(main_frame)
        exec_frame.pack(fill=tk.X, expand=True, pady=10)

        self.status_label = ttk.Label(exec_frame, text="Click 'Preview' to see affected rows.")
        self.status_label.pack(side=tk.LEFT, expand=True, fill=tk.X)

        ttk.Button(exec_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        self.exec_button = ttk.Button(exec_frame, text="Execute", command=self._execute_changes, state="disabled")
        self.exec_button.pack(side=tk.RIGHT, padx=5)
        ttk.Button(exec_frame, text="Preview", command=self._preview_changes).pack(side=tk.RIGHT)

    def _on_action_change(self):
        state = "normal" if self.action_var.get() == "update" else "disabled"
        for child in self.update_settings_frame.winfo_children():
            child.config(state=state)

    def _get_matching_rows(self):
        filter_col = self.filter_col_var.get()
        op = self.filter_op_var.get()
        val = self.filter_val_var.get()

        # A column must be selected, but the value can be empty.
        if not filter_col:
            return []

        master_data = self.app.table_data_cache[self.internal_key]
        matching_indices = []
        for i, row in enumerate(master_data):
            cell_value = str(row.get(filter_col, ""))
            match = False
            if op == "CONTAINS":
                match = val in cell_value
            elif op == "EQUALS":
                match = val == cell_value
            elif op == "STARTS WITH":
                match = cell_value.startswith(val)
            elif op == "ENDS WITH":
                match = cell_value.endswith(val)
            elif op == "DOES NOT CONTAIN":
                match = val not in cell_value
            elif op == "IS NOT EQUAL TO":
                match = val != cell_value

            if match:
                matching_indices.append(i)
        return matching_indices

    def _execute_delete(self):
        # Confirm with the user before deleting
        if not messagebox.askyesno("Confirm Delete",
                                   f"Are you sure you want to delete {len(self.matching_indices)} row(s)?\nThis action can be undone.",
                                   parent=self):
            return

        deleted_rows = []
        parent_element = self.current_tab.table_info.get("parent_element")

        for index in sorted(self.matching_indices, reverse=True):
            deleted_row_data = self.app.table_data_cache[self.internal_key].pop(index)
            deleted_rows.append({"index": index, "data": deleted_row_data})

            # Also remove from the actual XML tree if applicable
            if self.app.file_type == 'xml' and parent_element is not None:
                element_to_remove = deleted_row_data.get("_element")
                if element_to_remove is not None:
                    parent_element.remove(element_to_remove)

        if deleted_rows:
            self.app.push_undo(
                {"action": "batch_delete", "deleted_rows": deleted_rows, "internal_key": self.internal_key})
            self.current_tab._apply_filter_and_sort()
        self.destroy()

    def _preview_changes(self):
        self.matching_indices = self._get_matching_rows()
        count = len(self.matching_indices)
        action = self.action_var.get()
        if count > 0:
            self.status_label.config(text=f"This will {action} {count} row(s).")
            self.exec_button.config(state="normal")
        else:
            self.status_label.config(text="No matching rows found.")
            self.exec_button.config(state="disabled")

    def _execute_changes(self):
        action = self.action_var.get()
        if action == "update":
            self._execute_update()
        elif action == "delete":
            self._execute_delete()

    def _execute_update(self):
        update_col = self.update_col_var.get()
        update_val = self.update_val_var.get()
        if not update_col:
            messagebox.showerror("Error", "Please select a column to update.", parent=self)
            return

        changes = []
        for index in self.matching_indices:
            old_value = self.app.table_data_cache[self.internal_key][index][update_col]
            changes.append(
                {"original_index": index, "column": update_col, "old_value": old_value, "new_value": update_val})
            self.app.table_data_cache[self.internal_key][index][update_col] = update_val
            if self.app.file_type == 'xml':
                self.app.table_data_cache[self.internal_key][index]['_element'].find(update_col).text = update_val

        if changes:
            self.app.push_undo({"action": "batch_update", "changes": changes, "internal_key": self.internal_key})
            self.current_tab._apply_filter_and_sort()
        self.destroy()




class TableViewTab(ttk.Frame):
    def __init__(self, parent_notebook, app_instance, internal_key):
        super().__init__(parent_notebook)
        self.app = app_instance
        self.internal_key = internal_key
        self.table_info = self.app.potential_tables[internal_key]

        # State for this specific tab
        self.quick_filter_var = tk.StringVar()
        self.nav_status_var = tk.StringVar()
        self.sort_criteria = []
        self.current_view_data = []
        self.virtual_view_top_index = 0
        self.selected_data_index = None
        self.active_cell_editor = None
        self.right_clicked_column = None

        self._setup_widgets()
        self.display_table_view_data()

    def _setup_widgets(self):
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        # Header Context Menu for Statistics
        self.header_context_menu = tk.Menu(self, tearoff=0)
        self.header_context_menu.add_command(label="Show Column Statistics", command=self._calculate_and_show_stats)

        filter_frame = ttk.Frame(self)
        filter_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        ttk.Label(filter_frame, text="Quick Filter:").pack(side=tk.LEFT, padx=(0, 5))
        filter_entry = ttk.Entry(filter_frame, textvariable=self.quick_filter_var)
        filter_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.quick_filter_var.trace_add("write", self._on_quick_filter_change)

        self.table_treeview = ttk.Treeview(self, show='headings')
        self.table_treeview_vsb = ttk.Scrollbar(self, orient="vertical", command=self._on_virtual_scroll)
        self.table_treeview_hsb = ttk.Scrollbar(self, orient="horizontal", command=self.table_treeview.xview)
        self.table_treeview.configure(xscrollcommand=self.table_treeview_hsb.set)
        self.table_treeview.grid(row=1, column=0, sticky="nsew")
        self.table_treeview_vsb.grid(row=1, column=1, sticky="ns")
        self.table_treeview_hsb.grid(row=2, column=0, columnspan=1, sticky="ew")

        self.table_treeview.tag_configure('selected', background='#e6f3ff')
        self.table_treeview.bind("<Button-1>", self.on_table_tree_click)
        self.table_treeview.bind("<Button-3>", self._show_header_context_menu)
        self.table_treeview.bind("<Double-1>", self.on_table_cell_or_header_double_click)

        nav_frame = ttk.Frame(self, padding=2)
        nav_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        self.nav_btn_top = ttk.Button(nav_frame, text="<< Top", command=self._nav_top, state="disabled")
        self.nav_btn_top.pack(side=tk.LEFT, padx=2)
        self.nav_btn_prev = ttk.Button(nav_frame, text="< Prev", command=self._nav_prev, state="disabled")
        self.nav_btn_prev.pack(side=tk.LEFT, padx=2)
        self.nav_status_label = ttk.Label(nav_frame, textvariable=self.nav_status_var, anchor="center")
        self.nav_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.nav_btn_next = ttk.Button(nav_frame, text="Next >", command=self._nav_next, state="disabled")
        self.nav_btn_next.pack(side=tk.LEFT, padx=2)
        self.nav_btn_end = ttk.Button(nav_frame, text="End >>", command=self._nav_end, state="disabled")
        self.nav_btn_end.pack(side=tk.LEFT, padx=2)

    def _show_header_context_menu(self, event):
        region = self.table_treeview.identify("region", event.x, event.y)
        if region == "heading":
            column_id_str = self.table_treeview.identify_column(event.x)
            self.right_clicked_column = self.table_treeview.column(column_id_str, "id")
            if self.right_clicked_column != "#":
                self.header_context_menu.post(event.x_root, event.y_root)

    def _calculate_and_show_stats(self):
        if not self.right_clicked_column:
            return

        column_id = self.right_clicked_column
        all_values = [row.get(column_id, "") for row in self.current_view_data]

        total_rows = len(all_values)
        non_empty_values = [v for v in all_values if v is not None and str(v).strip() != ""]
        non_empty_count = len(non_empty_values)
        unique_values = set(non_empty_values)
        unique_count = len(unique_values)

        numeric_values = []
        for v in non_empty_values:
            try:
                numeric_values.append(float(v))
            except (ValueError, TypeError):
                continue

        stats_message = f"Statistics for Column: '{column_id}'\n"
        stats_message += "-" * 40
        stats_message += f"\nTotal Rows (in current view): {total_rows}"
        stats_message += f"\nNon-Empty Rows: {non_empty_count}"
        stats_message += f"\nUnique Values: {unique_count}"

        if numeric_values:
            stats_message += "\n\n--- Numeric Stats ---"
            stats_message += f"\nSum: {sum(numeric_values):,.2f}"
            stats_message += f"\nMean (Average): {sum(numeric_values) / len(numeric_values):,.2f}"
            stats_message += f"\nMinimum: {min(numeric_values):,.2f}"
            stats_message += f"\nMaximum: {max(numeric_values):,.2f}"
        else:
            stats_message += "\n\n(No numeric data found for stats)"

        messagebox.showinfo(f"Column Statistics", stats_message, parent=self)

    def display_table_view_data(self):
        if self.app.file_type == 'xml' and self.internal_key not in self.app.table_data_cache:
            parent_element = self.table_info["parent_element"]
            row_tag = self.table_info["row_tag"]
            row_elements = parent_element.findall(row_tag)
            parsed_data = []
            for i, row_element in enumerate(row_elements):
                row_data = {"_element": row_element, "_original_index": i}
                for col_name in self.table_info["columns"]:
                    col_data_element = row_element.find(col_name)
                    cell_text = col_data_element.text.strip() if col_data_element is not None and col_data_element.text else ""
                    row_data[col_name] = cell_text
                parsed_data.append(row_data)
            self.app.table_data_cache[self.internal_key] = parsed_data
        elif self.app.file_type == 'csv':
            if not self.app.table_data_cache[self.internal_key] or not any(
                    "_original_index" in row for row in self.app.table_data_cache[self.internal_key]):
                for i, row in enumerate(self.app.table_data_cache[self.internal_key]):
                    row["_original_index"] = i

        columns = self.table_info["columns"]
        display_columns = ["#"] + columns
        self.table_treeview["columns"] = display_columns
        self.table_treeview.column("#0", width=0, stretch=tk.NO)
        self.table_treeview.column("#", width=60, stretch=tk.NO, anchor='e')
        self.table_treeview.heading("#", text="#", anchor='e')
        for col_name in columns:
            self.table_treeview.column(col_name, width=100, anchor='w')
            self.table_treeview.heading(col_name, text=col_name, anchor='w')

        self.quick_filter_var.set("")
        self._apply_filter_and_sort()
        self.resize_columns()

    def _repopulate_virtual_table(self):
        for item in self.table_treeview.get_children():
            self.table_treeview.delete(item)

        start_index = self.virtual_view_top_index
        end_index = start_index + VIRTUAL_TABLE_ROW_COUNT
        data_slice = self.current_view_data[start_index:end_index]

        columns = self.table_info["columns"]
        all_display_columns = ["#"] + columns

        for i, row_data in enumerate(data_slice):
            actual_view_index = start_index + i
            row_data["#"] = str(actual_view_index + 1)
            display_values = [row_data.get(col_id, "") for col_id in all_display_columns]

            tags = (str(actual_view_index),)
            if self.selected_data_index == actual_view_index:
                tags += ('selected',)

            self.table_treeview.insert("", "end", values=display_values, tags=tags)

        self._update_nav_controls()

    def _update_virtual_scrollbar(self):
        total_rows = len(self.current_view_data)
        if total_rows <= VIRTUAL_TABLE_ROW_COUNT:
            self.table_treeview_vsb.set(0, 1)
        else:
            upper = self.virtual_view_top_index / total_rows if total_rows > 0 else 0
            lower = (self.virtual_view_top_index + VIRTUAL_TABLE_ROW_COUNT) / total_rows if total_rows > 0 else 1
            self.table_treeview_vsb.set(upper, lower)

    def _on_virtual_scroll(self, action, value, units=None):
        total_rows = len(self.current_view_data)
        if total_rows <= VIRTUAL_TABLE_ROW_COUNT: return

        max_top_index = total_rows - VIRTUAL_TABLE_ROW_COUNT
        if action == "moveto":
            new_top_index = int(float(value) * max_top_index)
        elif action == "scroll":
            if units == "pages":
                new_top_index = self.virtual_view_top_index + (int(value) * VIRTUAL_TABLE_ROW_COUNT)
            else:
                new_top_index = self.virtual_view_top_index + int(value)
        else:
            return

        new_top_index = max(0, min(new_top_index, max_top_index))
        if new_top_index != self.virtual_view_top_index:
            self.virtual_view_top_index = new_top_index
            self._repopulate_virtual_table()
            self._update_virtual_scrollbar()

    def _jump_to_virtual_index(self, data_index):
        total_rows = len(self.current_view_data)
        if not (0 <= data_index < total_rows): return

        self.virtual_view_top_index = max(0, min(data_index, total_rows - VIRTUAL_TABLE_ROW_COUNT))
        self.select_row_by_index(data_index)
        self._repopulate_virtual_table()
        self._update_virtual_scrollbar()

        for item_id in self.table_treeview.get_children():
            tags = self.table_treeview.item(item_id, "tags")
            if tags and int(tags[0]) == data_index:
                self.table_treeview.see(item_id)
                self.table_treeview.focus(item_id)
                break

    def on_table_tree_click(self, event):
        if self.active_cell_editor and event.widget != self.active_cell_editor:
            self._finish_cell_edit()
        region = self.table_treeview.identify("region", event.x, event.y)
        if region == "heading":
            self.on_table_header_click_for_sort(event)
        elif region == "cell":
            item_id = self.table_treeview.identify_row(event.y)
            if item_id:
                tags = self.table_treeview.item(item_id, "tags")
                if tags:
                    view_index = int(tags[0])
                    self.select_row_by_index(view_index)

    def on_table_cell_or_header_double_click(self, event):
        if self.active_cell_editor: self._finish_cell_edit()
        region = self.table_treeview.identify("region", event.x, event.y)
        if region == "heading":
            self.resize_columns()
            return

        item_id = self.table_treeview.identify_row(event.y)
        column_id_str = self.table_treeview.identify_column(event.x)
        if not item_id or not column_id_str: return

        actual_column_name = self.table_treeview.column(column_id_str, "id")
        if actual_column_name == "#": return

        tags = self.table_treeview.item(item_id, "tags")
        if not tags: return
        view_index = int(tags[0])
        self.select_row_by_index(view_index)

        bbox = self.table_treeview.bbox(item_id, column_id_str)
        if not bbox: return
        x, y, width, height = bbox
        original_value = self.table_treeview.item(item_id, "values")[
            self.table_treeview["columns"].index(actual_column_name)]
        self.active_cell_editor = ttk.Entry(self.table_treeview)
        self.active_cell_editor.insert(0, original_value)
        self.active_cell_editor.place(x=x, y=y, width=width, height=height, anchor='nw')
        self.active_cell_editor.focus_set()
        self.active_cell_editor.select_range(0, 'end')

        self.active_cell_editor.view_index = view_index
        self.active_cell_editor.column_name = actual_column_name
        self.active_cell_editor.bind("<Return>", lambda e: self._finish_cell_edit(commit=True))
        self.active_cell_editor.bind("<Escape>", lambda e: self._finish_cell_edit(commit=False))

    def select_row_by_index(self, view_index):
        if self.selected_data_index == view_index: return
        self.selected_data_index = view_index
        self._repopulate_virtual_table()

    def deselect_row(self, update_view=True):
        if self.selected_data_index is not None:
            self.selected_data_index = None
            if update_view: self._repopulate_virtual_table()
        self._update_nav_controls()

    def on_table_header_click_for_sort(self, event):
        column_id_str = self.table_treeview.identify_column(event.x)
        column_id = self.table_treeview.column(column_id_str, "id")
        if column_id == "#": return

        is_shift_click = (event.state & 0x0001) != 0
        if not is_shift_click:
            if self.sort_criteria and self.sort_criteria[0][0] == column_id:
                new_dir = 'desc' if self.sort_criteria[0][1] == 'asc' else 'asc'
                self.sort_criteria = [(column_id, new_dir)]
            else:
                self.sort_criteria = [(column_id, 'asc')]
        else:
            col_index = next((i for i, (col, _) in enumerate(self.sort_criteria) if col == column_id), -1)
            if col_index != -1:
                new_dir = 'desc' if self.sort_criteria[col_index][1] == 'asc' else 'asc'
                self.sort_criteria[col_index] = (column_id, new_dir)
            else:
                self.sort_criteria.append((column_id, 'asc'))
        self._apply_filter_and_sort()

    def _finish_cell_edit(self, commit=False):
        if not self.active_cell_editor: return
        editor = self.active_cell_editor
        self.active_cell_editor = None

        view_index, column, new_value = editor.view_index, editor.column_name, editor.get()
        editor.destroy()
        if not commit: return

        original_index = self.current_view_data[view_index]["_original_index"]
        old_value = self.app.table_data_cache[self.internal_key][original_index][column]

        if new_value != old_value:
            self.app.table_data_cache[self.internal_key][original_index][column] = new_value
            self.current_view_data[view_index][column] = new_value
            self.app.push_undo({
                "action": "edit", "column": column, "old_value": old_value,
                "new_value": new_value, "internal_key": self.internal_key,
                "original_index": original_index
            })
            if self.app.file_type == 'xml':
                element = self.app.table_data_cache[self.internal_key][original_index]['_element']
                col_data = element.find(column)
                if col_data is not None:
                    col_data.text = new_value
                else:
                    ET.SubElement(element, column).text = new_value
            self._repopulate_virtual_table()

    def _on_quick_filter_change(self, *args):
        self._apply_filter_and_sort()

    def _apply_filter_and_sort(self):
        if self.internal_key not in self.app.potential_tables:
            self.current_view_data = []
            self._update_virtual_table_view()
            return

        master_data = self.app.table_data_cache.get(self.internal_key, [])
        filter_text = self.quick_filter_var.get().lower()

        self.current_view_data = [
            row for row in master_data
            if not filter_text or any(filter_text in str(val).lower() for val in row.values())
        ]

        if self.sort_criteria:
            for col, direction in reversed(self.sort_criteria):
                self.current_view_data.sort(
                    key=lambda item, c=col: str(item.get(c, '')).lower(),
                    reverse=(direction == 'desc')
                )
        self.virtual_view_top_index = 0
        self.selected_data_index = None
        self._update_virtual_table_view()

    def _update_virtual_table_view(self):
        self._repopulate_virtual_table()
        self._update_virtual_scrollbar()
        self._update_nav_controls()
        self._update_header_sort_indicators()

    def _update_header_sort_indicators(self):
        for col in self.table_info['columns']:
            self.table_treeview.heading(col, text=col)
        for i, (col, direction) in enumerate(self.sort_criteria):
            arrow = " ▲" if direction == 'asc' else " ▼"
            self.table_treeview.heading(col, text=f"{col}{arrow}{i + 1}")

    def _update_nav_controls(self):
        total_rows = len(self.current_view_data)
        if total_rows == 0:
            self.nav_status_var.set("No matching rows")
            state = "disabled"
        else:
            current_selection = self.selected_data_index + 1 if self.selected_data_index is not None else "-"
            self.nav_status_var.set(f"Selected: {current_selection} / {total_rows}")
            state = "normal"

        self.nav_btn_top.config(state=state)
        self.nav_btn_end.config(state=state)
        self.nav_btn_prev.config(
            state="normal" if self.selected_data_index is not None and self.selected_data_index > 0 else "disabled")
        self.nav_btn_next.config(
            state="normal" if self.selected_data_index is not None and self.selected_data_index < total_rows - 1 else "disabled")

    def _nav_top(self):
        if self.current_view_data: self._jump_to_virtual_index(0)

    def _nav_end(self):
        if self.current_view_data: self._jump_to_virtual_index(len(self.current_view_data) - 1)

    def _nav_prev(self):
        if self.selected_data_index is not None and self.selected_data_index > 0: self._jump_to_virtual_index(
            self.selected_data_index - 1)

    def _nav_next(self):
        if self.selected_data_index is not None and self.selected_data_index < len(
            self.current_view_data) - 1: self._jump_to_virtual_index(self.selected_data_index + 1)

    def handle_goto_row(self):
        total_rows = len(self.current_view_data)
        if total_rows == 0: return
        row_num = simpledialog.askinteger("Go to Row", f"Enter row number (1 - {total_rows}):", parent=self.app.root,
                                          minvalue=1, maxvalue=total_rows)
        if row_num is not None: self._jump_to_virtual_index(row_num - 1)

    def resize_columns(self):
        tree, items_to_check = self.table_treeview, self.table_treeview.get_children()
        if not tree["columns"]: return
        style = ttk.Style()
        font_name = style.lookup("Treeview", "font") or "TkDefaultFont"
        tree_font = tkFont.Font(font=font_name)
        padding = 20
        for col_id in tree["columns"]:
            header_text = tree.heading(col_id, "text").split(" ")[0]
            max_width = tree_font.measure(header_text) + padding
            col_index = tree["columns"].index(col_id)
            for item_id in items_to_check:
                values = tree.item(item_id, "values")
                cell_value = str(values[col_index]) if values and col_index < len(values) else ""
                max_width = max(max_width, tree_font.measure(cell_value) + padding)
            tree.column(col_id, width=min(max(max_width, 50), 500), stretch=False)

    def get_current_table_data(self):
        headers = self.table_info["columns"]
        rows_to_write = [[row.get(h, "") for h in headers] for row in self.current_view_data]
        return headers, rows_to_write



class TransactionalDataChecker(tk.Toplevel):
    """
    A UI to perform transactional integrity checks between two tables.
    """

    def __init__(self, parent, app_instance, potential_tables, table_combobox_map, source_path, file_type,
                 table_data_cache):
        super().__init__(parent)
        self.title("Transactional Data Check")
        self.geometry("1000x700")
        self.transient(parent)
        self.grab_set()

        self.app = app_instance
        self.potential_tables = potential_tables
        self.table_combobox_map = table_combobox_map
        self.table_names = sorted(list(self.table_combobox_map.keys()))
        self.source_path = source_path
        self.file_type = file_type
        self.table_data_cache = table_data_cache

        # --- UI State Variables ---
        self.primary_table_var = tk.StringVar()
        self.primary_key_var = tk.StringVar()
        self.trans_table_var = tk.StringVar()
        self.trans_key_var = tk.StringVar()
        self.check_type_var = tk.StringVar(value="orphans")  # 'orphans' or 'unused'

        # --- Results State ---
        self.current_results_data = []

        self._setup_ui()
        self._populate_initial_dropdowns()

    def _setup_ui(self):
        main_paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Top Config Frame ---
        config_frame = ttk.Labelframe(main_paned, text="1. Configuration", padding=10)
        main_paned.add(config_frame, weight=1)

        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)

        # Primary Table
        ttk.Label(config_frame, text="Primary Table (e.g., Customer Master):").grid(row=0, column=0, padx=5, pady=5,
                                                                                    sticky="w")
        self.primary_table_combo = ttk.Combobox(config_frame, textvariable=self.primary_table_var, state="readonly")
        self.primary_table_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(config_frame, text="Primary Key Column:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.primary_key_combo = ttk.Combobox(config_frame, textvariable=self.primary_key_var, state="disabled")
        self.primary_key_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Transactional Table
        ttk.Label(config_frame, text="Transactional Table (e.g., Invoices):").grid(row=0, column=2, padx=5, pady=5,
                                                                                    sticky="w")
        self.trans_table_combo = ttk.Combobox(config_frame, textvariable=self.trans_table_var, state="readonly")
        self.trans_table_combo.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        ttk.Label(config_frame, text="Foreign Key Column:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.trans_key_combo = ttk.Combobox(config_frame, textvariable=self.trans_key_var, state="disabled")
        self.trans_key_combo.grid(row=1, column=3, padx=5, pady=5, sticky="ew")

        # Bindings
        self.primary_table_combo.bind("<<ComboboxSelected>>", self._on_primary_table_select)
        self.trans_table_combo.bind("<<ComboboxSelected>>", self._on_trans_table_select)

        # Check Type
        check_type_frame = ttk.Labelframe(config_frame, text="2. Check Type", padding=10)
        check_type_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=10, sticky="ew")

        ttk.Radiobutton(check_type_frame, text="Find transactions with no matching primary record (Orphans)",
                        variable=self.check_type_var, value="orphans").pack(anchor="w", pady=2)
        ttk.Radiobutton(check_type_frame, text="Find primary records with no matching transactions (Unused)",
                        variable=self.check_type_var, value="unused").pack(anchor="w", pady=2)

        # Run Button
        run_button = ttk.Button(config_frame, text="Run Check", command=self._run_check, style="Accent.TButton")
        run_button.grid(row=3, column=0, columnspan=4, pady=10)

        # --- Bottom Results Frame ---
        results_frame = ttk.Labelframe(main_paned, text="Results", padding=10)
        main_paned.add(results_frame, weight=3)
        results_frame.rowconfigure(1, weight=1)
        results_frame.columnconfigure(0, weight=1)

        # Results Actions
        action_frame = ttk.Frame(results_frame)
        action_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        ttk.Button(action_frame, text="Export Results as CSV...", command=self._export_results).pack(side=tk.LEFT)
        ttk.Button(action_frame, text="Resize Columns", command=self._resize_columns).pack(side=tk.LEFT, padx=5)
        self.results_status_label = ttk.Label(action_frame, text="No check performed yet.")
        self.results_status_label.pack(side=tk.RIGHT, padx=5)

        # Results Grid
        self.results_tree = ttk.Treeview(results_frame, show='headings', selectmode='extended')
        results_vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        results_hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=results_vsb.set, xscrollcommand=results_hsb.set)
        results_hsb.grid(row=2, column=0, sticky='ew')
        results_vsb.grid(row=1, column=1, sticky='ns')
        self.results_tree.grid(row=1, column=0, sticky='nsew')

    def _populate_initial_dropdowns(self):
        self.primary_table_combo['values'] = self.table_names
        self.trans_table_combo['values'] = self.table_names

    def _on_primary_table_select(self, event=None):
        self.primary_key_var.set('')
        table_name = self.primary_table_var.get()
        if not table_name:
            self.primary_key_combo.config(state="disabled")
            self.primary_key_combo['values'] = []
            return

        key = self.table_combobox_map.get(table_name)
        cols = self.potential_tables[key]['columns'] if key else []
        self.primary_key_combo['values'] = cols
        self.primary_key_combo.config(state="readonly" if cols else "disabled")

    def _on_trans_table_select(self, event=None):
        self.trans_key_var.set('')
        table_name = self.trans_table_var.get()
        if not table_name:
            self.trans_key_combo.config(state="disabled")
            self.trans_key_combo['values'] = []
            return

        key = self.table_combobox_map.get(table_name)
        cols = self.potential_tables[key]['columns'] if key else []
        self.trans_key_combo['values'] = cols
        self.trans_key_combo.config(state="readonly" if cols else "disabled")

    def _run_check(self):
        primary_table = self.primary_table_var.get()
        primary_key_col = self.primary_key_var.get()
        trans_table = self.trans_table_var.get()
        trans_key_col = self.trans_key_var.get()
        check_type = self.check_type_var.get()

        if not all([primary_table, primary_key_col, trans_table, trans_key_col]):
            messagebox.showerror("Incomplete Selection",
                                 "Please select both tables and their respective key columns.", parent=self)
            return

        try:
            primary_rows = self._get_rows_from_source(primary_table)
            trans_rows = self._get_rows_from_source(trans_table)

            results = []
            result_columns = []

            if check_type == 'orphans':
                primary_keys = {self._get_cell_value(row, primary_key_col) for row in primary_rows}
                result_columns = self._get_all_columns(trans_table)
                for row in trans_rows:
                    f_key_val = self._get_cell_value(row, trans_key_col)
                    if f_key_val not in primary_keys:
                        results.append(row)
            else:  # 'unused'
                trans_keys = {self._get_cell_value(row, trans_key_col) for row in trans_rows}
                result_columns = self._get_all_columns(primary_table)
                for row in primary_rows:
                    p_key_val = self._get_cell_value(row, primary_key_col)
                    if p_key_val not in trans_keys:
                        results.append(row)

            self.current_results_data = results
            self._display_results(result_columns)

        except Exception as e:
            messagebox.showerror("Execution Error", f"An error occurred during the check:\n{e}", parent=self)
            traceback.print_exc()

    def _display_results(self, columns):
        self.results_tree.delete(*self.results_tree.get_children())
        self.results_tree["columns"] = columns

        for col in columns:
            self.results_tree.heading(col, text=col, anchor='w')
            self.results_tree.column(col, width=120, anchor='w', stretch=True)

        for result_row in self.current_results_data:
            values = [result_row.get(col, "") for col in columns]
            self.results_tree.insert("", "end", values=values)

        count = len(self.current_results_data)
        if count == 0:
            self.results_status_label.config(text="No issues found.")
        else:
            self.results_status_label.config(text=f"Found {count} records.")

    def _get_rows_from_source(self, table_name):
        key = self.table_combobox_map[table_name]
        if key in self.table_data_cache:
            return self.table_data_cache.get(key, [])

        if self.file_type == 'xml':
            table_info = self.potential_tables.get(key)
            if not table_info:
                return []

            parent_element = table_info["parent_element"]
            row_tag = table_info["row_tag"]
            row_elements = parent_element.findall(row_tag)
            parsed_data = []
            for i, row_element in enumerate(row_elements):
                row_data = {"_element": row_element, "_original_index": i}
                for col_name in table_info["columns"]:
                    col_data_element = row_element.find(col_name)
                    cell_text = col_data_element.text.strip() if col_data_element is not None and col_data_element.text else ""
                    row_data[col_name] = cell_text
                parsed_data.append(row_data)
            self.table_data_cache[key] = parsed_data
            return parsed_data

        return []

    def _get_cell_value(self, row, column):
        return row.get(column, "").strip()

    def _get_all_columns(self, table_name):
        key = self.table_combobox_map[table_name]
        return self.potential_tables[key]['columns']

    def _export_results(self):
        if not self.results_tree.get_children():
            messagebox.showwarning("Export Error", "There are no results to export.", parent=self)
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                                parent=self)
        if not filepath:
            return
        try:
            delimiter = self.app.csv_delimiter if hasattr(self.app, 'csv_delimiter') else ','
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=delimiter)
                writer.writerow(self.results_tree["columns"])
                for item_id in self.results_tree.get_children():
                    writer.writerow(self.results_tree.item(item_id, "values"))
            messagebox.showinfo("Success", "Results exported successfully.", parent=self)
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred during export:\n{e}", parent=self)

    def _resize_columns(self):
        tree = self.results_tree
        if not tree["columns"]:
            return
        try:
            items_to_sample = tree.get_children()[:VIRTUAL_TABLE_ROW_COUNT]

            style = ttk.Style()
            font_name = style.lookup("Treeview", "font") or "TkDefaultFont"
            tree_font = tkFont.Font(font=font_name)
            padding = 20

            for col_id in tree["columns"]:
                header_text = tree.heading(col_id, "text")
                max_width = tree_font.measure(header_text) + padding

                col_index = tree["columns"].index(col_id)
                for item_id in items_to_sample:
                    values = tree.item(item_id, "values")
                    try:
                        cell_value = str(values[col_index]) if values and col_index < len(values) else ""
                        cell_width = tree_font.measure(cell_value) + padding
                        if cell_width > max_width:
                            max_width = cell_width
                    except (ValueError, IndexError):
                        continue

                max_width = min(max(max_width, 50), 500)
                tree.column(col_id, width=max_width, stretch=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resize columns: {e}", parent=self)



class QueryDesigner(tk.Toplevel):
    """
    An advanced query designer with a visual builder, simple text parser, and SQL view.
    """

    def __init__(self, parent, app_instance, potential_tables, table_combobox_map, source_path, file_type,
                 table_data_cache):
        super().__init__(parent)
        self.title("Advanced Query Designer")
        self.geometry("1300x900")
        self.transient(parent)
        self.grab_set()

        self.app = app_instance
        self.potential_tables = potential_tables
        self.table_combobox_map = table_combobox_map
        self.table_names = sorted(list(self.table_combobox_map.keys()))
        self.table_names_with_blank = [""] + self.table_names
        self.source_path = source_path
        self.file_type = file_type
        self.table_data_cache = table_data_cache

        # --- Visual Designer State ---
        self.table1_var = tk.StringVar()
        self.table2_var = tk.StringVar()
        self.condition_type_var = tk.StringVar(value="filter")
        self.field1_var = tk.StringVar()
        self.field2_var = tk.StringVar()
        self.filter_field_var = tk.StringVar()
        self.filter_op_var = tk.StringVar(value="CONTAINS")
        self.filter_value_var = tk.StringVar()
        self.query_type_var = tk.StringVar(value="INNER")
        self.visual_conditions = []
        self.aggregate_func_var = tk.StringVar(value="None")
        self.grouped_by_fields = []

        # --- Simple Query State ---
        self.simple_query_table_var = tk.StringVar()
        self.simple_query_intellisense_popup = None

        # --- Shared State ---
        self.limit_enabled_var = tk.BooleanVar(value=False)
        self.limit_value_var = tk.IntVar(value=100)
        self.manual_edit_mode = tk.BooleanVar(value=False)

        # --- Results State ---
        self.current_results_data = []
        self.results_sort_col = None
        self.results_sort_asc = True

        # --- Intellisense & Help ---
        self.intellisense_popup = None
        self.help_window = None

        self._setup_ui()
        self._populate_initial_dropdowns()
        self._on_table_select()

        self.bind("<F1>", self._show_help)
        self.bind("<Control-w>", lambda e: self._resize_query_results_columns())
        self.bind("<Control-e>", lambda e: self._export_results())
        self.config_notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

    def _show_help(self, event=None):
        if self.help_window is None or not self.help_window.winfo_exists():
            self.help_window = HelpWindow(self)
        self.help_window.focus()

    def _setup_ui(self):
        self.menubar = tk.Menu(self)

        self.querymenu = tk.Menu(self.menubar, tearoff=0)
        self.querymenu.add_command(label="Load Query...", command=self._load_config)
        self.querymenu.add_command(label="Save Query...", command=self._save_config)
        self.querymenu.add_separator()
        self.querymenu.add_command(label="Fix Column Widths", command=self._resize_query_results_columns,
                                   accelerator="Ctrl+W")
        self.querymenu.add_command(label="Export Results as CSV...", command=self._export_results, accelerator="Ctrl+E")
        self.querymenu.add_separator()
        self.querymenu.add_command(label="Exit", command=self.destroy)
        self.menubar.add_cascade(label="Query", menu=self.querymenu)

        self.helpmenu = tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="Help Topics", command=self._show_help, accelerator="F1")
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)

        self.config(menu=self.menubar)

        main_paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.config_notebook = ttk.Notebook(main_paned)
        main_paned.add(self.config_notebook, weight=2)

        visual_designer_frame = ttk.Frame(self.config_notebook, padding=10)
        self.config_notebook.add(visual_designer_frame, text="Visual Designer")
        self._create_visual_designer_widgets(visual_designer_frame)

        simple_query_frame = ttk.Frame(self.config_notebook, padding=10)
        self.config_notebook.add(simple_query_frame, text="Simple Query")
        self._create_simple_query_widgets(simple_query_frame)

        sql_view_frame = ttk.Frame(self.config_notebook, padding=10)
        self.config_notebook.add(sql_view_frame, text="SQL View")
        self._create_sql_view_widgets(sql_view_frame)

        results_pane = ttk.Frame(main_paned)
        main_paned.add(results_pane, weight=3)
        self._create_results_widgets(results_pane)

        self.table1_combo.bind("<<ComboboxSelected>>", self._on_table_select)
        self.table2_combo.bind("<<ComboboxSelected>>", self._on_table_select)
        self.results_tree.bind("<Button-1>", self._on_results_click)

        for var in [self.query_type_var, self.limit_enabled_var, self.limit_value_var]:
            var.trace_add("write", lambda *args: self._update_query_view())

    def _parse_simple_query(self, text, valid_columns):
        text = re.sub(r'\s+', ' ', text).strip()
        text = re.sub(r'\b(all fields|all)\b', '\*', text, flags=re.IGNORECASE)

        # Split the query into the part before 'where' and the part after.
        parts = re.split(r'\s+where\s+', text, maxsplit=1, flags=re.IGNORECASE)
        show_part = parts[0].strip()
        conditions_part = parts[1] if len(parts) > 1 else None

        # First, ensure the query actually starts with 'show'.
        if not show_part.lower().startswith('show'):
            return {'success': False, 'error': "Invalid query. Must start with 'show'."}

        # The fields are whatever is between 'show' and 'where'.
        fields_str = show_part[4:].strip()

        # If the fields string is empty, it implies all fields ('*').
        if not fields_str:
            fields_str = '*'

        # Validate the specified fields.
        fields = valid_columns if fields_str == '*' else [f.strip() for f in fields_str.split(',')]
        for field in fields:
            if field != '*' and field not in valid_columns:
                start_index = text.find(field)
                return {'success': False, 'error': f"Field '{field}' not found.",
                        'span': (start_index, start_index + len(field))}

        conditions = []
        if conditions_part:
            op_pattern = r"\s*(is not|is|contains|not contains|starts with|ends with|>=|<=|>|<|!=|=)\s*"
            condition_parts = re.split(r'\s+and\s+', conditions_part, flags=re.IGNORECASE)

            for part in condition_parts:
                part = part.strip()
                op_match = re.search(op_pattern, part, re.IGNORECASE)
                if not op_match:
                    date_func_match = re.match(r"(year|month|day)\s+of\s+([\w\.]+)", part, re.I)
                    if not date_func_match:
                        start = text.find(part)
                        return {'success': False, 'error': f"Invalid condition format: '{part}'",
                                'span': (start, start + len(part))}

                    func, field_str = date_func_match.groups()
                    op_part = part[date_func_match.end():].strip()
                    op_match2 = re.search(r"^(>=|<=|>|<|!=|=)", op_part)
                    if not op_match2:
                        start = text.find(part)
                        return {'success': False, 'error': "Invalid operator for date function.",
                                'span': (start, start + len(part))}

                    op = op_match2.group(1)
                    value = op_part[op_match2.end():].strip()
                    field = (func.lower(), field_str.strip())
                else:
                    field = part[:op_match.start()].strip()
                    op = op_match.group(1).strip().lower()
                    value = part[op_match.end():].strip()

                clean_field = field[1] if isinstance(field, tuple) else field
                if clean_field not in valid_columns:
                    start = text.find(clean_field)
                    return {'success': False, 'error': f"Field '{clean_field}' not found.",
                            'span': (start, start + len(clean_field))}

                if value.startswith("'") and value.endswith("'"):
                    value = value[1:-1]
                elif value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]

                op_map = {'is': '=', 'is not': '!=', 'not contains': 'NOT CONTAINS'}
                final_op = op_map.get(op, op).upper()
                conditions.append({'field': field, 'op': final_op, 'value': value})

        return {'success': True, 'fields': fields, 'conditions': conditions}

    def _row_matches_filters(self, row1_el, row2_el, condition_node):
        if not condition_node:
            return True

        is_group = 'group' in condition_node

        if not is_group:
            condition = condition_node
            table_alias = condition.get('table', 'T1')
            field, op, value = condition['field'], condition['op'], condition['value']

            target_row = row1_el if table_alias == "T1" else row2_el
            if target_row is None:
                return False

            cell_value = self._get_cell_value(target_row, field)
            cell_compare, value_compare = cell_value.lower(), value.lower()

            match = False
            if op == "CONTAINS":
                match = value_compare in cell_compare
            elif op == "NOT CONTAINS":
                match = value_compare not in cell_compare
            elif op == "STARTS WITH":
                match = cell_compare.startswith(value_compare)
            elif op == "ENDS WITH":
                match = cell_compare.endswith(value_compare)
            elif op == "=":
                match = cell_compare == value_compare
            elif op == "!=":
                match = cell_compare != value_compare
            else:
                try:
                    cell_num, val_num = float(cell_value), float(value)
                    if op == ">":
                        match = cell_num > val_num
                    elif op == "<":
                        match = cell_num < val_num
                    elif op == ">=":
                        match = cell_num >= val_num
                    elif op == "<=":
                        match = cell_num <= val_num
                except (ValueError, TypeError):
                    pass
            return match
        else:
            group_type = condition_node.get('group', 'AND')
            conditions = condition_node.get('conditions', [])

            if group_type == 'NOT':
                return not self._row_matches_filters(row1_el, row2_el, conditions[0])

            for condition in conditions:
                match = self._row_matches_filters(row1_el, row2_el, condition)

                if group_type == 'AND' and not match: return False
                if group_type == 'OR' and match: return True

            return True if group_type == 'AND' else False

    def _create_visual_designer_widgets(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        # --- TOP ROW: Tables and Conditions ---
        top_row_frame = ttk.Frame(parent)
        top_row_frame.grid(row=0, column=0, columnspan=2, sticky="ew")
        top_row_frame.grid_columnconfigure(1, weight=1)

        table_select_frame = ttk.Labelframe(top_row_frame, text="1. Select Tables", padding=10)
        table_select_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ns")
        ttk.Label(table_select_frame, text="Left Table (T1):").pack(anchor="w")
        self.table1_combo = ttk.Combobox(table_select_frame, textvariable=self.table1_var, state="readonly", width=30)
        self.table1_combo.pack(pady=2, fill=tk.X, expand=True)
        ttk.Label(table_select_frame, text="Right Table (T2):").pack(anchor="w", pady=(5, 0))
        self.table2_combo = ttk.Combobox(table_select_frame, textvariable=self.table2_var, state="readonly", width=30)
        self.table2_combo.pack(pady=2, fill=tk.X, expand=True)

        self.conditions_frame = ttk.Labelframe(top_row_frame, text="2. Define Conditions (WHERE / ON)", padding=10)
        self.conditions_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        self.conditions_frame.grid_columnconfigure(0, weight=1)
        self.conditions_frame.grid_rowconfigure(3, weight=1)

        condition_type_frame = ttk.Frame(self.conditions_frame)
        condition_type_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))
        self.join_radio = ttk.Radiobutton(condition_type_frame, text="Join (ON)", variable=self.condition_type_var,
                                          value="join", command=self._on_condition_type_change, state="disabled")
        self.join_radio.pack(side=tk.LEFT, padx=(0, 10))
        self.filter_radio = ttk.Radiobutton(condition_type_frame, text="Filter (WHERE)",
                                            variable=self.condition_type_var, value="filter",
                                            command=self._on_condition_type_change)
        self.filter_radio.pack(side=tk.LEFT)

        condition_input_frame = ttk.Frame(self.conditions_frame)
        condition_input_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

        self.join_controls_frame = ttk.Frame(condition_input_frame)
        self.field1_combo = ttk.Combobox(self.join_controls_frame, textvariable=self.field1_var, state="disabled",
                                         width=15);
        self.field1_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Label(self.join_controls_frame, text="=").pack(side=tk.LEFT, padx=5)
        self.field2_combo = ttk.Combobox(self.join_controls_frame, textvariable=self.field2_var, state="disabled",
                                         width=15);
        self.field2_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.filter_controls_frame = ttk.Frame(condition_input_frame)
        self.filter_field_combo = ttk.Combobox(self.filter_controls_frame, textvariable=self.filter_field_var,
                                               state="disabled", width=15);
        self.filter_field_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        op_values = ["CONTAINS", "NOT CONTAINS", "=", "!=", ">", "<", ">=", "<=", "STARTS WITH", "ENDS WITH"];
        self.filter_op_combo = ttk.Combobox(self.filter_controls_frame, textvariable=self.filter_op_var,
                                            values=op_values, state="readonly", width=12);
        self.filter_op_combo.pack(side=tk.LEFT, padx=5)
        self.filter_value_entry = ttk.Entry(self.filter_controls_frame, textvariable=self.filter_value_var, width=15);
        self.filter_value_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        condition_actions_frame = ttk.Frame(self.conditions_frame)
        condition_actions_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Button(condition_actions_frame, text="Add", command=self._add_condition, style="Accent.TButton").pack(
            side=tk.LEFT, padx=2)
        ttk.Button(condition_actions_frame, text="(", command=lambda: self._add_logical_operator("(")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(condition_actions_frame, text=")", command=lambda: self._add_logical_operator(")")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(condition_actions_frame, text="AND", command=lambda: self._add_logical_operator("AND")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(condition_actions_frame, text="OR", command=lambda: self._add_logical_operator("OR")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(condition_actions_frame, text="NOT", command=lambda: self._add_logical_operator("NOT")).pack(
            side=tk.LEFT, padx=2)

        self.conditions_listbox = tk.Listbox(self.conditions_frame, height=4);
        self.conditions_listbox.grid(row=3, column=0, columnspan=2, sticky="nsew")
        self.condition_status_label = ttk.Label(self.conditions_frame, text="Valid", foreground="green");
        self.condition_status_label.grid(row=4, column=0, sticky="w", pady=(2, 0))
        ttk.Button(self.conditions_frame, text="Remove", command=self._remove_condition_from_list).grid(row=4, column=1,
                                                                                                        sticky="e",
                                                                                                        pady=(2, 0))

        # --- MIDDLE ROW: Group By ---
        group_by_frame = ttk.Labelframe(parent, text="3. Group By (Optional)", padding=10)
        group_by_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        group_by_frame.grid_columnconfigure(0, weight=1);
        group_by_frame.grid_columnconfigure(2, weight=1)
        group_by_frame.grid_rowconfigure(1, weight=1)

        ttk.Label(group_by_frame, text="Available Fields").grid(row=0, column=0, sticky="w")
        group_avail_frame = ttk.Frame(group_by_frame)
        group_avail_frame.grid(row=1, column=0, sticky="nsew")
        self.group_available_lb = tk.Listbox(group_avail_frame, selectmode=tk.EXTENDED, height=4);
        group_avail_vsb = ttk.Scrollbar(group_avail_frame, orient=tk.VERTICAL, command=self.group_available_lb.yview)
        self.group_available_lb.config(yscrollcommand=group_avail_vsb.set)
        group_avail_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.group_available_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        group_shuttle_buttons = ttk.Frame(group_by_frame);
        group_shuttle_buttons.grid(row=1, column=1, padx=5)
        ttk.Button(group_shuttle_buttons, text=">", width=3, command=self._add_group_by_field).pack(pady=2)
        ttk.Button(group_shuttle_buttons, text="<", width=3, command=self._remove_group_by_field).pack(pady=2)

        ttk.Label(group_by_frame, text="Group By Fields").grid(row=0, column=2, sticky="w")
        group_sel_frame = ttk.Frame(group_by_frame)
        group_sel_frame.grid(row=1, column=2, sticky="nsew")
        self.grouped_by_lb = tk.Listbox(group_sel_frame, selectmode=tk.EXTENDED, height=4);
        group_sel_vsb = ttk.Scrollbar(group_sel_frame, orient=tk.VERTICAL, command=self.grouped_by_lb.yview)
        self.grouped_by_lb.config(yscrollcommand=group_sel_vsb.set)
        group_sel_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.grouped_by_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- BOTTOM ROW: Output Design ---
        output_frame = ttk.Labelframe(parent, text="4. Design Report Output", padding=10)
        output_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        output_frame.grid_columnconfigure(0, weight=1);
        output_frame.grid_columnconfigure(2, weight=1)
        output_frame.grid_rowconfigure(1, weight=1)

        ttk.Label(output_frame, text="Available Fields").grid(row=0, column=0, sticky="w")
        out_avail_frame = ttk.Frame(output_frame)
        out_avail_frame.grid(row=1, column=0, sticky="nsew")
        self.available_fields_lb = tk.Listbox(out_avail_frame, selectmode=tk.EXTENDED)
        out_avail_vsb = ttk.Scrollbar(out_avail_frame, orient=tk.VERTICAL, command=self.available_fields_lb.yview)
        self.available_fields_lb.config(yscrollcommand=out_avail_vsb.set)
        out_avail_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.available_fields_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        agg_frame = ttk.Frame(output_frame);
        agg_frame.grid(row=1, column=1, padx=5, sticky="n")
        agg_values = ["None", "COUNT", "SUM", "AVG", "MIN", "MAX"];
        self.agg_combo = ttk.Combobox(agg_frame, textvariable=self.aggregate_func_var, values=agg_values,
                                      state="readonly", width=8);
        self.agg_combo.pack(pady=2)
        ttk.Button(agg_frame, text="Add Field >", command=self._add_output_field).pack(pady=2)

        ttk.Label(output_frame, text="Selected Fields (in order)").grid(row=0, column=2, sticky="w")
        out_sel_frame = ttk.Frame(output_frame)
        out_sel_frame.grid(row=1, column=2, sticky="nsew")
        self.selected_fields_lb = tk.Listbox(out_sel_frame, selectmode=tk.EXTENDED)
        out_sel_vsb = ttk.Scrollbar(out_sel_frame, orient=tk.VERTICAL, command=self.selected_fields_lb.yview)
        self.selected_fields_lb.config(yscrollcommand=out_sel_vsb.set)
        out_sel_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.selected_fields_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        output_shuttle_frame = ttk.Frame(output_frame);
        output_shuttle_frame.grid(row=1, column=3, padx=5, sticky="n")
        ttk.Button(output_shuttle_frame, text="< Remove", command=self._remove_output_field).pack(pady=2)
        ttk.Button(output_shuttle_frame, text="Up", command=lambda: self._move_output_field(-1)).pack(pady=2)
        ttk.Button(output_shuttle_frame, text="Down", command=lambda: self._move_output_field(1)).pack(pady=2)

    def _add_group_by_field(self):
        selected_indices = self.group_available_lb.curselection()
        for i in reversed(selected_indices):
            field = self.group_available_lb.get(i)
            self.grouped_by_lb.insert(tk.END, field)
            self.group_available_lb.delete(i)
        self._update_query_view()

    def _remove_group_by_field(self):
        selected_indices = self.grouped_by_lb.curselection()
        for i in reversed(selected_indices):
            field = self.grouped_by_lb.get(i)
            self.group_available_lb.insert(tk.END, field)
            self.grouped_by_lb.delete(i)
        self._update_query_view()

    def _validate_and_highlight_conditions(self):
        self.condition_status_label.config(text="Validating...", foreground="orange")
        is_valid = True
        error_msg = "Valid"

        for i in range(self.conditions_listbox.size()):
            self.conditions_listbox.itemconfigure(i, foreground="black")

        paren_balance = 0
        last_item_type = 'op_or_start'
        for i, item in enumerate(self.visual_conditions):
            item_type = item['type']
            value = item.get('value')

            if item_type == 'op' and value in ['AND', 'OR', 'NOT']:
                self.conditions_listbox.itemconfigure(i, foreground="blue")

            if value == '(':
                paren_balance += 1
            elif value == ')':
                paren_balance -= 1
            if paren_balance < 0:
                is_valid, error_msg = False, "Error: Mismatched ')'"
                self.conditions_listbox.itemconfigure(i, foreground="red")
                break

            if (last_item_type == 'cond' and item_type == 'cond') or \
                    (last_item_type == 'op_or_start' and value in ['AND', 'OR']):
                is_valid, error_msg = False, "Error: Invalid operator sequence"
                self.conditions_listbox.itemconfigure(i, foreground="red")
                break

            last_item_type = 'cond' if item_type == 'cond' else 'op'
            if value == '(':
                last_item_type = 'op_or_start'

        if is_valid and paren_balance != 0:
            is_valid, error_msg = False, "Error: Mismatched '('"
            for i in range(self.conditions_listbox.size() - 1, -1, -1):
                if self.conditions_listbox.get(i) == '(':
                    self.conditions_listbox.itemconfigure(i, foreground="red")
                    break

        if is_valid:
            self.condition_status_label.config(text="Valid", foreground="green")
            self.run_designer_button.config(state="normal")
        else:
            self.condition_status_label.config(text=error_msg, foreground="red")
            self.run_designer_button.config(state="disabled")

        self._update_query_view()

    def _on_condition_type_change(self):
        cond_type = self.condition_type_var.get()
        if cond_type == "join":
            self.join_controls_frame.pack(fill=tk.X, expand=True)
            self.filter_controls_frame.pack_forget()
        else:
            self.filter_controls_frame.pack(fill=tk.X, expand=True)
            self.join_controls_frame.pack_forget()

    def _add_condition(self):
        cond_type = self.condition_type_var.get()

        if cond_type == "join":
            f1, f2 = self.field1_var.get(), self.field2_var.get()
            if not (f1 and f2):
                messagebox.showwarning("Incomplete Join", "Please select a field from both tables.", parent=self)
                return
            condition_data = {'type': 'join', 't1_field': f1, 't2_field': f2}
            display_text = f"T1.{f1} = T2.{f2}"
        else:
            field, op, value = self.filter_field_var.get(), self.filter_op_var.get(), self.filter_value_var.get()
            if not field or not op:
                messagebox.showwarning("Incomplete Filter", "Please select a field and an operator.", parent=self)
                return

            table_alias = "T1"
            clean_field = field
            if field.startswith("T1: ") or field.startswith("T2: "):
                parts = field.split(": ", 1)
                table_alias = parts[0]
                clean_field = parts[1]

            condition_data = {'type': 'filter', 'table': table_alias, 'field': clean_field, 'op': op, 'value': value}
            display_text = f"{table_alias}.{clean_field} {op} '{value}'"

        self.visual_conditions.append({'type': 'cond', 'data': condition_data})
        self.conditions_listbox.insert(tk.END, display_text)
        self._validate_and_highlight_conditions()

    def _add_logical_operator(self, operator):
        self.visual_conditions.append({'type': 'op', 'value': operator})

        display_text = operator
        if operator not in ['(', ')', 'NOT']:
            display_text = f"  {operator}"

        self.conditions_listbox.insert(tk.END, display_text)

        self.field1_var.set('')
        self.field2_var.set('')
        self.filter_field_var.set('')
        self.filter_value_var.set('')

        self._validate_and_highlight_conditions()

    def _remove_condition_from_list(self):
        selected_indices = self.conditions_listbox.curselection()
        if not selected_indices: return

        for i in sorted(selected_indices, reverse=True):
            self.conditions_listbox.delete(i)
            del self.visual_conditions[i]

        self._validate_and_highlight_conditions()

    def _on_table_select(self, event=None):
        t1_name, t2_name = self.table1_var.get(), self.table2_var.get()

        self.conditions_listbox.delete(0, tk.END)
        self.visual_conditions.clear()
        self._update_available_fields()
        self.group_available_lb.delete(0, tk.END)
        self.grouped_by_lb.delete(0, tk.END)

        all_fields = self.available_fields_lb.get(0, tk.END)
        for field in all_fields:
            self.group_available_lb.insert(tk.END, field)

        is_single_table = t1_name and not t2_name
        is_join = t1_name and t2_name

        if is_single_table:
            self.condition_type_var.set("filter")
            self.join_radio.config(state="disabled")
            self.filter_radio.config(state="normal")

            key1 = self.table_combobox_map.get(t1_name)
            cols = self.potential_tables[key1]['columns'] if key1 else []
            self.filter_field_combo['values'] = cols
            self.filter_field_combo.config(state="readonly" if cols else "disabled")
            self.join_type_rb_inner.config(state="disabled")
            self.join_type_rb_anti.config(state="disabled")
        elif is_join:
            self.condition_type_var.set("join")
            self.join_radio.config(state="normal")
            self.filter_radio.config(state="normal")

            key1 = self.table_combobox_map.get(t1_name)
            key2 = self.table_combobox_map.get(t2_name)
            cols1 = self.potential_tables[key1]['columns'] if key1 else []
            cols2 = self.potential_tables[key2]['columns'] if key2 else []
            self.field1_combo['values'] = cols1
            self.field1_combo.config(state="readonly" if key1 else "disabled")
            self.field2_combo['values'] = cols2
            self.field2_combo.config(state="readonly" if key2 else "disabled")

            filter_cols = [f"T1: {c}" for c in cols1] + [f"T2: {c}" for c in cols2]
            self.filter_field_combo['values'] = filter_cols
            self.filter_field_combo.config(state="readonly" if filter_cols else "disabled")

            self.join_type_rb_inner.config(state="normal")
            self.join_type_rb_anti.config(state="normal")
        else:
            self.join_radio.config(state="disabled")
            self.filter_radio.config(state="disabled")
            self.join_type_rb_inner.config(state="disabled")
            self.join_type_rb_anti.config(state="disabled")

        self._on_condition_type_change()
        self._validate_and_highlight_conditions()

    def _update_query_view(self):
        t1, t2 = self.table1_var.get(), self.table2_var.get()
        select_clause = ",\n  ".join(self.selected_fields_lb.get(0, tk.END)) or "[Select Output Fields]"
        limit_str = f"\nLIMIT {self.limit_value_var.get()}" if self.limit_enabled_var.get() else ""
        query_str = ""

        join_conditions = []
        filter_clause_parts = []

        for item in self.visual_conditions:
            if item['type'] == 'op':
                filter_clause_parts.append(item['value'])
            elif item['type'] == 'cond':
                data = item['data']
                if data['type'] == 'join':
                    join_conditions.append(f"T1.{data['t1_field']} = T2.{data['t2_field']}")
                else:
                    filter_clause_parts.append(f"{data['table']}.{data['field']} {data['op']} '{data['value']}'")

        if t1 and not t2:
            where_clause = " ".join(filter_clause_parts) or "[Define Filter Conditions]"
            query_str = f"SELECT\n  {select_clause}\nFROM\n  '{t1}' AS T1\nWHERE\n  {where_clause}{limit_str};"
        elif t1 and t2:
            join_type = self.query_type_var.get().replace("ANTI", "ANTI-JOIN")
            on_clause = "\n    AND ".join(join_conditions) if join_conditions else "[Define Join Conditions]"
            where_clause = " ".join(filter_clause_parts)
            where_str = f"\nWHERE\n  {where_clause}" if where_clause else ""
            query_str = f"SELECT\n  {select_clause}\nFROM\n  '{t1}' AS T1\n{join_type}\n  '{t2}' AS T2\n  ON {on_clause}{where_str}{limit_str};"
        else:
            query_str = "Please select at least one table to begin."

        is_editing = self.manual_edit_mode.get()
        if not is_editing:
            self.query_view_text.config(state="normal")
            self.query_view_text.delete("1.0", tk.END)
            self.query_view_text.insert("1.0", query_str)
            self.query_view_text.config(state="disabled")

    def _get_conditions_from_list(self):
        tokens = self.visual_conditions
        if not tokens:
            return None

        output_queue = []
        operator_stack = []
        precedence = {'OR': 1, 'AND': 2, 'NOT': 3}

        for token in tokens:
            if token['type'] == 'cond':
                output_queue.append(token['data'])
            elif token['type'] == 'op':
                op = token['value']
                if op == '(':
                    operator_stack.append(op)
                elif op == ')':
                    while operator_stack and operator_stack[-1] != '(':
                        output_queue.append(operator_stack.pop())
                    if not operator_stack or operator_stack[-1] != '(':
                        raise ValueError("Mismatched parentheses")
                    operator_stack.pop()
                else:
                    while (operator_stack and operator_stack[-1] != '(' and
                           precedence.get(operator_stack[-1], 0) >= precedence.get(op, 0)):
                        output_queue.append(operator_stack.pop())
                    operator_stack.append(op)

        while operator_stack:
            op = operator_stack.pop()
            if op == '(':
                raise ValueError("Mismatched parentheses")
            output_queue.append(op)

        if not output_queue: return None

        eval_stack = []
        for token in output_queue:
            if isinstance(token, dict):
                eval_stack.append(token)
            else:
                if token in ['AND', 'OR']:
                    if len(eval_stack) < 2: raise ValueError("Invalid syntax for AND/OR")
                    right = eval_stack.pop()
                    left = eval_stack.pop()
                    eval_stack.append({'group': token, 'conditions': [left, right]})
                elif token == 'NOT':
                    if len(eval_stack) < 1: raise ValueError("Invalid syntax for NOT")
                    operand = eval_stack.pop()
                    eval_stack.append({'group': 'NOT', 'conditions': [operand]})

        if len(eval_stack) != 1:
            return {'group': 'AND', 'conditions': eval_stack}

        return eval_stack[0]

    def _run_filter_query(self, config=None):
        try:
            t1_name = self.table1_var.get()
            condition_tree = self._get_conditions_from_list()

            rows1 = self._get_rows_from_source(t1_name)
            cols1 = self._get_all_columns(t1_name)
            base_results = []

            for row1 in rows1:
                if self._row_matches_filters(row1, None, condition_tree):
                    result_row = {f"T1: {col}": self._get_cell_value(row1, col) for col in cols1}
                    base_results.append(result_row)

            final_results = self._run_query_engine(base_results)
            self.current_results_data = final_results
            self._display_results_grid(self.selected_fields_lb.get(0, tk.END))
        except ValueError as e:
            messagebox.showerror("Query Error", f"Invalid condition logic: {e}", parent=self)
        except Exception as ex:
            messagebox.showerror("Execution Error", f"An error occurred: {ex}", parent=self)
            traceback.print_exc()

    def _update_results_header_style(self):
        for col in self.results_tree["columns"]:
            text = col
            if col == self.results_sort_col:
                text += " ▲" if self.results_sort_asc else " ▼"
            self.results_tree.heading(col, text=text)

    def _run_join_query(self, config=None):
        try:
            join_conditions = []
            filter_conditions_tree = None

            if config:
                t1_name = config["table1"]
                t2_name = config["table2"]
                join_conditions = config["join_conditions"]

                filter_conditions_tree = {'group': 'AND', 'conditions': [
                    {'type': 'filter', 'table': f[0], 'field': f[1], 'op': o, 'value': v} for f, o, v in
                    config.get("filter_conditions", [])
                ]}
                query_type = config["query_type"]
                output_fields = config["output_fields"]
                limit = config.get("limit_value") if config.get("limit_enabled") else -1
            else:
                t1_name, t2_name = self.table1_var.get(), self.table2_var.get()

                join_cond_data = [item['data'] for item in self.visual_conditions if
                                  item['type'] == 'cond' and item['data']['type'] == 'join']
                join_conditions = [(jc['t1_field'], jc['t2_field']) for jc in join_cond_data]

                if not join_conditions:
                    messagebox.showerror("Error", "Please add at least one join condition.", parent=self)
                    return

                filter_items = [item for item in self.visual_conditions if
                                not (item['type'] == 'cond' and item['data']['type'] == 'join')]
                original_items = self.visual_conditions
                self.visual_conditions = filter_items
                filter_conditions_tree = self._get_conditions_from_list()
                self.visual_conditions = original_items

                query_type = self.query_type_var.get()
                output_fields = self.selected_fields_lb.get(0, tk.END)
                limit = self.limit_value_var.get() if self.limit_enabled_var.get() else -1

            cols1, cols2 = self._get_all_columns(t1_name), self._get_all_columns(t2_name)
            rows2 = self._get_rows_from_source(t2_name)
            t2_join_fields = [jc[1] for jc in join_conditions]
            table2_index = defaultdict(list)
            for row_el in rows2:
                join_key = tuple(self._get_cell_value(row_el, f) for f in t2_join_fields)
                table2_index[join_key].append(row_el)

            results = []
            t1_join_fields = [jc[0] for jc in join_conditions]
            rows1 = self._get_rows_from_source(t1_name)
            for row1_el in rows1:
                join_key = tuple(self._get_cell_value(row1_el, f) for f in t1_join_fields)
                matching_rows2 = table2_index.get(join_key, [])

                if query_type == "INNER" and matching_rows2:
                    for row2_el in matching_rows2:
                        if self._row_matches_filters(row1_el, row2_el, filter_conditions_tree):
                            result_row = {"Match_Count": len(matching_rows2)}
                            result_row.update({f"T1: {c}": self._get_cell_value(row1_el, c) for c in cols1})
                            result_row.update({f"T2: {c}": self._get_cell_value(row2_el, c) for c in cols2})
                            results.append(result_row)
                elif query_type == "ANTI" and not matching_rows2:
                    if self._row_matches_filters(row1_el, None, filter_conditions_tree):
                        result_row = {"Match_Count": 0}
                        result_row.update({f"T1: {c}": self._get_cell_value(row1_el, c) for c in cols1})
                        result_row.update({f"T2: {c}": "" for c in cols2})
                        results.append(result_row)
                if limit != -1 and len(results) >= limit:
                    break

            self.current_results_data = results
            self._display_results_grid(output_fields)
        except ValueError as e:
            messagebox.showerror("Query Error", f"Invalid condition logic: {e}", parent=self)
        except Exception as ex:
            messagebox.showerror("Execution Error", f"An error occurred: {ex}", parent=self)
            traceback.print_exc()

    def _create_simple_query_widgets(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        top_frame = ttk.Frame(parent)
        top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(top_frame, text="Select Table:").pack(side=tk.LEFT, padx=(0, 5))
        self.simple_query_table_combo = ttk.Combobox(top_frame, textvariable=self.simple_query_table_var,
                                                     state="readonly", values=self.table_names)
        self.simple_query_table_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if self.table_names:
            self.simple_query_table_var.set(self.table_names[0])

        help_text = "Example: show name, city where age > 30 and status starts with 'active'"
        ttk.Label(parent, text=help_text, font=("Segoe UI", 9, "italic"), foreground="gray").grid(row=1, column=0,
                                                                                                  columnspan=2,
                                                                                                  sticky="w", padx=5)

        self.simple_query_text = tk.Text(parent, wrap=tk.WORD, height=5, font=("Courier New", 11))
        self.simple_query_text.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=5)
        self.simple_query_text.tag_configure("error", background="#FFDDDD", underline=True)
        self.simple_query_text.bind("<KeyRelease>", self._on_simple_query_key_release)

        bottom_frame = ttk.Frame(parent)
        bottom_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        ttk.Button(bottom_frame, text="Validate & Run Query", command=self._validate_and_run_simple_query,
                   style="Accent.TButton").pack(side=tk.LEFT)
        self.simple_query_status_label = ttk.Label(bottom_frame, text="Ready.", anchor="w", foreground="green")
        self.simple_query_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

    def _create_sql_view_widgets(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        manual_actions = ttk.Frame(parent)
        manual_actions.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        ttk.Button(manual_actions, text="Run SQL", command=self._run_sql_from_view, style="Accent.TButton").pack(
            side=tk.LEFT)
        ttk.Separator(manual_actions, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill='y')
        ttk.Button(manual_actions, text="Toggle Edit Mode", command=self._toggle_manual_edit).pack(side=tk.LEFT)
        self.apply_to_designer_button = ttk.Button(manual_actions, text="Apply to Designer",
                                                   command=self._parse_and_apply_manual_query, state="disabled")
        self.apply_to_designer_button.pack(side=tk.LEFT, padx=10)

        self.query_view_text = tk.Text(parent, wrap=tk.WORD, state="disabled", font=("Courier New", 10))
        self.query_view_text.grid(row=1, column=0, sticky="nsew")
        query_sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.query_view_text.yview)
        query_sb.grid(row=1, column=1, sticky="ns")
        self.query_view_text.config(yscrollcommand=query_sb.set)
        self.query_view_text.bind("<KeyRelease>", self._on_key_release)

    def _create_results_widgets(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        action_frame = ttk.Labelframe(parent, text="Actions", padding=10)
        action_frame.grid(row=0, column=0, sticky="ew", pady=5)

        options_frame = ttk.Frame(action_frame)
        options_frame.pack(side=tk.LEFT)
        self.join_type_rb_inner = ttk.Radiobutton(options_frame, text="Inner Join", variable=self.query_type_var,
                                                  value="INNER")
        self.join_type_rb_inner.pack(side=tk.LEFT, padx=5)
        self.join_type_rb_anti = ttk.Radiobutton(options_frame, text="Left Anti-Join", variable=self.query_type_var,
                                                 value="ANTI")
        self.join_type_rb_anti.pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(options_frame, text="Limit:", variable=self.limit_enabled_var,
                        command=self._toggle_limit_entry).pack(side=tk.LEFT, padx=(15, 0))
        self.limit_spinbox = ttk.Spinbox(action_frame, from_=1, to=1000000, textvariable=self.limit_value_var, width=8,
                                         state="disabled")
        self.limit_spinbox.pack(side=tk.LEFT)

        action_button_frame = ttk.Frame(action_frame)
        action_button_frame.pack(side=tk.RIGHT)
        self.run_designer_button = ttk.Button(action_button_frame, text="Run Designer Query",
                                              command=self._run_designer_query, style="Accent.TButton")
        self.run_designer_button.pack(side=tk.LEFT, padx=10)
        ttk.Style().configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

        results_grid_frame = ttk.Labelframe(parent, text="Results", padding=10)
        results_grid_frame.grid(row=1, column=0, sticky="nsew")
        results_grid_frame.rowconfigure(0, weight=1)
        results_grid_frame.columnconfigure(0, weight=1)

        nav_frame = ttk.Frame(results_grid_frame)
        nav_frame.pack(fill='x', pady=(0, 5))
        self.results_status_label = ttk.Label(nav_frame, text="No results.")
        self.results_status_label.pack(side=tk.LEFT, padx=5)
        ttk.Label(nav_frame, text="Go to Row:").pack(side=tk.LEFT, padx=(10, 2))
        self.goto_row_var = tk.StringVar()
        self.goto_row_entry = ttk.Entry(nav_frame, textvariable=self.goto_row_var, width=8)
        self.goto_row_entry.pack(side=tk.LEFT)
        self.goto_row_entry.bind("<Return>", self._go_to_row)
        ttk.Button(nav_frame, text="Next", command=self._go_to_next_selected).pack(side=tk.LEFT, padx=2)

        self.results_tree = ttk.Treeview(results_grid_frame, show='headings', selectmode='extended')
        results_vsb = ttk.Scrollbar(results_grid_frame, orient="vertical", command=self.results_tree.yview)
        results_hsb = ttk.Scrollbar(results_grid_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=results_vsb.set, xscrollcommand=results_hsb.set)
        results_hsb.pack(side='bottom', fill='x')
        results_vsb.pack(side='right', fill='y')
        self.results_tree.pack(fill='both', expand=True)

    def _on_tab_change(self, event):
        is_visual_tab = self.config_notebook.tab(self.config_notebook.select(), "text") == "Visual Designer"
        is_sql_tab = self.config_notebook.tab(self.config_notebook.select(), "text") == "SQL View"

        self.run_designer_button.config(state="normal" if is_visual_tab or is_sql_tab else "disabled")
        if is_sql_tab:
            self.run_designer_button.config(text="Run SQL Query", command=self._run_sql_from_view)
        else:
            self.run_designer_button.config(text="Run Designer Query", command=self._run_designer_query)

    def _row_matches_simple_filters(self, row_element, filters):
        if not filters: return True
        for f in filters:
            field, op, value = f['field'], f['op'], f['value']
            is_date_func = isinstance(field, tuple)

            if is_date_func:
                func, col_name = field
                cell_value = self._get_cell_value(row_element, col_name)
                try:
                    dt_obj = None
                    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%d-%b-%y"):
                        try:
                            dt_obj = datetime.strptime(cell_value, fmt)
                            break
                        except ValueError:
                            continue
                    if not dt_obj: raise ValueError("Date format not recognized")

                    cell_part = getattr(dt_obj, func)
                    value_num = int(value)

                    op_map = {'=': cell_part == value_num, '!=': cell_part != value_num, '>': cell_part > value_num,
                              '<': cell_part < value_num, '>=': cell_part >= value_num, '<=': cell_part <= value_num}
                    match = op_map.get(op, False)
                except (ValueError, TypeError):
                    match = False
            else:
                cell_value = self._get_cell_value(row_element, field)
                cell_compare, value_compare = cell_value.lower(), value.lower()
                match = False
                if op == "CONTAINS":
                    match = value_compare in cell_compare
                elif op == "NOT CONTAINS":
                    match = value_compare not in cell_compare
                elif op == "STARTS WITH":
                    match = cell_compare.startswith(value_compare)
                elif op == "ENDS WITH":
                    match = cell_compare.endswith(value_compare)
                elif op == "=":
                    match = cell_compare == value_compare
                elif op == "!=":
                    match = cell_compare != value_compare
                else:
                    try:
                        cell_num, val_num = float(cell_value), float(value)
                        if op == ">":
                            match = cell_num > val_num
                        elif op == "<":
                            match = cell_num < val_num
                        elif op == ">=":
                            match = cell_num >= val_num
                        elif op == "<=":
                            match = cell_num <= val_num
                    except (ValueError, TypeError):
                        pass

            if not match: return False
        return True

    def _validate_and_run_simple_query(self):
        self.simple_query_text.tag_remove("error", "1.0", tk.END)
        self.simple_query_status_label.config(text="Validating...", foreground="orange")
        self.update_idletasks()

        query_text = self.simple_query_text.get("1.0", tk.END).strip()
        table_name = self.simple_query_table_var.get()

        if not query_text or not table_name:
            self.simple_query_status_label.config(text="Error: Query or table is missing.", foreground="red")
            return

        key = self.table_combobox_map.get(table_name)
        valid_columns = self.potential_tables[key]['columns']

        parsed = self._parse_simple_query(query_text, valid_columns)

        if not parsed['success']:
            self.simple_query_status_label.config(text=f"Error: {parsed['error']}", foreground="red")
            if 'span' in parsed:
                start, end = parsed['span']
                self.simple_query_text.tag_add("error", f"1.0+{start}c", f"1.0+{end}c")
            return

        self.simple_query_status_label.config(text="Query is valid. Running...", foreground="blue")
        self.update_idletasks()

        try:
            rows = self._get_rows_from_source(table_name)
            results = []

            for row in rows:
                if self._row_matches_simple_filters(row, parsed['conditions']):
                    result_row = {}
                    display_fields = parsed['fields']
                    if display_fields == ['*']:
                        display_fields = valid_columns
                    for field in display_fields:
                        result_row[field] = self._get_cell_value(row, field)
                    results.append(result_row)

            if self.limit_enabled_var.get():
                results = results[:self.limit_value_var.get()]

            self.current_results_data = results
            final_fields = parsed['fields'] if parsed['fields'] != ['*'] else valid_columns
            self._display_results_grid(final_fields)
            self.simple_query_status_label.config(text=f"Success! Found {len(results)} records.", foreground="green")

        except Exception as e:
            self.simple_query_status_label.config(text=f"Execution Error: {e}", foreground="red")
            traceback.print_exc()

    def _update_available_fields(self):
        self.available_fields_lb.delete(0, tk.END)
        self.selected_fields_lb.delete(0, tk.END)
        t1_name, t2_name = self.table1_var.get(), self.table2_var.get()
        if t1_name:
            key1 = self.table_combobox_map.get(t1_name)
            if key1:
                for col in self.potential_tables[key1]['columns']: self.available_fields_lb.insert(tk.END, f"T1: {col}")
        if t2_name:
            key2 = self.table_combobox_map.get(t2_name)
            if key2:
                for col in self.potential_tables[key2]['columns']: self.available_fields_lb.insert(tk.END, f"T2: {col}")

    def _toggle_manual_edit(self):
        self.manual_edit_mode.set(not self.manual_edit_mode.get())
        is_editing = self.manual_edit_mode.get()
        self.query_view_text.config(state="normal" if is_editing else "disabled")
        self.apply_to_designer_button.config(state="normal" if is_editing else "disabled")
        if is_editing:
            messagebox.showinfo("Manual Edit Mode",
                                "You can now edit the query text.\n"
                                "Click 'Apply to Designer' to parse your changes back into the Visual Designer.",
                                parent=self)
        else:
            self._destroy_intellisense()
            self._update_query_view()

    def _populate_ui_from_config(self, config):
        visual_config = config.get("visual_query")
        if not visual_config: return

        self.table1_var.set(visual_config.get("table1", ""))
        self.table2_var.set(visual_config.get("table2", ""))
        self._on_table_select()

        self.visual_conditions = visual_config.get("conditions_list", [])
        self.conditions_listbox.delete(0, tk.END)
        for item in self.visual_conditions:
            if item['type'] == 'op':
                display_text = item['value']
                if display_text not in ['(', ')', 'NOT']: display_text = f"  {display_text}"
                self.conditions_listbox.insert(tk.END, display_text)
            elif item['type'] == 'cond':
                data = item['data']
                if data['type'] == 'join':
                    self.conditions_listbox.insert(tk.END, f"T1.{data['t1_field']} = T2.{data['t2_field']}")
                else:
                    self.conditions_listbox.insert(tk.END,
                                                   f"{data['table']}.{data['field']} {data['op']} '{data['value']}'")

        self.query_type_var.set(visual_config.get("query_type", "INNER"))

        self.selected_fields_lb.delete(0, tk.END)
        available_items = list(self.available_fields_lb.get(0, tk.END))
        for field in visual_config.get("output_fields", []):
            if field in available_items:
                self.selected_fields_lb.insert(tk.END, field)
                try:
                    idx = list(self.available_fields_lb.get(0, tk.END)).index(field)
                    self.available_fields_lb.delete(idx)
                except ValueError:
                    pass

        self.limit_enabled_var.set(config.get("limit_enabled", False))
        self.limit_value_var.set(config.get("limit_value", 100))
        self._toggle_limit_entry()
        self._update_query_view()
        self._validate_and_highlight_conditions()

    def _save_config(self):
        filepath = filedialog.asksaveasfilename(title="Save Query Configuration", defaultextension=".json",
                                                filetypes=[("Query Config Files", "*.json")], parent=self)
        if not filepath: return

        visual_query_config = {
            "table1": self.table1_var.get(),
            "table2": self.table2_var.get(),
            "query_type": self.query_type_var.get(),
            "conditions_list": self.visual_conditions,
            "output_fields": list(self.selected_fields_lb.get(0, tk.END))
        }

        simple_query_config = {
            "table": self.simple_query_table_var.get(),
            "text": self.simple_query_text.get("1.0", tk.END).strip()
        }

        sql_view_config = {
            "text": self.query_view_text.get("1.0", tk.END).strip()
        }

        config_data = {
            "source_file": self.source_path,
            "file_type": self.file_type,
            "active_tab": self.config_notebook.tab(self.config_notebook.select(), "text"),
            "limit_enabled": self.limit_enabled_var.get(),
            "limit_value": self.limit_value_var.get(),
            "visual_query": visual_query_config,
            "simple_query": simple_query_config,
            "sql_query": sql_view_config
        }

        try:
            with open(filepath, 'w') as f:
                json.dump(config_data, f, indent=2)
            messagebox.showinfo("Success", f"Query configuration saved to:\n{filepath}", parent=self)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save configuration:\n{e}", parent=self)

    def _load_config(self):
        filepath = filedialog.askopenfilename(title="Load Query Configuration",
                                              filetypes=[("Query Config Files", "*.json")], parent=self)
        if not filepath: return
        try:
            with open(filepath, 'r') as f:
                config_data = json.load(f)

            if config_data.get("source_file") != self.source_path:
                if not messagebox.askyesno("Warning", "This query was saved for a different source file.\nLoad anyway?",
                                           parent=self):
                    return

            self._populate_ui_from_config(config_data)

            if "simple_query" in config_data:
                self.simple_query_table_var.set(config_data["simple_query"].get("table", ""))
                self.simple_query_text.delete("1.0", tk.END)
                self.simple_query_text.insert("1.0", config_data["simple_query"].get("text", ""))

            if "sql_query" in config_data:
                self.query_view_text.config(state="normal")
                self.query_view_text.delete("1.0", tk.END)
                self.query_view_text.insert("1.0", config_data["sql_query"].get("text", ""))
                self.query_view_text.config(state="disabled" if not self.manual_edit_mode.get() else "normal")

            active_tab_name = config_data.get("active_tab")
            for i, tab_name in enumerate(self.config_notebook.tabs()):
                if self.config_notebook.tab(i, "text") == active_tab_name:
                    self.config_notebook.select(i)
                    break

            self._on_tab_change(None)

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load or apply configuration:\n{e}", parent=self)
            traceback.print_exc()

    def _run_designer_query(self):
        t1_name = self.table1_var.get()
        if not t1_name:
            messagebox.showerror("Error", "Please select a table (T1).", parent=self)
            return
        if not self.selected_fields_lb.get(0, tk.END):
            messagebox.showerror("Error", "Please select at least one field for the output.", parent=self)
            return

        if self.table2_var.get():
            self._run_join_query()
        else:
            self._run_filter_query()

    def _run_sql_from_view(self):
        query_text = self.query_view_text.get("1.0", tk.END)
        config = None
        try:
            if re.search(r"\s+(INNER|LEFT\s+ANTI)(?:\s+JOIN)?\s+", query_text, re.I):
                config = self._parse_join_query(query_text)
                self._run_join_query(config)
            elif re.search(r"\s+FROM\s+", query_text, re.I):
                config = self._parse_filter_query(query_text)
                self._run_filter_query(config)
            else:
                raise ValueError("Invalid SQL. Must contain at least a SELECT and FROM clause.")
        except ValueError as e:
            messagebox.showerror("Parsing Error", f"Could not run manual query:\n\n{e}", parent=self)
        except Exception as e:
            error_details = f"An unexpected error occurred during execution:\n\n{traceback.format_exc()}"
            messagebox.showerror("Unexpected Error", error_details, parent=self)

    def _display_results_grid(self, output_fields):
        selected_iids = self.results_tree.selection()

        self.results_tree.delete(*self.results_tree.get_children())

        display_columns = list(output_fields)
        is_join = self.table1_var.get() and self.table2_var.get()
        if is_join and self.query_type_var.get() == "INNER" and "Match_Count" not in display_columns:
            display_columns.insert(0, "Match_Count")

        self.results_tree["columns"] = display_columns
        for col in display_columns:
            self.results_tree.heading(col, text=col, anchor='w')
            self.results_tree.column(col, width=120, anchor='w', stretch=True)

        if "Match_Count" in display_columns:
            self.results_tree.column("Match_Count", width=80, anchor='center', stretch=False)

        self.results_sort_col = None
        self._sort_and_redisplay_results()

        if not self.current_results_data:
            self.results_status_label.config(text="No results.")
        else:
            self.results_status_label.config(text=f"{len(self.current_results_data)} records found.")

        final_selection = [iid for iid in selected_iids if self.results_tree.exists(iid)]
        if final_selection:
            self.results_tree.selection_set(final_selection)
            self.results_tree.focus(final_selection[0])
            self.results_tree.see(final_selection[0])

    def _sort_and_redisplay_results(self):
        selected_iids = self.results_tree.selection()

        def sort_key(item_tuple):
            item = item_tuple[1]
            value = item.get(self.results_sort_col)
            if value is None:
                return -float('inf') if self.results_sort_asc else float('inf')
            try:
                return float(value)
            except (ValueError, TypeError):
                return str(value).lower()

        indexed_data = list(enumerate(self.current_results_data))

        if self.results_sort_col:
            indexed_data.sort(key=sort_key, reverse=not self.results_sort_asc)

        self._update_results_header_style()
        self.results_tree.delete(*self.results_tree.get_children())

        for original_index, result_row in indexed_data:
            iid = f"row_{original_index}"
            values = [result_row.get(col, "") for col in self.results_tree["columns"]]
            self.results_tree.insert("", "end", iid=iid, values=values)

        if selected_iids:
            self.results_tree.selection_set(selected_iids)

    def _go_to_row(self, event=None):
        try:
            row_num_str = self.goto_row_var.get()
            if not row_num_str: return
            row_num = int(row_num_str)
            children = self.results_tree.get_children()
            if 1 <= row_num <= len(children):
                target_iid = children[row_num - 1]
                self.results_tree.selection_set(target_iid)
                self.results_tree.focus(target_iid)
                self.results_tree.see(target_iid)
            else:
                messagebox.showwarning("Invalid Row", f"Please enter a row number between 1 and {len(children)}.",
                                       parent=self)
        except (ValueError, IndexError):
            messagebox.showwarning("Invalid Row", "Please enter a valid row number.", parent=self)

    def _go_to_next_selected(self):
        selection = self.results_tree.selection()
        if not selection: return

        children = self.results_tree.get_children()
        try:
            last_selected_iid = selection[-1]
            current_index = children.index(last_selected_iid)
            if current_index + 1 < len(children):
                next_iid = children[current_index + 1]
                self.results_tree.selection_set(next_iid)
                self.results_tree.focus(next_iid)
                self.results_tree.see(next_iid)
        except ValueError:
            pass

    def _populate_initial_dropdowns(self):
        self.table1_combo['values'] = self.table_names
        self.table2_combo['values'] = self.table_names_with_blank

    def _add_output_field(self):
        selected_indices = self.available_fields_lb.curselection()
        agg_func = self.aggregate_func_var.get()

        for i in reversed(selected_indices):
            field = self.available_fields_lb.get(i)
            if agg_func != "None":
                output_text = f"{agg_func}({field})"
            else:
                output_text = field
            self.selected_fields_lb.insert(tk.END, output_text)
        self._update_query_view()

    def _add_all_output_fields(self):
        all_items = self.available_fields_lb.get(0, tk.END)
        for item in all_items: self.selected_fields_lb.insert(tk.END, item)
        self.available_fields_lb.delete(0, tk.END)
        self._update_query_view()

    def _remove_output_field(self):
        selected = self.selected_fields_lb.curselection()
        for i in reversed(selected):
            self.selected_fields_lb.delete(i)
        self._update_query_view()

    def _run_query_engine(self, base_results):
        group_by_fields = self.grouped_by_lb.get(0, tk.END)
        output_fields = self.selected_fields_lb.get(0, tk.END)

        if not group_by_fields:
            limit = self.limit_value_var.get() if self.limit_enabled_var.get() else -1
            if limit != -1:
                return base_results[:limit]
            return base_results

        grouped_data = defaultdict(list)
        for row in base_results:
            key = tuple(row.get(field, "") for field in group_by_fields)
            grouped_data[key].append(row)

        aggregated_results = []
        for key, rows in grouped_data.items():
            agg_row = {}
            for i, field in enumerate(group_by_fields):
                agg_row[field] = key[i]

            for out_field in output_fields:
                if out_field in agg_row:
                    continue

                match = re.match(r"(\w+)\((.+)\)", out_field)
                if not match:
                    agg_row[out_field] = "N/A (Non-aggregated field in GROUP BY query)"
                    continue

                func, field_to_agg = match.groups()

                values_to_agg = [row.get(field_to_agg, 0) for row in rows]

                numeric_values = []
                for v in values_to_agg:
                    try:
                        numeric_values.append(float(v))
                    except (ValueError, TypeError):
                        continue

                if func.upper() == "COUNT":
                    agg_row[out_field] = len(values_to_agg)
                elif func.upper() == "SUM":
                    agg_row[out_field] = sum(numeric_values)
                elif func.upper() == "AVG":
                    agg_row[out_field] = sum(numeric_values) / len(numeric_values) if numeric_values else 0
                elif func.upper() == "MIN":
                    agg_row[out_field] = min(numeric_values) if numeric_values else ""
                elif func.upper() == "MAX":
                    agg_row[out_field] = max(numeric_values) if numeric_values else ""

            aggregated_results.append(agg_row)

        limit = self.limit_value_var.get() if self.limit_enabled_var.get() else -1
        if limit != -1:
            return aggregated_results[:limit]
        return aggregated_results

    def _remove_all_output_fields(self):
        all_items = self.selected_fields_lb.get(0, tk.END)
        for item in all_items: self.available_fields_lb.insert(tk.END, item)
        self.selected_fields_lb.delete(0, tk.END)
        self._update_query_view()

    def _move_output_field(self, direction):
        selected_indices = self.selected_fields_lb.curselection()
        if not selected_indices: return
        for i in selected_indices:
            if not 0 <= i + direction < self.selected_fields_lb.size(): continue
            text = self.selected_fields_lb.get(i)
            self.selected_fields_lb.delete(i)
            self.selected_fields_lb.insert(i + direction, text)
            self.selected_fields_lb.selection_set(i + direction)
        self._update_query_view()

    def _toggle_limit_entry(self):
        state = "normal" if self.limit_enabled_var.get() else "disabled"
        self.limit_spinbox.config(state=state)
        self._update_query_view()

    def _parse_and_apply_manual_query(self):
        query_text = self.query_view_text.get("1.0", tk.END)
        config = None
        try:
            if re.search(r"\s+(INNER|LEFT\s+ANTI)(?:\s+JOIN)?\s+", query_text, re.I):
                config = self._parse_join_query(query_text)
            elif re.search(r"\s+FROM\s+", query_text, re.I):
                config = self._parse_filter_query(query_text)
            else:
                raise ValueError("Invalid SQL. Must contain at least a SELECT and FROM clause.")

            self.table1_var.set(config.get("table1", ""))
            self.table2_var.set(config.get("table2", ""))
            self._on_table_select()

            if config.get("mode") == "join":
                for t1_field, t2_field in config.get("join_conditions", []):
                    self.condition_type_var.set("join")
                    self.field1_var.set(t1_field)
                    self.field2_var.set(t2_field)
                    self._add_condition()

            self.condition_type_var.set("filter")
            for table_alias, field, op, value in config.get("filter_conditions", []):
                full_field = f"{table_alias}: {field}" if config.get("table2") else field
                self.filter_field_var.set(full_field)
                self.filter_op_var.set(op)
                self.filter_value_var.set(value)
                self._add_condition()
                self._add_logical_operator("AND")

            if self.conditions_listbox.size() > 0:
                last_item = self.visual_conditions[-1]
                if last_item['type'] == 'op' and last_item['value'] == 'AND':
                    self.conditions_listbox.delete(tk.END)
                    self.visual_conditions.pop()

            self._update_available_fields()
            self.selected_fields_lb.delete(0, tk.END)
            for field in config.get("output_fields", []):
                self.selected_fields_lb.insert(tk.END, field)
                try:
                    idx = list(self.available_fields_lb.get(0, tk.END)).index(field)
                    self.available_fields_lb.delete(idx)
                except ValueError:
                    pass
            self._validate_and_highlight_conditions()
            messagebox.showinfo("Success", "Manual query applied to the designer.", parent=self)
            self._update_query_view()
        except ValueError as e:
            messagebox.showerror("Parsing Error", f"Could not apply manual query:\n\n{e}", parent=self)
        except Exception:
            error_details = f"An unexpected error occurred:\n\n{traceback.format_exc()}"
            messagebox.showerror("Unexpected Error", error_details, parent=self)

    def _parse_join_query(self, text):
        select_match = re.search(r"SELECT\s*(.*?)\s*FROM", text, re.S | re.I)
        if not select_match: raise ValueError("Could not find SELECT clause.")
        fields_str = select_match.group(1).strip()
        output_fields = [f.strip() for f in re.split(r'\s*,\s*', fields_str) if f.strip()]

        from_join_match = re.search(
            r"FROM\s+'([^']+)'\s+AS\s+(T1)\s+(INNER(?:\s+JOIN)?|LEFT\s+ANTI(?:-JOIN)?)\s+'([^']+)'\s+AS\s+(T2)", text,
            re.I | re.S)
        if not from_join_match: raise ValueError("Could not parse FROM/JOIN clause.")
        t1, a1, join_type_raw, t2, a2 = from_join_match.groups()
        if a1.upper() != "T1" or a2.upper() != "T2": raise ValueError("Must use aliases T1 and T2.")
        join_type = "ANTI" if "ANTI" in join_type_raw.upper() else "INNER"

        on_match = re.search(r"\s+ON\s+(.*?)(?=\s*(?:WHERE|LIMIT|;|$))", text, re.S | re.I)
        if not on_match: raise ValueError("Could not find ON clause.")
        conditions_str = on_match.group(1).strip()
        join_conditions = []
        for part in re.split(r"\s+AND\s+", conditions_str, flags=re.I):
            match = re.fullmatch(r"T1\.([\w\.\s-]+)\s*=\s*T2\.([\w\.\s-]+)", part.strip(), re.I)
            if not match: raise ValueError(f"Invalid join condition: '{part}'")
            join_conditions.append([f.strip() for f in match.groups()])

        where_match = re.search(r"\s+WHERE\s+(.*?)(?=\s*(?:LIMIT|;|$))", text, re.S | re.I)
        filter_conditions = []
        if where_match:
            conditions_str = where_match.group(1).strip()
            op_pattern = r"CONTAINS|NOT\s*CONTAINS|STARTS\s*WITH|ENDS\s*WITH|[<>=!]+"
            for part in re.split(r"\s+AND\s+", conditions_str, flags=re.I):
                match = re.match(fr"(T1|T2)\.([\w\.\s-]+)\s+({op_pattern})\s+'([^']*)'", part.strip(), re.I)
                if not match: continue
                alias, field, op, value = match.groups()
                filter_conditions.append((alias.upper(), field.strip(), op.upper().replace(" ", ""), value))

        limit_match = re.search(r"LIMIT\s*(\d+)", text, re.I)
        return {
            "mode": "join", "table1": t1, "table2": t2, "join_conditions": join_conditions,
            "filter_conditions": filter_conditions,
            "output_fields": output_fields, "query_type": join_type,
            "limit_enabled": bool(limit_match), "limit_value": int(limit_match.group(1)) if limit_match else 100
        }

    def _parse_filter_query(self, text):
        select_match = re.search(r"SELECT\s*(.*?)\s*FROM", text, re.S | re.I)
        if not select_match: raise ValueError("Could not find SELECT clause.")
        fields_str = select_match.group(1).strip()
        output_fields = [f.strip() for f in re.split(r'\s*,\s*', fields_str) if f.strip()]

        from_match = re.search(r"FROM\s+'([^']+)'(?:\s+AS\s+(T1))?", text, re.I | re.S)
        if not from_match: raise ValueError("Could not parse FROM clause.")
        t1, alias = from_match.groups()
        if alias and alias.upper() != "T1": raise ValueError("Alias for single table must be T1 if provided.")

        where_match = re.search(r"\s+WHERE\s+(.*?)(?=\s*(?:LIMIT|;|$))", text, re.S | re.I)
        filter_conditions = []
        if where_match:
            conditions_str = where_match.group(1).strip()
            op_pattern = r"CONTAINS|NOT\s*CONTAINS|STARTS\s*WITH|ENDS\s*WITH|[<>=!]+"
            for part in re.split(r"\s+AND\s+", conditions_str, flags=re.I):
                match = re.match(fr"(?:T1\.)?([\w\.\s-]+)\s+({op_pattern})\s+'([^']*)'", part.strip(), re.I)
                if not match: continue
                field, op, value = match.groups()
                filter_conditions.append(("T1", field.strip(), op.upper().replace(" ", ""), value))

        limit_match = re.search(r"LIMIT\s*(\d+)", text, re.I)
        return {
            "mode": "filter", "table1": t1, "table2": "", "filter_conditions": filter_conditions,
            "output_fields": output_fields,
            "limit_enabled": bool(limit_match), "limit_value": int(limit_match.group(1)) if limit_match else 100
        }

    def _export_results(self):
        if not self.results_tree.get_children():
            messagebox.showwarning("Export Error", "There are no results to export.", parent=self)
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                                parent=self)
        if not filepath: return
        try:
            delimiter = self.app.csv_delimiter if hasattr(self.app, 'csv_delimiter') else ','
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=delimiter)
                writer.writerow(self.results_tree["columns"])
                for item_id in self.results_tree.get_children():
                    writer.writerow(self.results_tree.item(item_id, "values"))
            messagebox.showinfo("Success", "Results exported successfully.", parent=self)
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred during export:\n{e}", parent=self)

    def _get_rows_from_source(self, table_name):
        key = self.table_combobox_map[table_name]
        if self.file_type == 'xml':
            table_info = self.potential_tables[key]
            return table_info["parent_element"].findall(table_info["row_tag"])
        elif self.file_type == 'csv':
            return self.table_data_cache.get(key, [])
        return []

    def _get_cell_value(self, row, column):
        if self.file_type == 'xml':
            return row.findtext(column, "").strip()
        elif self.file_type == 'csv':
            return row.get(column, "").strip()
        return ""

    def _get_all_columns(self, table_name):
        key = self.table_combobox_map[table_name]
        return self.potential_tables[key]['columns']

    def _on_results_click(self, event):
        if self.results_tree.identify("region", event.x, event.y) == "heading":
            col_id_str = self.results_tree.identify_column(event.x)
            col_id = self.results_tree.column(col_id_str, "id")
            self._sort_by_column(col_id)

    def _sort_by_column(self, col_name):
        if self.results_sort_col == col_name:
            self.results_sort_asc = not self.results_sort_asc
        else:
            self.results_sort_col, self.results_sort_asc = col_name, True
        self._sort_and_redisplay_results()

    def _resize_query_results_columns(self):
        tree = self.results_tree
        if not tree["columns"]: return
        try:
            items_to_sample = tree.get_children()[:VIRTUAL_TABLE_ROW_COUNT]

            style = ttk.Style()
            font_name = style.lookup("Treeview", "font") or "TkDefaultFont"
            tree_font = tkFont.Font(font=font_name)
            padding = 20

            for col_id in tree["columns"]:
                header_text = tree.heading(col_id, "text").split(" ")[0]
                max_width = tree_font.measure(header_text) + padding

                col_index = tree["columns"].index(col_id)
                for item_id in items_to_sample:
                    values = tree.item(item_id, "values")
                    try:
                        cell_value = str(values[col_index]) if values and col_index < len(values) else ""
                        cell_width = tree_font.measure(cell_value) + padding
                        if cell_width > max_width: max_width = cell_width
                    except (ValueError, IndexError):
                        continue

                max_width = min(max(max_width, 50), 500)
                tree.column(col_id, width=max_width, stretch=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resize columns: {e}", parent=self)

    def _on_simple_query_key_release(self, event):
        self._destroy_simple_query_intellisense()
        cursor_index = self.simple_query_text.index(tk.INSERT)
        line_text_before_cursor = self.simple_query_text.get("1.0", cursor_index)

        if event.char == '.':
            last_word = re.split(r'[\s,]', line_text_before_cursor[:-1])[-1]
            if last_word.lower() == self.simple_query_table_var.get().lower():
                self._show_simple_query_intellisense('')
        elif event.char.isalnum() or event.keysym == 'BackSpace':
            last_separator_pos = -1
            for sep in [' ', ',', '.']:
                last_separator_pos = max(last_separator_pos, line_text_before_cursor.rfind(sep))

            start_of_word = last_separator_pos + 1
            current_word = line_text_before_cursor[start_of_word:]

            if len(current_word) > 0:
                self._show_simple_query_intellisense(current_word)

    def _show_simple_query_intellisense(self, prefix):
        self._destroy_simple_query_intellisense()
        table_name = self.simple_query_table_var.get()
        if not table_name: return
        key = self.table_combobox_map.get(table_name)
        if not key: return

        all_columns = self.potential_tables[key]['columns']
        matching_columns = [c for c in all_columns if c.lower().startswith(prefix.lower())]

        if not matching_columns: return

        bbox = self.simple_query_text.bbox(tk.INSERT)
        if not bbox: return
        x, y, _, height = bbox

        self.simple_query_intellisense_popup = tk.Toplevel(self)
        self.simple_query_intellisense_popup.wm_overrideredirect(True)

        listbox = tk.Listbox(self.simple_query_intellisense_popup, height=min(10, len(matching_columns)))
        listbox.pack()

        for col in matching_columns:
            listbox.insert(tk.END, col)

        popup_x = self.simple_query_text.winfo_rootx() + x
        popup_y = self.simple_query_text.winfo_rooty() + y + height
        self.simple_query_intellisense_popup.wm_geometry(f"+{popup_x}+{popup_y}")

        listbox.focus_set()
        listbox.bind("<Double-Button-1>", self._on_simple_intellisense_select)
        listbox.bind("<Return>", self._on_simple_intellisense_select)
        listbox.bind("<Escape>", lambda e: self._destroy_simple_query_intellisense())

    def _on_simple_intellisense_select(self, event):
        listbox = event.widget
        selection_indices = listbox.curselection()
        if not selection_indices: return

        selected_field = listbox.get(selection_indices[0])

        cursor_index = self.simple_query_text.index(tk.INSERT)
        line_text_before_cursor = self.simple_query_text.get("1.0", cursor_index)

        last_separator_pos = -1
        for sep in [' ', ',', '.']:
            last_separator_pos = max(last_separator_pos, line_text_before_cursor.rfind(sep))

        start_of_word = last_separator_pos + 1

        self.simple_query_text.delete(f"1.0+{start_of_word}c", cursor_index)
        self.simple_query_text.insert(f"1.0+{start_of_word}c", selected_field)

        self._destroy_simple_query_intellisense()
        return "break"

    def _destroy_simple_query_intellisense(self):
        if self.simple_query_intellisense_popup:
            self.simple_query_intellisense_popup.destroy()
            self.simple_query_intellisense_popup = None
            self.simple_query_text.focus_set()

    def _on_key_release(self, event):
        if not self.manual_edit_mode.get():
            return
        if event.char != ":":
            self._destroy_intellisense()
            return
        cursor_index = self.query_view_text.index(tk.INSERT)
        line, char = map(int, cursor_index.split('.'))
        if char < 3:
            self._destroy_intellisense()
            return
        start_index = f"{line}.{char - 3}"
        prefix = self.query_view_text.get(start_index, cursor_index).upper()
        if prefix == "T1:":
            self._show_intellisense("T1")
        elif prefix == "T2:":
            self._show_intellisense("T2")
        else:
            self._destroy_intellisense()

    def _show_intellisense(self, alias):
        self._destroy_intellisense()
        table_var = self.table1_var if alias == "T1" else self.table2_var
        table_name = table_var.get()
        if not table_name: return
        key = self.table_combobox_map.get(table_name)
        if not key: return
        columns = self.potential_tables[key]['columns']
        if not columns: return
        bbox = self.query_view_text.bbox(tk.INSERT)
        if not bbox: return
        x, y, _, height = bbox
        self.intellisense_popup = tk.Frame(self, relief='solid', borderwidth=1)
        self.intellisense_popup.alias = alias
        listbox = tk.Listbox(self.intellisense_popup, height=min(10, len(columns)))
        scrollbar = ttk.Scrollbar(self.intellisense_popup, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for col in columns: listbox.insert(tk.END, col)
        popup_x = self.query_view_text.winfo_rootx() + x
        popup_y = self.query_view_text.winfo_rooty() + y + height
        self.intellisense_popup.place(x=popup_x - self.winfo_rootx(), y=popup_y - self.winfo_rooty())
        listbox.focus_set()
        listbox.bind("<Double-Button-1>", self._on_intellisense_select)
        listbox.bind("<Return>", self._on_intellisense_select)
        listbox.bind("<Escape>", lambda e: self._destroy_intellisense())

    def _on_intellisense_select(self, event):
        listbox = event.widget
        selection_indices = listbox.curselection()
        if not selection_indices: return
        selected_field = listbox.get(selection_indices[0])
        self.query_view_text.insert(tk.INSERT, selected_field)
        self._destroy_intellisense()
        return "break"

    def _destroy_intellisense(self):
        if self.intellisense_popup:
            self.intellisense_popup.destroy()
            self.intellisense_popup = None
            self.query_view_text.focus_set()


class XMLNotepad:
    def __init__(self, root):
        self.root = root
        self.root.title("XML/CSV Notepad")
        self.root.geometry("1300x850")
        self.file_type = None
        self.xml_tree_root = None
        self.tree_item_to_element = {}
        self.selected_element_for_context_menu = None
        self.potential_tables = {}
        self.table_combobox_map = {}
        self.current_loaded_filepath = ""
        self.csv_delimiter = ','
        self.undo_stack = deque(maxlen=UNDO_STACK_SIZE)
        self.redo_stack = deque(maxlen=UNDO_STACK_SIZE)
        self.table_data_cache = {}
        self.open_table_tabs = {}
        self._pressed_close_tab_index = None

        self.status_var = tk.StringVar()
        self.filename_display_var = tk.StringVar(value="No file loaded.")
        self.selected_table_var = tk.StringVar()

        self.help_window = None
        self.transactional_checker_window = None

        self._setup_styles()
        self.setup_menu()
        self.setup_top_controls_frame()
        self.setup_paned_window()
        self.setup_context_menu()
        self.setup_status_frame()

        self.root.bind_all("<Control-o>", lambda event: self.open_xml_file_threaded())
        self.root.bind_all("<Control-s>", self.handle_ctrl_s)
        self.root.bind_all("<Control-f>", lambda event: self.show_find_dialog())
        self.root.bind_all("<Control-g>", lambda event: self._handle_goto_row())
        self.root.bind_all("<Control-z>", lambda event: self.undo_action())
        self.root.bind_all("<Control-y>", lambda event: self.redo_action())



    def _setup_styles(self):
        style = ttk.Style()

        active_tab_bg = "#D0E8FF"
        inactive_tab_bg = "#EAEAEA"

        try:
            # Create the image elements needed for the tabs
            self.close_image = tk.PhotoImage("close_image", data='''
                R0lGODlhCAAIAMIBAAAAADs7O4+Pj9nZ2Ts7Ozs7Ozs7Ozs7OyH+EUNyZWF0ZWQg
                d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU
                5kEJADs=
            ''')
            self.spacer_image = tk.PhotoImage("spacer_image", data='''
                R0lGODlhBQAFAPAAAAAAAP///yH5BAEAAAAALAAAAAAFAAUAAAIIhI+ZweAIAOw==
            ''')

            # Define our custom elements with a unique prefix to avoid conflicts
            style.element_create("Custom.close", "image", "close_image")
            style.element_create("Custom.spacer", "image", "spacer_image")
        except tk.TclError:
            # This can happen if the code is run multiple times in the same session
            pass

        # Clone the default TNotebook layout into our new CustomNotebook style
        style.layout("CustomNotebook", style.layout("TNotebook"))

        # Configure the default appearance for tabs in our custom style
        style.configure("CustomNotebook.Tab", padding=[5, 1], background=inactive_tab_bg)

        # Map the 'selected' state to a different background color. This is the most reliable method.
        style.map("CustomNotebook.Tab", background=[("selected", active_tab_bg)])

        # Modify the layout for a single tab, using the ORIGINAL ttk element names
        style.layout("CustomNotebook.Tab", [
            ("Notebook.tab", {  # Use the original 'Notebook.tab' element
                "sticky": "nswe",
                "children": [
                    ("Notebook.padding", {  # Use 'Notebook.padding'
                        "side": "top",
                        "sticky": "nswe",
                        "children": [
                            ("Notebook.focus", {  # Use 'Notebook.focus'
                                "side": "top",
                                "sticky": "nswe",
                                "children": [
                                    ("Notebook.label", {"side": "left", "sticky": ''}),
                                    ("Custom.spacer", {"side": "left", "sticky": ''}),
                                    ("Custom.close", {"side": "left", "sticky": ''}),
                                ]
                            })
                        ]
                    })
                ]
            })
        ])

    def setup_menu(self):
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)

        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Open XML...", command=self.open_xml_file_threaded, accelerator="Ctrl+O")
        self.filemenu.add_command(label="Open CSV...", command=self.open_csv_file_threaded)
        self.filemenu.add_command(label="Close File", command=self.close_current_file, state="disabled")
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Save XML As...", command=self.save_xml_as, state="disabled")
        self.filemenu.add_command(label="Save Table as CSV...", command=self._save_current_table_as_csv,
                                  accelerator="Ctrl+S", state="disabled")
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=self.root.quit)
        self.menubar.add_cascade(label="File", menu=self.filemenu)

        self.editmenu = tk.Menu(self.menubar, tearoff=0)
        self.editmenu.add_command(label="Undo", command=self.undo_action, accelerator="Ctrl+Z", state="disabled")
        self.editmenu.add_command(label="Redo", command=self.redo_action, accelerator="Ctrl+Y", state="disabled")
        self.editmenu.add_separator()
        self.editmenu.add_command(label="Find...", command=self.show_find_dialog, accelerator="Ctrl+F",
                                  state="disabled")
        self.editmenu.add_command(label="Batch Operations...", command=self.open_batch_operations_dialog,
                                  state="disabled")
        self.editmenu.add_separator()
        self.editmenu.add_command(label="Go to Row...", command=self._handle_goto_row, accelerator="Ctrl+G",
                                  state="disabled")
        self.editmenu.add_command(label="Resize Columns", command=self.resize_columns, state="disabled")
        self.editmenu.add_command(label="Reorder Tabs...", command=self.open_tab_reorder_dialog, state="disabled")
        self.menubar.add_cascade(label="Edit", menu=self.editmenu)

        self.utilsmenu = tk.Menu(self.menubar, tearoff=0)
        self.utilsmenu.add_command(label="Query Designer...", command=self.open_query_designer, state="disabled")
        self.utilsmenu.add_command(label="Transactional Data Check...", command=self.open_transactional_checker,
                                   state="disabled")
        self.utilsmenu.add_separator()
        self.utilsmenu.add_command(label="Generate XSD from XML...", command=self._generate_xsd)
        self.utilsmenu.add_command(label="Validate XML with XSD...", command=self._validate_with_xsd)
        self.utilsmenu.add_separator()
        self.utilsmenu.add_command(label="Set CSV Delimiter...", command=self._set_csv_delimiter)
        self.menubar.add_cascade(label="Utils", menu=self.utilsmenu)

        self.helpmenu = tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="Help Topics", command=self.show_help_window)
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)

    def open_tab_reorder_dialog(self):
        TabReorderDialog(self.root, self.content_notebook)

        current_tab = self.get_current_table_tab()
        if current_tab:
            TabReorderDialog(self.root, self, current_tab)

    def open_batch_operations_dialog(self):
        current_tab = self.get_current_table_tab()
        if current_tab:
            BatchOperationsDialog(self.root, self, current_tab)

    def setup_top_controls_frame(self):
        self.top_controls_frame = ttk.Frame(self.root)
        self.top_controls_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))
        self.tables_label = ttk.Label(self.top_controls_frame, text="Tables:")
        self.tables_label.pack(side=tk.LEFT, padx=(0, 5))
        self.tables_combobox = ttk.Combobox(self.top_controls_frame, textvariable=self.selected_table_var,
                                            state="disabled", width=35)
        self.tables_combobox.pack(side=tk.LEFT, padx=(0, 10))
        self.tables_combobox.bind("<<ComboboxSelected>>", self.on_table_combobox_select)
        self.filename_label_prefix = ttk.Label(self.top_controls_frame, text="File:")
        self.filename_label_prefix.pack(side=tk.LEFT, padx=(10, 2))
        self.filename_display_label = ttk.Label(self.top_controls_frame, textvariable=self.filename_display_var,
                                                relief="sunken", padding=2, anchor="w")
        self.filename_display_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def setup_paned_window(self):
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(expand=True, fill='both', padx=5, pady=5)

        self.left_panel_container = ttk.Frame(self.paned_window)
        self.paned_window.add(self.left_panel_container, weight=1)
        self.tree_frame_outer = ttk.Labelframe(self.left_panel_container, text="XML Tree", padding="5")
        self.tree_frame_outer.pack(side=tk.TOP, expand=True, fill='both')
        self.tree = ttk.Treeview(self.tree_frame_outer, selectmode='browse')
        self.tree_vsb = ttk.Scrollbar(self.tree_frame_outer, orient="vertical", command=self.tree.yview)
        self.tree_hsb = ttk.Scrollbar(self.tree_frame_outer, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_vsb.set, xscrollcommand=self.tree_hsb.set)
        self.tree_vsb.pack(side='right', fill='y')
        self.tree_hsb.pack(side='bottom', fill='x')
        self.tree.pack(side='left', expand=True, fill='both')
        self.tree.bind('<<TreeviewSelect>>', self.on_xml_tree_node_select)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<<TreeviewOpen>>", self.on_tree_node_expand)

        self.right_panel_outer = ttk.Labelframe(self.paned_window, text="Content", padding="5")
        self.paned_window.add(self.right_panel_outer, weight=3)

        self.content_notebook = ttk.Notebook(self.right_panel_outer, style="CustomNotebook")
        self.content_notebook.pack(expand=True, fill='both')
        self.content_notebook.bind("<ButtonPress-1>", self._on_tab_close_press)
        self.content_notebook.bind("<ButtonRelease-1>", self._on_tab_close_release)

        details_frame = ttk.Frame(self.content_notebook)
        self.content_notebook.add(details_frame, text="Node Details")

        self.detail_tree = ttk.Treeview(details_frame, columns=("Property", "Value"), show='headings')
        self.detail_tree.heading("Property", text="Property", anchor='w')
        self.detail_tree.heading("Value", text="Value", anchor='w')
        self.detail_tree.column("Property", width=220, stretch=tk.NO, anchor='w')
        self.detail_tree.column("Value", width=400, anchor='w')
        self.detail_tree_vsb = ttk.Scrollbar(details_frame, orient="vertical", command=self.detail_tree.yview)
        self.detail_tree.configure(yscrollcommand=self.detail_tree_vsb.set)
        self.detail_tree.pack(side='left', fill="both", expand=True)
        self.detail_tree_vsb.pack(side='right', fill='y')

    def setup_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Export Node as CSV...", command=self.export_node_as_csv, state="disabled")

    def setup_status_frame(self):
        self.status_frame = ttk.Frame(self.root, padding=(2, 2))
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = ttk.Label(self.status_frame, textvariable=self.status_var, anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(self.status_frame, orient='horizontal', mode='determinate')

    def update_status(self, message, show_progress=False, progress_value=0):
        self.status_var.set(message)
        if show_progress:
            if not self.progress_bar.winfo_ismapped():
                self.progress_bar.pack(side=tk.RIGHT, padx=5, pady=2, fill=tk.X, expand=False, ipadx=50)
            self.progress_bar['value'] = progress_value
        else:
            if self.progress_bar.winfo_ismapped():
                self.progress_bar.pack_forget()
        self.root.update_idletasks()

    def _on_tab_close_press(self, event):
        element = self.content_notebook.identify(event.x, event.y)
        if "close" in element:
            index = self.content_notebook.index(f"@{event.x},{event.y}")
            self.content_notebook.state(['pressed'])
            self._pressed_close_tab_index = index
            return "break"

    def _on_tab_close_release(self, event):
        if not self.content_notebook.instate(['pressed']):
            return
        element = self.content_notebook.identify(event.x, event.y)
        try:
            index = self.content_notebook.index(f"@{event.x},{event.y}")
            if "close" in element and self._pressed_close_tab_index == index:
                self.close_tab(index)
        except tk.TclError:
            pass
        finally:
            self.content_notebook.state(['!pressed'])
            self._pressed_close_tab_index = None

    def close_tab(self, index):
        if index == 0: return
        tab_widget = self.content_notebook.winfo_children()[index]
        if hasattr(tab_widget, 'internal_key'):
            key_to_remove = tab_widget.internal_key
            if key_to_remove in self.open_table_tabs:
                del self.open_table_tabs[key_to_remove]
        self.content_notebook.forget(index)

        if len(self.content_notebook.tabs()) <= 2:
            self.editmenu.entryconfig("Reorder Tabs...", state="disabled")
        if len(self.content_notebook.tabs()) <= 1:
            self.filemenu.entryconfig("Save Table as CSV...", state="disabled")
            self.editmenu.entryconfig("Find...", state="disabled")
            self.editmenu.entryconfig("Batch Operations...", state="disabled")
            self.editmenu.entryconfig("Go to Row...", state="disabled")
            self.editmenu.entryconfig("Resize Columns", state="disabled")


    def get_current_table_tab(self):
        try:
            selected_tab = self.content_notebook.select()
            if not selected_tab: return None
            tab_widget = self.root.nametowidget(selected_tab)
            if isinstance(tab_widget, TableViewTab):
                return tab_widget
            return None
        except (tk.TclError, KeyError):
            return None

    def push_undo(self, undo_data):
        self.undo_stack.append(undo_data)
        self.redo_stack.clear()
        self.editmenu.entryconfig("Undo", state="normal")
        self.editmenu.entryconfig("Redo", state="disabled")

    def handle_ctrl_s(self, event=None):
        current_tab = self.get_current_table_tab()
        if current_tab:
            self._save_current_table_as_csv()
        return "break"

    def resize_columns(self):
        current_tab = self.get_current_table_tab()
        if current_tab:
            current_tab.resize_columns()

    def show_help_window(self):
        if self.help_window is None or not self.help_window.winfo_exists():
            self.help_window = HelpWindow(self.root)
        self.help_window.focus()

    def open_query_designer(self):
        if not self.potential_tables:
            messagebox.showwarning("Query Designer", "No tables found to query.", parent=self.root)
            return
        QueryDesigner(self.root, self, self.potential_tables, self.table_combobox_map, self.current_loaded_filepath,
                      self.file_type, self.table_data_cache)

    def open_transactional_checker(self):
        if not self.potential_tables:
            messagebox.showwarning("Transactional Check", "No tables found to check.", parent=self.root)
            return
        if self.transactional_checker_window is None or not self.transactional_checker_window.winfo_exists():
            self.transactional_checker_window = TransactionalDataChecker(
                self.root, self, self.potential_tables, self.table_combobox_map,
                self.current_loaded_filepath, self.file_type, self.table_data_cache
            )
        self.transactional_checker_window.focus()

    def _set_csv_delimiter(self):
        new_delimiter = simpledialog.askstring("Set Delimiter", "Enter the single character for CSV delimiter:",
                                               parent=self.root)
        if new_delimiter and len(new_delimiter) == 1:
            self.csv_delimiter = new_delimiter
            messagebox.showinfo("Success", f"CSV delimiter has been set to '{self.csv_delimiter}'.", parent=self.root)
        elif new_delimiter is not None:
            messagebox.showwarning("Invalid Input", "Delimiter must be a single character.", parent=self.root)

    def _generate_xsd(self):
        if self.file_type != 'xml' or not self.current_loaded_filepath:
            messagebox.showerror("Error", "Please open an XML file first.", parent=self.root)
            return

        class XsdGenerator:
            def __init__(self):
                self.XS_NS = "http://www.w3.org/2001/XMLSchema"
                self.NSMAP = {'xs': self.XS_NS}
                self.defined_types = {}

            def _get_type_name(self, element):
                return f"{element.tag}Type"

            def _infer_type(self, text):
                if not text or not text.strip():
                    return "xs:string"
                try:
                    int(text)
                    return "xs:integer"
                except (ValueError, TypeError):
                    pass
                try:
                    float(text)
                    return "xs:decimal"
                except (ValueError, TypeError):
                    pass
                if re.match(r'\d{4}-\d{2}-\d{2}', text.strip()):
                    return "xs:date"
                return "xs:string"

            def _build_element_definition(self, element):
                if not list(element) and not element.attrib:
                    el_xsd = etree.Element(f"{{{self.XS_NS}}}element", name=element.tag)
                    el_xsd.set("type", self._infer_type(element.text))
                    return el_xsd

                type_name = self._get_type_name(element)
                if type_name not in self.defined_types:
                    complex_type = etree.Element(f"{{{self.XS_NS}}}complexType", name=type_name)
                    self.defined_types[type_name] = complex_type

                    if list(element):
                        sequence = etree.Element(f"{{{self.XS_NS}}}sequence")
                        unique_child_tags = sorted(list(set(c.tag for c in element)))
                        for tag in unique_child_tags:
                            first_child = element.find(tag)
                            child_xsd = self._build_element_definition(first_child)
                            child_xsd.set("minOccurs", "0")
                            child_xsd.set("maxOccurs", "unbounded")
                            sequence.append(child_xsd)
                        complex_type.append(sequence)

                    for name, value in element.attrib.items():
                        attr_xsd = etree.Element(f"{{{self.XS_NS}}}attribute", name=name)
                        attr_xsd.set("type", self._infer_type(value))
                        attr_xsd.set("use", "optional")
                        complex_type.append(attr_xsd)

                el_xsd = etree.Element(f"{{{self.XS_NS}}}element", name=element.tag, type=type_name)
                return el_xsd

            def generate(self, root_element):
                schema = etree.Element(f"{{{self.XS_NS}}}schema", nsmap=self.NSMAP)
                schema.append(self._build_element_definition(root_element))
                for type_name in sorted(self.defined_types.keys()):
                    schema.append(self.defined_types[type_name])
                return schema

        try:
            self.update_status("Generating XSD Schema...")
            self.root.update_idletasks()

            parsed_xml = etree.parse(self.current_loaded_filepath)
            root_element = parsed_xml.getroot()

            generator = XsdGenerator()
            schema_tree = generator.generate(root_element)

            generated_xsd_string = etree.tostring(schema_tree, pretty_print=True, xml_declaration=True,
                                                  encoding='UTF-8')

            filepath = filedialog.asksaveasfilename(
                title="Save Generated XSD",
                defaultextension=".xsd",
                filetypes=[("XSD Schema Files", "*.xsd"), ("All files", "*.*")],
                parent=self.root
            )
            if not filepath:
                self.update_status("Ready")
                return

            with open(filepath, 'wb') as f:
                f.write(generated_xsd_string)

            self.update_status("Ready")
            messagebox.showinfo("Success", f"Basic XSD schema successfully generated and saved to:\n{filepath}",
                                parent=self.root)
        except Exception as e:
            self.update_status("Error")
            messagebox.showerror("XSD Generation Failed", f"An unexpected error occurred:\n{type(e).__name__}: {e}",
                                 parent=self.root)
            traceback.print_exc()

    def _validate_with_xsd(self):
        if self.file_type != 'xml' or not self.current_loaded_filepath:
            messagebox.showerror("Error", "Please open an XML file first.", parent=self.root)
            return

        xsd_path = filedialog.askopenfilename(
            title="Open XSD Schema File",
            filetypes=[("XSD Schema Files", "*.xsd"), ("All files", "*.*")],
            parent=self.root
        )
        if not xsd_path:
            return

        try:
            self.update_status("Validating with XSD...")
            self.root.update_idletasks()

            xml_doc = etree.parse(self.current_loaded_filepath)
            xsd_doc = etree.parse(xsd_path)
            xml_schema = etree.XMLSchema(xsd_doc)

            xml_schema.assertValid(xml_doc)

            self.update_status("Ready")
            messagebox.showinfo("Validation Successful", "The XML document is valid against the provided XSD schema.",
                                parent=self.root)

        except etree.DocumentInvalid as e:
            self.update_status("Validation failed")
            error_message = "XML validation failed.\n\nErrors:\n"
            for i, error in enumerate(e.error_log):
                if i >= 5:
                    error_message += "\n(and more...)"
                    break
                error_message += f"- Line {error.line}, Col {error.column}: {error.message}\n"
            messagebox.showerror("Validation Failed", error_message, parent=self.root)
        except Exception as e:
            self.update_status("Error during validation")
            messagebox.showerror("Validation Error", f"An unexpected error occurred:\n{e}", parent=self.root)
            traceback.print_exc()

    def _reset_ui_for_new_file(self):
        self.update_status("Resetting UI...", show_progress=False)
        self.file_type = None
        self.current_loaded_filepath = ""
        self.filename_display_var.set("No file loaded.")
        self.selected_table_var.set('')
        self.tables_combobox.set('')
        self.tables_combobox['values'] = []
        self.tables_combobox.config(state="disabled")
        self.table_combobox_map.clear()
        self.potential_tables.clear()
        self.clear_xml_treeview()

        while len(self.content_notebook.tabs()) > 1:
            self.content_notebook.forget(1)
        self.open_table_tabs.clear()
        self.content_notebook.select(0)

        self.xml_tree_root = None
        self.selected_element_for_context_menu = None
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.table_data_cache.clear()

        try:
            self.paned_window.sashpos(0, 300)
        except tk.TclError:
            pass
        self.tree_frame_outer.pack_forget()
        self.filemenu.entryconfig("Close File", state="disabled")
        self.filemenu.entryconfig("Save Table as CSV...", state="disabled")
        self.filemenu.entryconfig("Save XML As...", state="disabled")
        self.editmenu.entryconfig("Go to Row...", state="disabled")
        self.editmenu.entryconfig("Find...", state="disabled")
        self.editmenu.entryconfig("Batch Operations...", state="disabled")
        self.editmenu.entryconfig("Resize Columns", state="disabled")
        self.editmenu.entryconfig("Reorder Tabs...", state="disabled")
        self.editmenu.entryconfig("Undo", state="disabled")
        self.editmenu.entryconfig("Redo", state="disabled")
        self.context_menu.entryconfig("Export Node as CSV...", state="disabled")
        self.utilsmenu.entryconfig("Query Designer...", state="disabled")
        self.utilsmenu.entryconfig("Transactional Data Check...", state="disabled")
        self.root.update_idletasks()

    def open_xml_file_threaded(self):
        filepath = filedialog.askopenfilename(title="Open XML File",
                                              filetypes=(("XML files", "*.xml"), ("All files", "*.*")),
                                              parent=self.root)
        if not filepath:
            return
        self._reset_ui_for_new_file()
        self.file_type = 'xml'
        self.current_loaded_filepath = filepath
        self.filemenu.entryconfig("Open XML...", state="disabled")
        self.filemenu.entryconfig("Open CSV...", state="disabled")
        threading.Thread(target=self._parse_and_populate_worker, args=(filepath,), daemon=True).start()

    def open_csv_file_threaded(self):
        filepath = filedialog.askopenfilename(title="Open CSV File",
                                              filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
                                              parent=self.root)
        if not filepath:
            return
        self._reset_ui_for_new_file()
        self.file_type = 'csv'
        self.current_loaded_filepath = filepath
        self.filemenu.entryconfig("Open XML...", state="disabled")
        self.filemenu.entryconfig("Open CSV...", state="disabled")
        threading.Thread(target=self._parse_csv_and_populate_worker, args=(filepath,), daemon=True).start()

    def _parse_and_populate_worker(self, filepath):
        try:
            if os.path.getsize(filepath) == 0:
                self.root.after(0, self._finish_loading_error, "The selected XML file is empty.")
                return
            self.root.after(0, self.update_status, f"Loading {os.path.basename(filepath)}...", True, 0)
            total_size = os.path.getsize(filepath)
            bytes_read = 0
            parser = ET.XMLParser(target=ET.TreeBuilder())
            with open(filepath, 'rb') as f:
                while True:
                    chunk = f.read(CHUNK_SIZE)
                    if not chunk:
                        break
                    parser.feed(chunk)
                    bytes_read += len(chunk)
                    progress = (bytes_read / total_size) * 100 if total_size > 0 else 100
                    if bytes_read % (CHUNK_SIZE * 5) == 0:
                        self.root.after(0, self.update_status,
                                        f"Parsing {os.path.basename(filepath)} ({int(progress)}%)...", True, progress)
            self.xml_tree_root = parser.close()
            self.root.after(0, self._finish_loading_success, filepath)
        except ET.ParseError as e:
            self.xml_tree_root = None
            self.root.after(0, self._finish_loading_error, f"XML Parse Error: {str(e)[:200]}")
        except Exception as e:
            self.xml_tree_root = None
            error_type = type(e).__name__
            self.root.after(0, self._finish_loading_error, f"{error_type}: {str(e)[:200]}")
        finally:
            self.root.after(0, lambda: self.filemenu.entryconfig("Open XML...", state="normal"))
            self.root.after(0, lambda: self.filemenu.entryconfig("Open CSV...", state="normal"))

    def _parse_csv_and_populate_worker(self, filepath):
        try:
            total_size = os.path.getsize(filepath)
            if total_size == 0:
                self.root.after(0, self._finish_loading_error, "The selected CSV file is empty.")
                return

            self.root.after(0, self.update_status, f"Loading {os.path.basename(filepath)}...", True, 0)

            with open(filepath, 'r', encoding='utf-8-sig') as f:
                try:
                    dialect = csv.Sniffer().sniff(f.read(2048))
                    self.csv_delimiter = dialect.delimiter
                    messagebox.showinfo("CSV Delimiter Detected",
                                        f"Detected '{self.csv_delimiter}' as the CSV delimiter.", parent=self.root)
                except csv.Error:
                    self.csv_delimiter = ','
                f.seek(0)

                reader = csv.reader(f, delimiter=self.csv_delimiter)
                headers = next(reader)
                data = []
                for row in reader:
                    data.append({h: v for h, v in zip(headers, row)})

            internal_key = os.path.basename(filepath)
            self.potential_tables[internal_key] = {
                "columns": headers,
                "display_name_candidate": internal_key
            }
            self.table_data_cache[internal_key] = data
            self.root.after(0, self._finish_loading_success, filepath)

        except Exception as e:
            self.root.after(0, self._finish_loading_error, f"CSV Load Error: {str(e)[:200]}")
        finally:
            self.root.after(0, lambda: self.filemenu.entryconfig("Open XML...", state="normal"))
            self.root.after(0, lambda: self.filemenu.entryconfig("Open CSV...", state="normal"))

    def _finish_loading_success(self, loaded_filepath):
        self.current_loaded_filepath = loaded_filepath
        self.filename_display_var.set(os.path.basename(loaded_filepath))
        self.editmenu.entryconfig("Find...", state="normal")
        self.filemenu.entryconfig("Close File", state="normal")

        if self.file_type == 'xml':
            self.filemenu.entryconfig("Save XML As...", state="normal")
            self.tree_frame_outer.pack(side=tk.TOP, expand=True, fill='both')
            self.update_status(f"Populating tree for {os.path.basename(loaded_filepath)}...", True, 50)
            self.populate_main_xml_treeview()
            self.update_status("Discovering tables...", True, 85)
            self.discover_potential_tables()
        elif self.file_type == 'csv':
            self.tree_frame_outer.pack_forget()
            try:
                self.paned_window.sashpos(0, 0)
            except tk.TclError:
                pass
            self.update_status("Processing CSV data...", True, 90)

        self.populate_table_combobox()

        if self.potential_tables:
            self.utilsmenu.entryconfig("Query Designer...", state="normal")
            self.utilsmenu.entryconfig("Transactional Data Check...", state="normal")
            if self.file_type == 'csv':
                if self.tables_combobox['values']:
                    first_table = self.tables_combobox['values'][0]
                    self.selected_table_var.set(first_table)
                    self.on_table_combobox_select()

        self.root.update_idletasks()
        self.update_status(f"Loaded: {os.path.basename(loaded_filepath)}", show_progress=False)

    def _finish_loading_error(self, error_message):
        self.filename_display_var.set("Load Error.")
        self.current_loaded_filepath = ""
        self.update_status(f"Error: {error_message[:100]}...", show_progress=False)
        messagebox.showerror("Load Error", error_message, parent=self.root)
        self.tables_combobox.set('')
        self.tables_combobox['values'] = []
        self.tables_combobox.config(state="disabled")
        self.filemenu.entryconfig("Save XML As...", state="disabled")
        self.editmenu.entryconfig("Find...", state="disabled")
        self.utilsmenu.entryconfig("Query Designer...", state="disabled")
        self.utilsmenu.entryconfig("Transactional Data Check...", state="disabled")
        self.root.update_idletasks()

    def clear_xml_treeview(self):
        children_to_delete = list(self.tree.get_children(""))
        for item in children_to_delete:
            try:
                self.tree.delete(item)
            except tk.TclError:
                pass
        self.tree_item_to_element.clear()

    def close_current_file(self):
        self._reset_ui_for_new_file()

    def populate_main_xml_treeview(self):
        self.clear_xml_treeview()
        if self.xml_tree_root is None:
            self.update_status("XML root not found.", False)
            return
        self.add_node_to_tree(self.xml_tree_root, "")
        self.update_status("Tree populated.", False)

    def add_node_to_tree(self, element, parent_item_id):
        node_text = f"<{element.tag}>"
        if element.attrib:
            first_attr_key, first_attr_val = list(element.attrib.items())[0]
            val_str = str(first_attr_val)
            attr_preview = f'{first_attr_key}="{val_str[:15]}"'
            if len(val_str) > 15 or len(element.attrib) > 1:
                attr_preview += "..."
            node_text += f" [{attr_preview}]"

        child_count = len(element)
        if child_count > 0:
            node_text += f" ({child_count})"

        item_id = self.tree.insert(parent_item_id, 'end', text=node_text, open=False)
        self.tree_item_to_element[item_id] = element

        if child_count > 0:
            self.tree.insert(item_id, 'end', text="Loading...", iid=f"{item_id}_dummy")

    def on_tree_node_expand(self, event):
        item_id = self.tree.focus()
        dummy_node_id = f"{item_id}_dummy"

        if not self.tree.exists(dummy_node_id):
            return

        self.tree.delete(dummy_node_id)
        element = self.tree_item_to_element.get(item_id)
        if element is None:
            return

        for child_element in element:
            self.add_node_to_tree(child_element, item_id)

    def discover_potential_tables(self):
        self.potential_tables.clear()
        if self.xml_tree_root is None:
            return
        for parent_element in self.xml_tree_root.iter():
            if len(parent_element) < MIN_ROWS_FOR_TABLE:
                continue
            child_tags = [child.tag for child in parent_element if child.tag is not None]
            if not child_tags:
                continue
            tag_counts = Counter(child_tags)
            if not tag_counts:
                continue
            most_common_tag, count = tag_counts.most_common(1)[0]
            if count >= MIN_ROWS_FOR_TABLE and (count / len(parent_element)) >= MIN_PERCENT_SIMILAR:
                row_elements = [child for child in parent_element if child.tag == most_common_tag]
                if not row_elements:
                    continue
                column_headers = set()
                for row_el in row_elements[:50]:
                    for col_el in row_el:
                        if col_el.tag is not None:
                            column_headers.add(col_el.tag)
                if column_headers:
                    internal_key = f"{parent_element.tag}_rows_{most_common_tag}_id{id(parent_element)}"

                    parent_tag_lower = parent_element.tag.lower()
                    child_tag_lower = most_common_tag.lower()
                    is_plural_of_child = parent_tag_lower.endswith('s') and parent_tag_lower.rstrip(
                        's') == child_tag_lower

                    if parent_tag_lower == child_tag_lower or is_plural_of_child:
                        display_name_candidate = parent_element.tag.capitalize()
                    elif parent_element.tag != self.xml_tree_root.tag:
                        display_name_candidate = f"{parent_element.tag.capitalize()}/{most_common_tag.capitalize()}"
                    else:
                        display_name_candidate = most_common_tag.capitalize()

                    self.potential_tables[internal_key] = {
                        "parent_element": parent_element, "row_tag": most_common_tag,
                        "columns": sorted(list(column_headers)), "display_name_candidate": display_name_candidate,
                        "original_parent_tag": parent_element.tag
                    }

    def populate_table_combobox(self):
        self.tables_combobox.set('')
        self.tables_combobox['values'] = []
        self.table_combobox_map.clear()
        if not self.potential_tables:
            self.tables_combobox.config(state="disabled")
            return
        sorted_internal_keys = sorted(
            self.potential_tables.keys(),
            key=lambda k: (self.potential_tables[k]['display_name_candidate'].lower(), k)
        )
        combobox_values = []
        temp_label_counts = {}
        for internal_key in sorted_internal_keys:
            info = self.potential_tables[internal_key]
            base_label = info['display_name_candidate']
            final_label = base_label
            if base_label in temp_label_counts:
                temp_label_counts[base_label] += 1
                final_label = f"{base_label} ({temp_label_counts[base_label]})"
            else:
                temp_label_counts[base_label] = 1
            combobox_values.append(final_label)
            self.table_combobox_map[final_label] = internal_key
        self.tables_combobox['values'] = combobox_values
        if combobox_values:
            self.tables_combobox.config(state="readonly")
        else:
            self.tables_combobox.config(state="disabled")

    def show_find_dialog(self):
        try:
            if not self.potential_tables:
                messagebox.showwarning("Warning", "No tables available to search.", parent=self.root)
                return
            find_dialog = tk.Toplevel(self.root)
            find_dialog.title("Find and Replace")
            find_dialog.geometry("600x450")
            find_dialog.transient(self.root)
            find_dialog.grab_set()
            find_dialog.columnconfigure(1, weight=1)

            # --- Widgets ---
            ttk.Label(find_dialog, text="Table:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            table_var = tk.StringVar()
            table_combobox = ttk.Combobox(find_dialog, textvariable=table_var, state="readonly")
            table_combobox.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

            ttk.Label(find_dialog, text="Field:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
            field_var = tk.StringVar()
            field_combobox = ttk.Combobox(find_dialog, textvariable=field_var, state="readonly")
            field_combobox.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

            ttk.Label(find_dialog, text="Find what:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
            search_var = tk.StringVar()
            search_entry = ttk.Entry(find_dialog, textvariable=search_var)
            search_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
            search_entry.focus_set()

            ttk.Label(find_dialog, text="Replace with:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
            replace_var = tk.StringVar()
            replace_entry = ttk.Entry(find_dialog, textvariable=replace_var)
            replace_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

            match_case_var = tk.BooleanVar()
            ttk.Checkbutton(find_dialog, text="Match case", variable=match_case_var).grid(row=4, column=1, padx=5,
                                                                                          pady=5, sticky="w")

            results_frame = ttk.Labelframe(find_dialog, text="Find Results", padding=5)
            results_frame.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
            find_dialog.rowconfigure(5, weight=1)
            results_tree = ttk.Treeview(results_frame, columns=("Row", "Column", "Value"), show="headings", height=5)
            results_tree.heading("Row", text="Row #")
            results_tree.heading("Column", text="Column")
            results_tree.heading("Value", text="Value")
            results_tree.column("Row", width=60, anchor="e")
            results_tree.column("Column", width=120)
            results_tree.column("Value", width=250)
            results_tree.pack(side="left", fill="both", expand=True)
            results_vsb = ttk.Scrollbar(results_frame, orient="vertical", command=results_tree.yview)
            results_vsb.pack(side="right", fill="y")
            results_tree.configure(yscrollcommand=results_vsb.set)

            # --- Buttons ---
            button_frame = ttk.Frame(find_dialog)
            button_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky="e")
            find_button = ttk.Button(button_frame, text="Find All", command=lambda: perform_search())
            find_button.pack(side=tk.LEFT, padx=5)
            replace_button = ttk.Button(button_frame, text="Replace", state="disabled")
            replace_button.pack(side=tk.LEFT, padx=5)
            replace_all_button = ttk.Button(button_frame, text="Replace All", state="disabled")
            replace_all_button.pack(side=tk.LEFT, padx=5)

            # --- State and Logic ---
            matches = []

            def update_fields(*args):
                selected_table = table_var.get()
                internal_key = self.table_combobox_map.get(selected_table)
                if internal_key:
                    columns = self.potential_tables[internal_key]["columns"]
                    field_combobox['values'] = ["(Any Field)"] + columns
                    field_var.set("(Any Field)")

            table_names = list(self.table_combobox_map.keys())
            table_combobox['values'] = table_names
            current_tab = self.get_current_table_tab()
            if current_tab:
                current_name = next(
                    (name for name, key in self.table_combobox_map.items() if key == current_tab.internal_key), None)
                if current_name: table_var.set(current_name)
            elif table_names:
                table_var.set(table_names[0])
            table_var.trace("w", update_fields)
            update_fields()

            def perform_search():
                nonlocal matches
                for item in results_tree.get_children(): results_tree.delete(item)
                matches = []
                replace_button.config(state="disabled")
                replace_all_button.config(state="disabled")

                selected_table_name = table_var.get()
                internal_key = self.table_combobox_map.get(selected_table_name)
                if not internal_key: return

                search_value = search_var.get()
                if not search_value: return

                if internal_key not in self.table_data_cache:
                    self.selected_table_var.set(selected_table_name)
                    self.on_table_combobox_select()
                    self.root.after(100, perform_search)
                    return

                full_data = self.table_data_cache.get(internal_key, [])
                field_to_search = field_var.get()
                match_case = match_case_var.get()
                search_term = search_value if match_case else search_value.lower()
                columns_to_search = [field_to_search] if field_to_search != "(Any Field)" else \
                self.potential_tables[internal_key]["columns"]

                for i, row_data in enumerate(full_data):
                    for col in columns_to_search:
                        cell_text = str(row_data.get(col, ""))
                        compare_text = cell_text if match_case else cell_text.lower()
                        if search_term in compare_text:
                            matches.append((i, col))
                            results_tree.insert("", "end", values=(i + 1, col, cell_text))

                if matches:
                    replace_button.config(state="normal")
                    replace_all_button.config(state="normal")
                    results_tree.selection_set(results_tree.get_children()[0])
                    results_tree.focus(results_tree.get_children()[0])

            def do_replace():
                selection = results_tree.selection()
                if not selection: return

                selected_item_id = selection[0]
                item_index = results_tree.index(selected_item_id)
                data_index, column = matches[item_index]

                internal_key = self.table_combobox_map.get(table_var.get())
                old_value = self.table_data_cache[internal_key][data_index][column]
                new_value = old_value.replace(search_var.get(), replace_var.get())

                self.table_data_cache[internal_key][data_index][column] = new_value
                if self.file_type == 'xml':
                    element = self.table_data_cache[internal_key][data_index]['_element']
                    element.find(column).text = new_value

                self.push_undo({"action": "edit", "column": column, "old_value": old_value, "new_value": new_value,
                                "internal_key": internal_key, "original_index": data_index})

                # Refresh the main table view
                target_tab = self.get_current_table_tab()
                if target_tab and target_tab.internal_key == internal_key:
                    target_tab._apply_filter_and_sort()

                results_tree.item(selected_item_id, values=(data_index + 1, column, new_value))
                next_index = item_index + 1
                if next_index < len(results_tree.get_children()):
                    next_item = results_tree.get_children()[next_index]
                    results_tree.selection_set(next_item)
                    results_tree.focus(next_item)
                    results_tree.see(next_item)
                else:
                    replace_button.config(state="disabled")

            def do_replace_all():
                if not matches: return

                internal_key = self.table_combobox_map.get(table_var.get())
                search_term = search_var.get()
                replace_term = replace_var.get()

                changes = []
                for data_index, column in matches:
                    old_value = self.table_data_cache[internal_key][data_index][column]
                    if search_term in old_value:
                        new_value = old_value.replace(search_term, replace_term)
                        self.table_data_cache[internal_key][data_index][column] = new_value
                        if self.file_type == 'xml':
                            element = self.table_data_cache[internal_key][data_index]['_element']
                            element.find(column).text = new_value
                        changes.append({"original_index": data_index, "column": column, "old_value": old_value,
                                        "new_value": new_value})

                if changes:
                    self.push_undo({"action": "batch_replace", "changes": changes, "internal_key": internal_key})
                    target_tab = self.get_current_table_tab()
                    if target_tab and target_tab.internal_key == internal_key:
                        target_tab._apply_filter_and_sort()
                    messagebox.showinfo("Replace All", f"Replaced {len(changes)} occurrence(s).", parent=find_dialog)
                    find_dialog.destroy()

            replace_button.config(command=do_replace)
            replace_all_button.config(command=do_replace_all)
            search_entry.bind("<Return>", lambda e: perform_search())

        except Exception as e:
            messagebox.showerror("Error", f"Failed to show find dialog: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def on_table_combobox_select(self, event=None):
        selected_display_name = self.selected_table_var.get()
        if not selected_display_name:
            return
        internal_key = self.table_combobox_map.get(selected_display_name)
        if not internal_key:
            return

        if internal_key in self.open_table_tabs:
            self.content_notebook.select(self.open_table_tabs[internal_key])
            return

        if len(self.open_table_tabs) >= 5:
            proceed = messagebox.askokcancel(
                "Memory Warning",
                "You have 5 or more data tabs open.\nOpening more may impact application performance.\n\nDo you want to continue?",
                parent=self.root
            )
            if not proceed:
                self.selected_table_var.set("")
                return

        try:
            new_tab = TableViewTab(self.content_notebook, self, internal_key)
            tab_text = selected_display_name
            if len(tab_text) > 25:
                tab_text = tab_text[:22] + "..."
            self.content_notebook.add(new_tab, text=tab_text)
            tab_id = self.content_notebook.tabs()[-1]
            self.open_table_tabs[internal_key] = tab_id
            self.content_notebook.select(tab_id)

            self.filemenu.entryconfig("Save Table as CSV...", state="normal")
            self.editmenu.entryconfig("Find...", state="normal")
            self.editmenu.entryconfig("Batch Operations...", state="normal")
            self.editmenu.entryconfig("Go to Row...", state="normal")
            self.editmenu.entryconfig("Resize Columns", state="normal")
            if len(self.content_notebook.tabs()) > 2:
                self.editmenu.entryconfig("Reorder Tabs...", state="normal")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to open table tab: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def on_xml_tree_node_select(self, event=None):
        current_tab = self.get_current_table_tab()
        if current_tab and hasattr(current_tab, 'active_cell_editor') and current_tab.active_cell_editor:
            current_tab._finish_cell_edit()

        selected_item_id = self.tree.focus()
        if not selected_item_id:
            return
        element = self.tree_item_to_element.get(selected_item_id)
        if element is None:
            return

        self.selected_element_for_context_menu = element
        self._update_context_menu_state(element)

        # Check if the clicked element is a parent of a detected table
        is_table_node = False
        for internal_key, table_info in self.potential_tables.items():
            if element is table_info.get("parent_element"):
                # Find the display name associated with this table's internal key
                display_name = next((name for name, key in self.table_combobox_map.items() if key == internal_key),
                                    None)
                if display_name:
                    # Set the combobox variable and call the selection handler to open the tab
                    self.selected_table_var.set(display_name)
                    self.on_table_combobox_select()
                    is_table_node = True
                    break

        # If the node was not a table, show the default details view
        if not is_table_node:
            self.content_notebook.select(0)
            self.update_node_detail_panel(element)

    def update_node_detail_panel(self, element):
        try:
            for item in self.detail_tree.get_children():
                self.detail_tree.delete(item)
            if element is None:
                return
            if len(element) > 0:
                for child_element in element:
                    child_tag_display = f"<{child_element.tag}>"
                    child_value_display = ""
                    if child_element.text and child_element.text.strip():
                        child_value_display = child_element.text.strip()
                    elif len(child_element) > 0:
                        child_value_display = f"({len(child_element)} sub-elements)"
                    elif child_element.attrib:
                        child_value_display = "(empty, has attributes)"
                    else:
                        child_value_display = "(empty)"
                    child_item_id = self.detail_tree.insert("", "end", values=(child_tag_display, child_value_display))
                    if child_element.attrib:
                        for k, v in child_element.attrib.items():
                            self.detail_tree.insert(child_item_id, 'end', values=(f"  @{k}", v))
            else:
                self.detail_tree.insert("", "end", values=("Tag", f"<{element.tag}>"))
                element_text = "(empty)"
                if element.text and element.text.strip():
                    element_text = element.text.strip()
                self.detail_tree.insert("", "end", values=("Text", element_text))
                if element.attrib:
                    attr_node_id = self.detail_tree.insert("", "end",
                                                           values=("Attributes", f"({len(element.attrib)} items)"))
                    for k, v in element.attrib.items():
                        self.detail_tree.insert(attr_node_id, "end", values=(f"  @{k}", v))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update detail panel: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def undo_action(self):
        if not self.undo_stack: return
        action = self.undo_stack.pop()
        internal_key = action["internal_key"]

        if action["action"] == "edit":
            original_index, column, old_value = action["original_index"], action["column"], action["old_value"]
            self.table_data_cache[internal_key][original_index][column] = old_value
            if self.file_type == 'xml':
                element = self.table_data_cache[internal_key][original_index]['_element']
                element.find(column).text = old_value
        elif action["action"] == "batch_replace" or action["action"] == "batch_update":
            for change in action["changes"]:
                original_index, column, old_value = change["original_index"], change["column"], change["old_value"]
                self.table_data_cache[internal_key][original_index][column] = old_value
                if self.file_type == 'xml':
                    self.table_data_cache[internal_key][original_index]['_element'].find(column).text = old_value
        elif action["action"] == "batch_delete":
            parent_element = self.potential_tables[internal_key].get(
                "parent_element") if self.file_type == 'xml' else None
            for item in sorted(action["deleted_rows"], key=lambda x: x['index']):
                self.table_data_cache[internal_key].insert(item['index'], item['data'])
                if parent_element is not None:
                    element_to_insert = item['data'].get("_element")
                    if element_to_insert is not None:
                        parent_element.insert(item['index'], element_to_insert)

        self.redo_stack.append(action)
        if not self.undo_stack: self.editmenu.entryconfig("Undo", state="disabled")
        self.editmenu.entryconfig("Redo", state="normal")

        if internal_key in self.open_table_tabs:
            tab_id = self.open_table_tabs[internal_key]
            tab_widget = self.root.nametowidget(tab_id)
            if isinstance(tab_widget, TableViewTab):
                tab_widget._apply_filter_and_sort()

    def redo_action(self):
        if not self.redo_stack: return
        action = self.redo_stack.pop()
        internal_key = action["internal_key"]

        if action["action"] == "edit":
            original_index, column, new_value = action["original_index"], action["column"], action["new_value"]
            self.table_data_cache[internal_key][original_index][column] = new_value
            if self.file_type == 'xml':
                self.table_data_cache[internal_key][original_index]['_element'].find(column).text = new_value
        elif action["action"] == "batch_replace" or action["action"] == "batch_update":
            for change in action["changes"]:
                original_index, column, new_value = change["original_index"], change["column"], change["new_value"]
                self.table_data_cache[internal_key][original_index][column] = new_value
                if self.file_type == 'xml':
                    self.table_data_cache[internal_key][original_index]['_element'].find(column).text = new_value
        elif action["action"] == "batch_delete":
            for item in sorted(action["deleted_rows"], key=lambda x: x['index'], reverse=True):
                self.table_data_cache[internal_key].pop(item['index'])

        self.undo_stack.append(action)
        if not self.redo_stack: self.editmenu.entryconfig("Redo", state="disabled")
        self.editmenu.entryconfig("Undo", state="normal")

        if internal_key in self.open_table_tabs:
            tab_id = self.open_table_tabs[internal_key]
            tab_widget = self.root.nametowidget(tab_id)
            if isinstance(tab_widget, TableViewTab):
                tab_widget._apply_filter_and_sort()

    def _update_context_menu_state(self, element):
        try:
            if self.file_type == 'xml' and element is not None and len(element) > 0 and element[0].tag is not None:
                if all(child.tag == element[0].tag for child in element):
                    self.context_menu.entryconfig("Export Node as CSV...", state="normal")
                else:
                    self.context_menu.entryconfig("Export Node as CSV...", state="disabled")
            else:
                self.context_menu.entryconfig("Export Node as CSV...", state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update context menu: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def show_context_menu(self, event):
        try:
            if self.file_type != 'xml': return
            current_tab = self.get_current_table_tab()
            if current_tab and hasattr(current_tab, 'active_cell_editor') and current_tab.active_cell_editor:
                current_tab._finish_cell_edit()
            item_id = self.tree.identify_row(event.y)
            if item_id:
                if self.tree.selection() != (item_id,):
                    self.tree.selection_set(item_id)
                    self.tree.focus(item_id)
                self._update_context_menu_state(self.tree_item_to_element.get(item_id))
                self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to show context menu: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def save_xml_as(self):
        try:
            if self.file_type != 'xml' or self.xml_tree_root is None: return
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xml",
                filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
                title="Save XML File",
                parent=self.root
            )
            if not filepath: return
            self.update_status(f"Saving XML to {os.path.basename(filepath)}...", show_progress=True, progress_value=0)
            try:
                tree = ET.ElementTree(self.xml_tree_root)
                ET.indent(tree)
                tree.write(filepath, encoding='utf-8', xml_declaration=True)
                self.update_status(f"Success: saved to {os.path.basename(filepath)}", show_progress=False)
                messagebox.showinfo("Success", f"Successfully saved to {filepath}", parent=self.root)
                self.current_loaded_filepath = filepath
                self.filename_display_var.set(os.path.basename(filepath))
            except Exception as e:
                self.update_status(f"Error saving XML: {str(e)[:50]}...", show_progress=False)
                messagebox.showerror("Error", f"Failed to save XML: {e}\n{traceback.format_exc()[:200]}",
                                     parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save XML: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def _save_current_table_as_csv(self):
        current_tab = self.get_current_table_tab()
        if not current_tab:
            messagebox.showerror("Export Error", "No table tab is currently selected.", parent=self.root)
            return

        headers, rows = current_tab.get_current_table_data()
        if not headers or not rows:
            messagebox.showerror("Export Error", "The selected table has no data to export.", parent=self.root)
            return

        initial_filename = self.content_notebook.tab(current_tab, "text").replace("...", "")
        filepath = filedialog.asksaveasfilename(
            title="Save Current Table as CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"{initial_filename}.csv",
            parent=self.root
        )
        if not filepath: return

        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=self.csv_delimiter)
                writer.writerow(headers)
                writer.writerows(rows)
            messagebox.showinfo("Success", f"Table '{initial_filename}' successfully saved as CSV.", parent=self.root)
        except Exception as e:
            messagebox.showerror("Save Error", f"An error occurred while saving the CSV file:\n{e}", parent=self.root)

    def export_node_as_csv(self):
        try:
            if self.file_type != 'xml': return
            selected_element = self.selected_element_for_context_menu
            if selected_element is None or len(selected_element) == 0: return
            child_elements = list(selected_element)
            first_child_tag = child_elements[0].tag
            if not all(child.tag == first_child_tag for child in child_elements):
                messagebox.showwarning("Warning", "All child elements must have the same tag.", parent=self.root)
                return
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Node as CSV",
                parent=self.root
            )
            if not filepath: return
            self.update_status(f"Exporting to {os.path.basename(filepath)}...", show_progress=True, progress_value=0)
            try:
                headers = set()
                for child in child_elements:
                    headers.update(child.attrib.keys())
                    for sub_elem in child:
                        if sub_elem.tag is not None:
                            headers.add(sub_elem.tag)
                if not headers and any(child.text and child.text.strip() for child in child_elements):
                    headers.add(f"{first_child_tag}_text")
                if not headers:
                    messagebox.showerror("Error", "No valid headers found for CSV export.", parent=self.root)
                    self.update_status("Export failed: no headers.", show_progress=False)
                    return
                sorted_headers = sorted(list(headers))
                with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=sorted_headers, extrasaction='ignore')
                    writer.writeheader()
                    total_children = len(child_elements)
                    for i, child in enumerate(child_elements):
                        row_data = {k: v for k, v in child.attrib.items() if k in sorted_headers}
                        for sub_elem in child:
                            if sub_elem.tag in sorted_headers:
                                row_data[sub_elem.tag] = (sub_elem.text or "").strip()
                        if f"{first_child_tag}_text" in sorted_headers and child.text and child.text.strip() and not child.attrib and len(
                                child) == 0:
                            row_data[f"{first_child_tag}_text"] = child.text.strip()
                        writer.writerow(row_data)
                        if (i + 1) % 100 == 0 or (i + 1) == total_children:
                            progress = ((i + 1) / total_children) * 100 if total_children > 0 else 100
                            self.update_status(f"Exporting node ({int(progress)}%)...", show_progress=True,
                                               progress_value=progress)
                self.update_status(f"Success: exported to {os.path.basename(filepath)}", show_progress=False)
                messagebox.showinfo("Success", f"Successfully exported node to {filepath}", parent=self.root)
            except Exception as e:
                self.update_status(f"Error exporting CSV: {str(e)[:50]}...", show_progress=False)
                messagebox.showerror("Error", f"Failed to export CSV: {e}\n{traceback.format_exc()[:200]}",
                                     parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export CSV: {e}\n{traceback.format_exc()[:200]}",
                                 parent=self.root)

    def _handle_goto_row(self):
        current_tab = self.get_current_table_tab()
        if current_tab:
            current_tab.handle_goto_row()

if __name__ == "__main__":
    root = tk.Tk()
    app = XMLNotepad(root)
    root.app = app
    root.mainloop()