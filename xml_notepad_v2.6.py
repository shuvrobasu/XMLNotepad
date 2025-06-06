# ------------------------------#
# XML Notepad ver 2.8 6th June #
# ------------------------------#
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkFont
import xml.etree.ElementTree as ET
import csv
import threading
from collections import Counter, deque, defaultdict
import traceback
import os
import json
import re

CHUNK_SIZE = 1024 * 1024
MIN_ROWS_FOR_TABLE = 3
MIN_PERCENT_SIMILAR = 0.6
UNDO_STACK_SIZE = 20

# ------------------------------#
# XML Notepad ver 3.4 6th June #
# ------------------------------#
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkFont
import xml.etree.ElementTree as ET
import csv
import threading
from collections import Counter, deque, defaultdict
import traceback
import os
import json
import re


class QueryDesigner(tk.Toplevel):
    """
    An advanced query designer with Intellisense for the manual query editor.
    """

    def __init__(self, parent, potential_tables, table_combobox_map, source_xml_path):
        super().__init__(parent)
        self.title("Advanced Query Designer")
        self.geometry("1200x800")
        self.transient(parent)
        self.grab_set()

        self.potential_tables = potential_tables
        self.table_combobox_map = table_combobox_map
        self.table_names = sorted(list(self.table_combobox_map.keys()))
        self.table_names_with_blank = [""] + self.table_names
        self.source_xml_path = source_xml_path

        self.join_conditions = []
        self.filter_conditions = []

        self.table1_var = tk.StringVar()
        self.table2_var = tk.StringVar()
        self.field1_var = tk.StringVar()
        self.field2_var = tk.StringVar()

        self.filter_field_var = tk.StringVar()
        self.filter_op_var = tk.StringVar(value="CONTAINS")
        self.filter_value_var = tk.StringVar()

        self.query_type_var = tk.StringVar(value="INNER")
        self.limit_enabled_var = tk.BooleanVar(value=False)
        self.limit_value_var = tk.IntVar(value=100)
        self.manual_edit_mode = tk.BooleanVar(value=False)

        self.current_results_data = []
        self.results_sort_col = None
        self.results_sort_asc = True

        # New variable to hold the Intellisense popup widget
        self.intellisense_popup = None

        self._setup_ui()
        self._populate_initial_dropdowns()
        self._on_table_select()

        self.bind("<Control-w>", lambda e: self._resize_query_results_columns())
        self.bind("<Control-e>", lambda e: self._export_results())

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
        self.config(menu=self.menubar)

        main_paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        config_notebook = ttk.Notebook(main_paned)
        main_paned.add(config_notebook, weight=2)

        visual_designer_frame = ttk.Frame(config_notebook, padding=10)
        config_notebook.add(visual_designer_frame, text="Visual Designer")
        self._create_visual_designer_widgets(visual_designer_frame)

        manual_query_frame = ttk.Frame(config_notebook, padding=10)
        config_notebook.add(manual_query_frame, text="Query View")
        self._create_manual_query_widgets(manual_query_frame)

        results_pane = ttk.Frame(main_paned)
        main_paned.add(results_pane, weight=3)
        self._create_results_widgets(results_pane)

        self.table1_combo.bind("<<ComboboxSelected>>", self._on_table_select)
        self.table2_combo.bind("<<ComboboxSelected>>", self._on_table_select)
        self.results_tree.bind("<Button-1>", self._on_results_click)

        for var in [self.query_type_var, self.limit_enabled_var, self.limit_value_var]:
            var.trace_add("write", lambda *args: self._update_query_view())

    def _create_visual_designer_widgets(self, parent):
        parent.grid_columnconfigure(1, weight=1)
        table_select_frame = ttk.Labelframe(parent, text="1. Select Tables", padding=10)
        table_select_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Label(table_select_frame, text="Left Table (T1):").pack(anchor="w")
        self.table1_combo = ttk.Combobox(table_select_frame, textvariable=self.table1_var, state="readonly", width=30)
        self.table1_combo.pack(pady=2, fill=tk.X, expand=True)
        ttk.Label(table_select_frame, text="Right Table (T2):").pack(anchor="w", pady=(5, 0))
        self.table2_combo = ttk.Combobox(table_select_frame, textvariable=self.table2_var, state="readonly", width=30)
        self.table2_combo.pack(pady=2, fill=tk.X, expand=True)

        self.conditions_frame = ttk.Labelframe(parent, text="2. Define Conditions", padding=10)
        self.conditions_frame.grid(row=0, column=1, rowspan=2, padx=5, pady=5, sticky="nsew")
        self.conditions_frame.grid_columnconfigure(0, weight=1)
        self.conditions_frame.grid_rowconfigure(1, weight=1)

        self.join_controls_frame = ttk.Frame(self.conditions_frame)
        self.field1_combo = ttk.Combobox(self.join_controls_frame, textvariable=self.field1_var, state="disabled",
                                         width=15)
        self.field1_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Label(self.join_controls_frame, text="=").pack(side=tk.LEFT, padx=5)
        self.field2_combo = ttk.Combobox(self.join_controls_frame, textvariable=self.field2_var, state="disabled",
                                         width=15)
        self.field2_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.filter_controls_frame = ttk.Frame(self.conditions_frame)
        self.filter_field_combo = ttk.Combobox(self.filter_controls_frame, textvariable=self.filter_field_var,
                                               state="disabled", width=15)
        self.filter_field_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        op_values = ["CONTAINS", "NOT CONTAINS", "=", "!=", ">", "<", ">=", "<="]
        self.filter_op_combo = ttk.Combobox(self.filter_controls_frame, textvariable=self.filter_op_var,
                                            values=op_values, state="readonly", width=12)
        self.filter_op_combo.pack(side=tk.LEFT, padx=5)
        self.filter_value_entry = ttk.Entry(self.filter_controls_frame, textvariable=self.filter_value_var, width=15)
        self.filter_value_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.add_condition_button = ttk.Button(self.conditions_frame, text="Add...", command=self._add_condition)
        self.add_condition_button.grid(row=0, column=1, padx=5, sticky="e")

        self.conditions_listbox = tk.Listbox(self.conditions_frame, height=4)
        self.conditions_listbox.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=5)
        ttk.Button(self.conditions_frame, text="Remove Selected", command=self._remove_condition).grid(row=2, column=0,
                                                                                                       columnspan=2,
                                                                                                       sticky="e")

        output_frame = ttk.Labelframe(parent, text="3. Design Report Output", padding=10)
        output_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        ttk.Label(output_frame, text="Available Fields").grid(row=0, column=0, sticky="w")
        self.available_fields_lb = tk.Listbox(output_frame, selectmode=tk.EXTENDED)
        self.available_fields_lb.grid(row=1, column=0, sticky="nsew")
        shuttle_buttons = ttk.Frame(output_frame)
        shuttle_buttons.grid(row=1, column=1, padx=5)
        ttk.Button(shuttle_buttons, text=">", width=3, command=self._add_output_field).pack(pady=2)
        ttk.Button(shuttle_buttons, text=">>", width=3, command=self._add_all_output_fields).pack(pady=2)
        ttk.Button(shuttle_buttons, text="<", width=3, command=self._remove_output_field).pack(pady=2)
        ttk.Button(shuttle_buttons, text="<<", width=3, command=self._remove_all_output_fields).pack(pady=2)
        ttk.Label(output_frame, text="Selected Fields (in order)").grid(row=0, column=2, sticky="w")
        self.selected_fields_lb = tk.Listbox(output_frame, selectmode=tk.EXTENDED)
        self.selected_fields_lb.grid(row=1, column=2, sticky="nsew")
        order_buttons = ttk.Frame(output_frame)
        order_buttons.grid(row=1, column=3, padx=5)
        ttk.Button(order_buttons, text="Up", command=lambda: self._move_output_field(-1)).pack(pady=2)
        ttk.Button(order_buttons, text="Down", command=lambda: self._move_output_field(1)).pack(pady=2)
        output_frame.grid_columnconfigure(0, weight=1)
        output_frame.grid_columnconfigure(2, weight=1)
        output_frame.grid_rowconfigure(1, weight=1)

    def _create_manual_query_widgets(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        manual_actions = ttk.Frame(parent)
        manual_actions.grid(row=0, column=0, sticky="ew", pady=5)
        ttk.Button(manual_actions, text="Toggle Edit Mode", command=self._toggle_manual_edit).pack(side=tk.LEFT)
        self.validate_button = ttk.Button(manual_actions, text="Validate & Apply to Designer",
                                          command=self._parse_and_apply_manual_query, state="disabled")
        self.validate_button.pack(side=tk.LEFT, padx=10)
        self.query_view_text = tk.Text(parent, wrap=tk.WORD, state="disabled", font=("Courier New", 10))
        self.query_view_text.grid(row=1, column=0, sticky="nsew")
        query_sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.query_view_text.yview)
        query_sb.grid(row=1, column=1, sticky="ns")
        self.query_view_text.config(yscrollcommand=query_sb.set)

        # Bind the key release event for Intellisense
        self.query_view_text.bind("<KeyRelease>", self._on_key_release)

    def _create_results_widgets(self, parent):
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        action_frame = ttk.Labelframe(parent, text="4. Set Options and Run", padding=10)
        action_frame.grid(row=0, column=0, sticky="ew", pady=5)
        self.join_type_rb_inner = ttk.Radiobutton(action_frame, text="Inner Join", variable=self.query_type_var,
                                                  value="INNER")
        self.join_type_rb_inner.pack(side=tk.LEFT, padx=5)
        self.join_type_rb_anti = ttk.Radiobutton(action_frame, text="Left Anti-Join", variable=self.query_type_var,
                                                 value="ANTI")
        self.join_type_rb_anti.pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(action_frame, text="Limit:", variable=self.limit_enabled_var,
                        command=self._toggle_limit_entry).pack(side=tk.LEFT, padx=(15, 0))
        self.limit_spinbox = ttk.Spinbox(action_frame, from_=1, to=1000000, textvariable=self.limit_value_var, width=8,
                                         state="disabled")
        self.limit_spinbox.pack(side=tk.LEFT)
        action_button_frame = ttk.Frame(action_frame)
        action_button_frame.pack(side=tk.RIGHT)
        ttk.Button(action_button_frame, text="Run Query", command=self._run_query, style="Accent.TButton").pack(
            side=tk.LEFT, padx=10)
        ttk.Style().configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        results_grid_frame = ttk.Labelframe(parent, text="Results", padding=10)
        results_grid_frame.grid(row=1, column=0, sticky="nsew")
        results_grid_frame.rowconfigure(0, weight=1)
        results_grid_frame.columnconfigure(0, weight=1)
        self.results_tree = ttk.Treeview(results_grid_frame, show='headings')
        results_vsb = ttk.Scrollbar(results_grid_frame, orient="vertical", command=self.results_tree.yview)
        results_hsb = ttk.Scrollbar(results_grid_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=results_vsb.set, xscrollcommand=results_hsb.set)
        results_hsb.pack(side='bottom', fill='x')
        results_vsb.pack(side='right', fill='y')
        self.results_tree.pack(fill='both', expand=True)

    ### START INTELLISENSE METHODS ###

    def _on_key_release(self, event):
        """Handle key releases in the text widget to trigger intellisense."""
        if not self.manual_edit_mode.get():
            return

        # Always destroy any existing popup if the user types anything else
        if event.char != ":":
            self._destroy_intellisense()
            return

        # Check for T1: or T2:
        cursor_index = self.query_view_text.index(tk.INSERT)
        line, char = map(int, cursor_index.split('.'))
        if char < 3:  # Not enough characters for "T1:"
            self._destroy_intellisense()
            return

        # Get the text immediately preceding the cursor
        start_index = f"{line}.{char - 3}"
        prefix = self.query_view_text.get(start_index, cursor_index).upper()

        if prefix == "T1:":
            self._show_intellisense("T1")
        elif prefix == "T2:":
            self._show_intellisense("T2")
        else:
            self._destroy_intellisense()

    def _show_intellisense(self, alias):
        """Creates and displays the intellisense dropdown list."""
        self._destroy_intellisense()
        table_name = (self.table1_var if alias == "T1" else self.table2_var).get()
        if not table_name: return
        key = self.table_combobox_map.get(table_name)
        if not key: return
        columns = self.potential_tables[key]['columns']
        if not columns: return

        bbox = self.query_view_text.bbox(tk.INSERT)
        if not bbox: return
        x, y, _, height = bbox

        self.intellisense_popup = tk.Frame(self, relief='solid', borderwidth=1)
        # --- NEW: Store the alias for context checking ---
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
        """Handles selection from the intellisense list with context-aware validation."""
        listbox = event.widget
        selection_indices = listbox.curselection()
        if not selection_indices: return

        selected_field = listbox.get(selection_indices[0])

        # Determine the alias ("T1" or "T2") that triggered the popup
        alias = self.intellisense_popup.alias
        full_field_name = f"{alias}:  {selected_field}"

        # --- CONTEXT-AWARE VALIDATION ---
        # Only check for duplicates if the cursor is inside the SELECT clause
        if self._is_in_select_clause():
            all_text = self.query_view_text.get("1.0", "end-1c")
            select_clause_text = re.search(r"SELECT(.*?)FROM", all_text, re.S | re.I)

            if select_clause_text and full_field_name in select_clause_text.group(1):
                messagebox.showwarning("Duplicate Field",
                                       f"The field '{full_field_name}' is already in the SELECT clause.", parent=self)
                self._destroy_intellisense()
                return "break"  # Stop the event

        # If validation passes (or is not applicable), insert the text.
        self.query_view_text.insert(tk.INSERT, selected_field)
        self._destroy_intellisense()
        return "break"

    def _is_in_select_clause(self):
        """Checks if the cursor is currently between SELECT and FROM keywords."""
        all_text = self.query_view_text.get("1.0", "end-1c")
        cursor_pos = self.query_view_text.index(tk.INSERT)

        # Find positions of SELECT and FROM
        select_match = re.search(r"SELECT", all_text, re.I)
        from_match = re.search(r"FROM", all_text, re.I)

        if not select_match or not from_match:
            return False  # Can't determine context

        # Convert text widget indices to comparable integers
        cursor_offset = self.query_view_text.count("1.0", cursor_pos)[0]
        select_offset = self.query_view_text.count("1.0", f"1.0 + {select_match.end()} chars")[0]
        from_offset = self.query_view_text.count("1.0", f"1.0 + {from_match.start()} chars")[0]

        return select_offset <= cursor_offset < from_offset

    def _destroy_intellisense(self):
        """Destroys the intellisense popup if it exists."""
        if self.intellisense_popup:
            self.intellisense_popup.destroy()
            self.intellisense_popup = None
            self.query_view_text.focus_set()  # Return focus to the text editor

    ### END INTELLISENSE METHODS ###

    def _resize_query_results_columns(self):
        tree = self.results_tree
        if not tree["columns"]: return
        try:
            style = ttk.Style()
            font_name = style.lookup("Treeview", "font") or "TkDefaultFont"
            tree_font = tkFont.Font(font=font_name)
            padding = 20
            for col_id in tree["columns"]:
                header_text = tree.heading(col_id, "text").split(" ")[0]
                max_width = tree_font.measure(header_text) + padding
                for item_id in tree.get_children():
                    values = tree.item(item_id, "values")
                    try:
                        col_index = tree["columns"].index(col_id)
                        cell_value = str(values[col_index]) if values and col_index < len(values) else ""
                        cell_width = tree_font.measure(cell_value) + padding
                        if cell_width > max_width: max_width = cell_width
                    except (ValueError, IndexError):
                        continue
                max_width = min(max(max_width, 50), 500)
                tree.column(col_id, width=max_width, stretch=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resize columns: {e}", parent=self)

    def _populate_initial_dropdowns(self):
        self.table1_combo['values'] = self.table_names
        self.table2_combo['values'] = self.table_names_with_blank

    def _on_table_select(self, event=None):
        t1_name, t2_name = self.table1_var.get(), self.table2_var.get()
        self._update_available_fields()
        self.join_conditions.clear()
        self.filter_conditions.clear()
        self.conditions_listbox.delete(0, tk.END)

        if t1_name and not t2_name:
            self.conditions_frame.config(text="2. Define Filters (WHERE)")
            self.add_condition_button.config(text="Add Filter ->")
            self.join_controls_frame.grid_forget()
            self.filter_controls_frame.grid(row=0, column=0, sticky="ew")
            key1 = self.table_combobox_map.get(t1_name)
            cols = self.potential_tables[key1]['columns'] if key1 else []
            self.filter_field_combo['values'] = cols
            self.filter_field_combo.config(state="readonly" if cols else "disabled")
            self.join_type_rb_inner.config(state="disabled")
            self.join_type_rb_anti.config(state="disabled")
        elif t1_name and t2_name:
            self.conditions_frame.config(text="2. Define Join Conditions (ON)")
            self.add_condition_button.config(text="Add Join ->")
            self.filter_controls_frame.grid_forget()
            self.join_controls_frame.grid(row=0, column=0, sticky="ew")
            key1 = self.table_combobox_map.get(t1_name)
            self.field1_combo['values'] = self.potential_tables[key1]['columns'] if key1 else []
            self.field1_combo.config(state="readonly" if key1 else "disabled")
            key2 = self.table_combobox_map.get(t2_name)
            self.field2_combo['values'] = self.potential_tables[key2]['columns'] if key2 else []
            self.field2_combo.config(state="readonly" if key2 else "disabled")
            self.join_type_rb_inner.config(state="normal")
            self.join_type_rb_anti.config(state="normal")
        else:
            self.conditions_frame.config(text="2. Define Conditions")
            self.add_condition_button.config(text="Add...")
            self.join_controls_frame.grid_forget()
            self.filter_controls_frame.grid_forget()
            self.join_type_rb_inner.config(state="disabled")
            self.join_type_rb_anti.config(state="disabled")
        self._update_query_view()

    def _add_condition(self):
        if self.table2_var.get():
            self._add_join_condition()
        else:
            self._add_filter_condition()

    def _add_join_condition(self):
        f1, f2 = self.field1_var.get(), self.field2_var.get()
        if not (f1 and f2):
            messagebox.showwarning("Incomplete Join", "Please select a field from both tables.", parent=self)
            return
        self.join_conditions.append((f1, f2))
        self.conditions_listbox.insert(tk.END, f"T1.{f1} = T2.{f2}")
        self._update_query_view()

    def _add_filter_condition(self):
        field, op, value = self.filter_field_var.get(), self.filter_op_var.get(), self.filter_value_var.get()
        if not field or not op:
            messagebox.showwarning("Incomplete Filter", "Please select a field and an operator.", parent=self)
            return
        self.filter_conditions.append((field, op, value))
        self.conditions_listbox.insert(tk.END, f"T1.{field} {op} '{value}'")
        self._update_query_view()

    def _remove_condition(self):
        selected_indices = self.conditions_listbox.curselection()
        if not selected_indices: return
        target_list = self.join_conditions if self.table2_var.get() else self.filter_conditions
        for i in sorted(selected_indices, reverse=True):
            self.conditions_listbox.delete(i)
            del target_list[i]
        self._update_query_view()

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

    def _add_output_field(self):
        selected = self.available_fields_lb.curselection()
        for i in reversed(selected):
            self.selected_fields_lb.insert(tk.END, self.available_fields_lb.get(i))
            self.available_fields_lb.delete(i)
        self._update_query_view()

    def _add_all_output_fields(self):
        all_items = self.available_fields_lb.get(0, tk.END)
        for item in all_items: self.selected_fields_lb.insert(tk.END, item)
        self.available_fields_lb.delete(0, tk.END)
        self._update_query_view()

    def _remove_output_field(self):
        selected = self.selected_fields_lb.curselection()
        for i in reversed(selected):
            self.available_fields_lb.insert(tk.END, self.selected_fields_lb.get(i))
            self.selected_fields_lb.delete(i)
        self._update_query_view()

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

    def _update_query_view(self):
        t1, t2 = self.table1_var.get(), self.table2_var.get()
        select_clause = ",\n  ".join(self.selected_fields_lb.get(0, tk.END)) or "  [Select Output Fields]"
        limit_str = f"\nLIMIT {self.limit_value_var.get()}" if self.limit_enabled_var.get() else ""
        query_str = ""

        if t1 and not t2:
            clauses = [f"T1.{f} {op} '{v}'" for f, op, v in self.filter_conditions]
            where_clause = "\n    AND ".join(clauses) if clauses else "[Define Filter Conditions]"
            query_str = f"SELECT\n  {select_clause}\nFROM\n  {t1} AS T1\nWHERE\n  {where_clause}{limit_str};"
        elif t1 and t2:
            join_type = self.query_type_var.get().replace("ANTI", "ANTI-JOIN")
            clauses = [f"T1.{f1} = T2.{f2}" for f1, f2 in self.join_conditions]
            on_clause = "\n    AND ".join(clauses) if clauses else "[Define Join Conditions]"
            query_str = f"SELECT\n  {select_clause}\nFROM\n  {t1} AS T1\n{join_type}\n  {t2} AS T2\n  ON {on_clause}{limit_str};"
        else:
            query_str = "Please select at least one table to begin."

        is_editing = self.manual_edit_mode.get()
        if not is_editing:
            self.query_view_text.config(state="normal")
            self.query_view_text.delete("1.0", tk.END)
            self.query_view_text.insert("1.0", query_str)
            self.query_view_text.config(state="disabled")

    def _toggle_manual_edit(self):
        self.manual_edit_mode.set(not self.manual_edit_mode.get())
        is_editing = self.manual_edit_mode.get()
        self.query_view_text.config(state="normal" if is_editing else "disabled")
        self.validate_button.config(state="normal" if is_editing else "disabled")
        if is_editing:
            messagebox.showinfo("Manual Edit Mode",
                                "You can now edit the query text.\n"
                                "Type 'T1:' or 'T2:' to trigger Intellisense.\n"
                                "Click 'Validate & Apply' to parse your changes.",
                                parent=self)
        else:
            self._destroy_intellisense()
            self._update_query_view()  # Sync text view with visual designer state

    def _parse_and_apply_manual_query(self):
        query_text = self.query_view_text.get("1.0", tk.END)
        config = None
        try:
            if re.search(r"\s+(INNER|LEFT\s+ANTI)(?:\s+JOIN)?\s+", query_text, re.I):
                config = self._parse_join_query(query_text)
            elif re.search(r"\s+WHERE\s+", query_text, re.I):
                config = self._parse_filter_query(query_text)
            else:
                raise ValueError("Could not find a valid JOIN or WHERE clause.")

            self._populate_ui_from_config(config)
            messagebox.showinfo("Success", "Manual query validated and applied to the designer.", parent=self)

        # FIX: Catch the specific error and show a clean message.
        except ValueError as e:
            messagebox.showerror("Parsing Error", f"Could not apply manual query:\n\n{e}", parent=self)
        # Catch any other unexpected errors and show the full traceback for debugging.
        except Exception:
            error_details = f"An unexpected error occurred:\n\n{traceback.format_exc()}"
            messagebox.showerror("Unexpected Error", error_details, parent=self)

    def _parse_join_query(self, text):
        select_match = re.search(r"SELECT\s*(.*?)\s*FROM", text, re.S | re.I)
        if not select_match: raise ValueError("Could not find SELECT clause.")
        fields_str = select_match.group(1).strip()
        output_fields = [f.strip() for f in re.split(r'\s*,\s*', fields_str) if f.strip()]

        from_join_match = re.search(
            r"FROM\s+([\w-]+)\s+AS\s+T1\s+(INNER(?:\s+JOIN)?|LEFT\s+ANTI(?:-JOIN)?)\s+([\w-]+)\s+AS\s+T2", text,
            re.I | re.S)
        if not from_join_match: raise ValueError("Could not parse FROM/JOIN clause.")
        t1, join_type_raw, t2 = from_join_match.groups()
        join_type = "ANTI" if "ANTI" in join_type_raw.upper() else "INNER"

        on_match = re.search(r"\s+ON\s+(.*?)(?=\s*(?:LIMIT|;|$))", text, re.S | re.I)
        if not on_match: raise ValueError("Could not find ON clause.")
        conditions_str = on_match.group(1).strip()
        join_conditions = []
        for part in re.split(r"\s+AND\s+", conditions_str, flags=re.I):
            match = re.fullmatch(r"T1\.([\w-]+)\s*=\s*T2\.([\w-]+)", part.strip(), re.I)
            if not match: raise ValueError(f"Invalid join condition: '{part}'")
            join_conditions.append(match.groups())

        limit_match = re.search(r"LIMIT\s*(\d+)", text, re.I)
        return {
            "mode": "join", "table1": t1, "table2": t2, "join_conditions": join_conditions,
            "output_fields": output_fields, "query_type": join_type,
            "limit_enabled": bool(limit_match), "limit_value": int(limit_match.group(1)) if limit_match else 100
        }

    def _parse_filter_query(self, text):
        select_match = re.search(r"SELECT\s*(.*?)\s*FROM", text, re.S | re.I)
        if not select_match: raise ValueError("Could not find SELECT clause.")
        fields_str = select_match.group(1).strip()
        output_fields = [f.strip() for f in re.split(r'\s*,\s*', fields_str) if f.strip()]

        from_match = re.search(r"FROM\s+([\w-]+)\s+AS\s+T1", text, re.I | re.S)
        if not from_match: raise ValueError("Could not parse FROM clause.")
        t1 = from_match.group(1)

        where_match = re.search(r"\s+WHERE\s+(.*?)(?=\s*(?:LIMIT|;|$))", text, re.S | re.I)
        if not where_match: raise ValueError("Could not find WHERE clause.")
        conditions_str = where_match.group(1).strip()
        filter_conditions = []
        op_pattern = r"CONTAINS|NOT\s*CONTAINS|[<>=!]+"
        for part in re.split(r"\s+AND\s+", conditions_str, flags=re.I):
            match = re.match(fr"T1\.([\w-]+)\s+({op_pattern})\s+'([^']*)'", part.strip(), re.I)
            if not match: raise ValueError(f"Invalid filter condition: '{part}'")
            field, op, value = match.groups()
            filter_conditions.append((field, op.upper().replace(" ", ""), value))

        limit_match = re.search(r"LIMIT\s*(\d+)", text, re.I)
        return {
            "mode": "filter", "table1": t1, "table2": "", "filter_conditions": filter_conditions,
            "output_fields": output_fields,
            "limit_enabled": bool(limit_match), "limit_value": int(limit_match.group(1)) if limit_match else 100
        }

    def _save_config(self):
        t1_name = self.table1_var.get()
        if not t1_name:
            messagebox.showerror("Error", "Please select at least one table to save.", parent=self)
            return
        config_data = {
            "source_xml": self.source_xml_path,
            "output_fields": list(self.selected_fields_lb.get(0, tk.END)),
            "limit_enabled": self.limit_enabled_var.get(),
            "limit_value": self.limit_value_var.get()
        }
        if self.table2_var.get():
            config_data.update({
                "mode": "join", "table1": t1_name, "table2": self.table2_var.get(),
                "join_conditions": self.join_conditions, "query_type": self.query_type_var.get()
            })
        else:
            config_data.update({
                "mode": "filter", "table1": t1_name, "table2": "", "filter_conditions": self.filter_conditions
            })
        filepath = filedialog.asksaveasfilename(title="Save Query Configuration", defaultextension=".json",
                                                filetypes=[("Query Config Files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'w') as f:
                json.dump(config_data, f, indent=2)
            messagebox.showinfo("Success", f"Query configuration saved to:\n{filepath}", parent=self)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save configuration:\n{e}", parent=self)

    def _load_config(self):
        filepath = filedialog.askopenfilename(title="Load Query Configuration",
                                              filetypes=[("Query Config Files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'r') as f:
                config_data = json.load(f)
            if config_data.get("source_xml") != self.source_xml_path and not messagebox.askyesno("Warning",
                                                                                                 "This query was saved for a different XML file.\nLoad anyway?",
                                                                                                 parent=self):
                return
            self._populate_ui_from_config(config_data)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load or apply configuration:\n{e}", parent=self)
            traceback.print_exc()

    def _populate_ui_from_config(self, config):
        mode = config.get("mode", "join")
        self.table1_var.set(config["table1"])
        self.table2_var.set(config.get("table2", ""))
        self._on_table_select()

        self.conditions_listbox.delete(0, tk.END)
        self.join_conditions.clear()
        self.filter_conditions.clear()

        if mode == "join":
            self.join_conditions = config.get("join_conditions", [])
            for f1, f2 in self.join_conditions: self.conditions_listbox.insert(tk.END, f"T1.{f1} = T2.{f2}")
            self.query_type_var.set(config.get("query_type", "INNER"))
            if self.join_conditions:
                self.field1_var.set(self.join_conditions[0][0])
                self.field2_var.set(self.join_conditions[0][1])
        elif mode == "filter":
            self.filter_conditions = config.get("filter_conditions", [])
            for field, op, value in self.filter_conditions: self.conditions_listbox.insert(tk.END,
                                                                                           f"T1.{field} {op} '{value}'")
            if self.filter_conditions:
                self.filter_field_var.set(self.filter_conditions[0][0])
                self.filter_op_var.set(self.filter_conditions[0][1])
                self.filter_value_var.set(self.filter_conditions[0][2])

        self.selected_fields_lb.delete(0, tk.END)
        available_items = list(self.available_fields_lb.get(0, tk.END))
        for field in config.get("output_fields", []):
            if field in available_items:
                self.selected_fields_lb.insert(tk.END, field)
                try:
                    self.available_fields_lb.delete(list(self.available_fields_lb.get(0, tk.END)).index(field))
                except ValueError:
                    pass

        self.limit_enabled_var.set(config.get("limit_enabled", False))
        self.limit_value_var.set(config.get("limit_value", 100))
        self._toggle_limit_entry()
        self._update_query_view()

    def _export_results(self):
        if not self.results_tree.get_children():
            messagebox.showwarning("Export Error", "There are no results to export.", parent=self)
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not filepath: return
        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(self.results_tree["columns"])
                for item_id in self.results_tree.get_children():
                    writer.writerow(self.results_tree.item(item_id, "values"))
            messagebox.showinfo("Success", f"Results exported successfully.", parent=self)
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred during export:\n{e}", parent=self)

    def _run_query(self):
        t1_name, t2_name = self.table1_var.get(), self.table2_var.get()
        if not t1_name:
            messagebox.showerror("Error", "Please select a table (T1).", parent=self)
            return
        if not self.selected_fields_lb.get(0, tk.END):
            messagebox.showerror("Error", "Please select at least one field for the output.", parent=self)
            return
        try:
            if t1_name and not t2_name:
                self._run_filter_query()
            elif t1_name and t2_name:
                self._run_join_query()
        except Exception as e:
            messagebox.showerror("Query Error", f"An error occurred while running the query:\n{e}", parent=self)
            traceback.print_exc()

    def _run_filter_query(self):
        if not self.filter_conditions and not messagebox.askyesno("No Filters", "No filters are defined. Continue?",
                                                                  parent=self):
            return
        t1_name = self.table1_var.get()
        table1_info = self.potential_tables[self.table_combobox_map[t1_name]]
        rows1 = table1_info["parent_element"].findall(table1_info["row_tag"])
        results = []
        for row1_el in rows1:
            if self._row_matches_filters(row1_el, self.filter_conditions):
                result_row = {f"T1: {col}": row1_el.findtext(col, "").strip() for col in table1_info['columns']}
                results.append(result_row)
        if self.limit_enabled_var.get(): results = results[:self.limit_value_var.get()]
        self.current_results_data = results
        self._display_results_grid(self.selected_fields_lb.get(0, tk.END))

    def _row_matches_filters(self, row_element, filters):
        if not filters: return True
        for field, op, value in filters:
            cell_value = row_element.findtext(field, "").strip()
            cell_compare, value_compare = cell_value.lower(), value.lower()
            match = False
            if op == "CONTAINS":
                match = value_compare in cell_compare
            elif op == "NOT CONTAINS":
                match = value_compare not in cell_compare
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

    def _run_join_query(self):
        if not self.join_conditions:
            messagebox.showerror("Error", "Please add at least one join condition.", parent=self)
            return
        t1_name, t2_name = self.table1_var.get(), self.table2_var.get()
        key1, key2 = self.table_combobox_map[t1_name], self.table_combobox_map[t2_name]
        table1_info, table2_info = self.potential_tables[key1], self.potential_tables[key2]
        t2_join_fields = [jc[1] for jc in self.join_conditions]
        table2_index = defaultdict(list)
        for row_el in table2_info["parent_element"].findall(table2_info["row_tag"]):
            join_key = tuple(row_el.findtext(f, "").strip() for f in t2_join_fields)
            table2_index[join_key].append(row_el)
        results = []
        t1_join_fields = [jc[0] for jc in self.join_conditions]
        for row1_el in table1_info["parent_element"].findall(table1_info["row_tag"]):
            join_key = tuple(row1_el.findtext(f, "").strip() for f in t1_join_fields)
            matching_rows2 = table2_index.get(join_key, [])
            if self.query_type_var.get() == "INNER" and matching_rows2:
                for row2_el in matching_rows2:
                    result_row = {"Match_Count": len(matching_rows2)}
                    result_row.update({f"T1: {c}": row1_el.findtext(c, "").strip() for c in table1_info['columns']})
                    result_row.update({f"T2: {c}": row2_el.findtext(c, "").strip() for c in table2_info['columns']})
                    results.append(result_row)
            elif self.query_type_var.get() == "ANTI" and not matching_rows2:
                result_row = {"Match_Count": 0}
                result_row.update({f"T1: {c}": row1_el.findtext(c, "").strip() for c in table1_info['columns']})
                result_row.update({f"T2: {c}": "" for c in table2_info['columns']})
                results.append(result_row)
        if self.limit_enabled_var.get(): results = results[:self.limit_value_var.get()]
        self.current_results_data = results
        self._display_results_grid(self.selected_fields_lb.get(0, tk.END))

    def _display_results_grid(self, output_fields):
        for item in self.results_tree.get_children(): self.results_tree.delete(item)
        display_columns = list(output_fields)
        if self.table2_var.get() and self.query_type_var.get() == "INNER" and "Match_Count" not in display_columns:
            display_columns.insert(0, "Match_Count")
        self.results_tree["columns"] = display_columns
        for col in display_columns:
            self.results_tree.heading(col, text=col, anchor='w')
            self.results_tree.column(col, width=120, anchor='w', stretch=True)
        if "Match_Count" in display_columns: self.results_tree.column("Match_Count", width=80, anchor='center',
                                                                      stretch=False)
        if not self.current_results_data: messagebox.showinfo("Query Results", "The query returned no records.",
                                                              parent=self)
        self.results_sort_col = None
        self._sort_and_redisplay_results()

    def _on_results_click(self, event):
        if self.results_tree.identify("region", event.x, event.y) == "heading":
            col_id = self.results_tree.identify_column(event.x)
            self._sort_by_column(self.results_tree.column(col_id, "id"))

    def _sort_by_column(self, col_name):
        if self.results_sort_col == col_name:
            self.results_sort_asc = not self.results_sort_asc
        else:
            self.results_sort_col, self.results_sort_asc = col_name, True
        self._sort_and_redisplay_results()

    def _sort_and_redisplay_results(self):
        def sort_key(item):
            value = item.get(self.results_sort_col, "")
            try:
                return float(value)
            except (ValueError, TypeError):
                return str(value).lower()

        if self.results_sort_col: self.current_results_data.sort(key=sort_key, reverse=not self.results_sort_asc)
        self._update_results_header_style()
        for item in self.results_tree.get_children(): self.results_tree.delete(item)
        for result_row in self.current_results_data:
            values = [result_row.get(col, "") for col in self.results_tree["columns"]]
            self.results_tree.insert("", "end", values=values)

    def _update_results_header_style(self):
        for col in self.results_tree["columns"]:
            text = col
            if col == self.results_sort_col: text += " ▲" if self.results_sort_asc else " ▼"
            self.results_tree.heading(col, text=text)


# --------------------------------#
# -----XMLNotepad --------#
# -----------------------------------#
class XMLNotepad:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Notepad")
        self.root.geometry("1300x850")
        self.xml_tree_root = None
        self.tree_item_to_element = {}
        self.selected_element_for_context_menu = None
        self.potential_tables = {}
        self.current_right_panel_view = "details"
        self.table_combobox_map = {}
        self.table_sort_column_id = None
        self.table_sort_direction_is_asc = True
        self.current_loaded_filepath = ""
        self.active_cell_editor = None
        self.selected_table_row_id = None
        self.clipboard_row_data = None
        self.undo_stack = deque(maxlen=UNDO_STACK_SIZE)
        self.redo_stack = deque(maxlen=UNDO_STACK_SIZE)

        # In-memory cache for loaded table data
        self.table_data_cache = {}

        self.setup_menu()
        self.setup_top_controls_frame()
        self.setup_paned_window()
        self.setup_context_menu()
        self.setup_status_frame()
        self.root.bind_all("<Control-o>", lambda event: self.open_xml_file_threaded())
        self.root.bind_all("<Control-s>", self.handle_ctrl_s)
        self.root.bind_all("<Control-f>", lambda event: self.show_find_dialog())

    def setup_menu(self):
        self.menubar = tk.Menu(self.root)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Open XML...", command=self.open_xml_file_threaded, accelerator="Ctrl+O")
        self.filemenu.add_command(label="Save XML As...", command=self.save_xml_as, state="disabled")
        self.filemenu.add_command(label="Save Table as CSV...", command=self.save_current_table_as_csv,
                                  state="disabled", accelerator="Ctrl+S")
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=self.root.quit)
        self.menubar.add_cascade(label="File", menu=self.filemenu)
        self.editmenu = tk.Menu(self.menubar, tearoff=0)
        self.editmenu.add_command(label="Cut Row", command=self.cut_row, state="disabled")
        self.editmenu.add_command(label="Copy Row", command=self.copy_row, state="disabled")
        self.editmenu.add_command(label="Paste Row", command=self.paste_row, state="disabled")
        self.editmenu.add_command(label="Delete Row", command=self.delete_row, state="disabled")
        self.editmenu.add_separator()
        self.editmenu.add_command(label="Find...", command=self.show_find_dialog, state="disabled")
        self.editmenu.add_command(label="Resize Columns", command=self.resize_columns)
        self.editmenu.add_command(label="Undo", command=self.undo_action, state="disabled")
        self.editmenu.add_command(label="Redo", command=self.redo_action, state="disabled")
        self.menubar.add_cascade(label="Edit", menu=self.editmenu)
        self.utilsmenu = tk.Menu(self.menubar, tearoff=0)
        self.utilsmenu.add_command(label="Query Designer...", command=self.open_query_designer, state="disabled")
        self.menubar.add_cascade(label="Utils", menu=self.utilsmenu)
        self.root.config(menu=self.menubar)

    def open_query_designer(self):
        if not self.potential_tables:
            messagebox.showwarning("Query Designer", "No tables found in the XML file to query.")
            return
        QueryDesigner(self.root, self.potential_tables, self.table_combobox_map, self.current_loaded_filepath)

    def setup_top_controls_frame(self):
        self.top_controls_frame = ttk.Frame(self.root)
        self.top_controls_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))
        self.tables_label = ttk.Label(self.top_controls_frame, text="Tables:")
        self.tables_label.pack(side=tk.LEFT, padx=(0, 5))
        self.selected_table_var = tk.StringVar()
        self.tables_combobox = ttk.Combobox(self.top_controls_frame, textvariable=self.selected_table_var,
                                            state="disabled", width=35)
        self.tables_combobox.pack(side=tk.LEFT, padx=(0, 10))
        self.tables_combobox.bind("<<ComboboxSelected>>", self.on_table_combobox_select)
        self.filename_label_prefix = ttk.Label(self.top_controls_frame, text="File:")
        self.filename_label_prefix.pack(side=tk.LEFT, padx=(10, 2))
        self.filename_display_var = tk.StringVar(value="No file loaded.")
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
        self.right_panel_outer = ttk.Labelframe(self.paned_window, text="Content", padding="5")
        self.paned_window.add(self.right_panel_outer, weight=2)
        self.detail_tree = ttk.Treeview(self.right_panel_outer, columns=("Property", "Value"), show='headings')
        self.detail_tree.heading("Property", text="Property", anchor='w')
        self.detail_tree.heading("Value", text="Value", anchor='w')
        self.detail_tree.column("Property", width=220, stretch=tk.NO, anchor='w')
        self.detail_tree.column("Value", width=400, anchor='w')
        self.detail_tree_vsb = ttk.Scrollbar(self.right_panel_outer, orient="vertical", command=self.detail_tree.yview)
        self.detail_tree.configure(yscrollcommand=self.detail_tree_vsb.set)
        self.detail_tree_vsb.pack(side='right', fill='y')
        self.detail_tree.pack(side='left', expand=True, fill='both')
        self.table_treeview = ttk.Treeview(self.right_panel_outer, show='headings')
        self.table_treeview_vsb = ttk.Scrollbar(self.right_panel_outer, orient="vertical",
                                                command=self.table_treeview.yview)
        self.table_treeview_hsb = ttk.Scrollbar(self.right_panel_outer, orient="horizontal",
                                                command=self.table_treeview.xview)
        self.table_treeview.configure(yscrollcommand=self.table_treeview_vsb.set,
                                      xscrollcommand=self.table_treeview_hsb.set)
        self.table_treeview.tag_configure('selected', background='#e6f3ff')
        self.table_treeview.bind("<Button-1>", self.on_table_tree_click)
        self.table_treeview.bind("<Double-1>", self.on_table_cell_or_header_double_click)

    def setup_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Export Node as CSV...", command=self.export_node_as_csv, state="disabled")

    def setup_status_frame(self):
        self.status_frame = ttk.Frame(self.root, padding=(2, 2))
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = ttk.Label(self.status_frame, text="Ready", anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(self.status_frame, orient='horizontal', mode='determinate')

    def update_status(self, message, show_progress=False, progress_value=0):
        self.status_label.config(text=message)
        if show_progress:
            if not self.progress_bar.winfo_ismapped():
                self.progress_bar.pack(side=tk.RIGHT, padx=5, pady=2, fill=tk.X, expand=False, ipadx=50)
            self.progress_bar['value'] = progress_value
        else:
            if self.progress_bar.winfo_ismapped():
                self.progress_bar.pack_forget()
        self.root.update_idletasks()

    def handle_ctrl_s(self, event=None):
        if self.current_right_panel_view == "table" and self.table_treeview["columns"]:
            self.save_current_table_as_csv()
        return "break"

    def resize_columns(self):
        try:
            if not self.table_treeview["columns"]:
                return
            style = ttk.Style()
            font_name = style.lookup("Treeview", "font") or "TkDefaultFont"
            tree_font = tkFont.Font(font=font_name)
            padding = 20
            for col_id in self.table_treeview["columns"]:
                header_text = self.table_treeview.heading(col_id, "text")
                max_width = tree_font.measure(header_text) + padding
                for item_id in self.table_treeview.get_children():
                    values = self.table_treeview.item(item_id, "values")
                    col_index = self.table_treeview["columns"].index(col_id)
                    cell_value = values[col_index].strip() if values and col_index < len(values) else ""
                    cell_width = tree_font.measure(cell_value) + padding
                    max_width = max(max_width, cell_width)
                max_width = min(max(max_width, 50), 500)
                self.table_treeview.column(col_id, width=max_width, stretch=False)
            self.table_treeview_hsb.pack(side='bottom', fill='x')
            self.table_treeview.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resize columns: {str(e)}\n{traceback.format_exc()}")

    def _reset_ui_for_new_file(self):
        self.update_status("Resetting UI...", show_progress=False)
        self.current_loaded_filepath = ""
        self.filename_display_var.set("No file loaded.")
        self.selected_table_var.set('')
        self.tables_combobox.set('')
        self.tables_combobox['values'] = []
        self.tables_combobox.config(state="disabled")
        self.table_combobox_map.clear()
        self.potential_tables.clear()
        self.clear_xml_treeview()
        self.switch_to_details_view(None)
        self.xml_tree_root = None
        self.selected_element_for_context_menu = None
        self.selected_table_row_id = None
        self.clipboard_row_data = None
        self.undo_stack.clear()
        self.redo_stack.clear()

        # Clear the cache when a new file is loaded
        self.table_data_cache.clear()

        self.filemenu.entryconfig("Save Table as CSV...", state="disabled")
        self.filemenu.entryconfig("Save XML As...", state="disabled")
        self.editmenu.entryconfig("Cut Row", state="disabled")
        self.editmenu.entryconfig("Copy Row", state="disabled")
        self.editmenu.entryconfig("Paste Row", state="disabled")
        self.editmenu.entryconfig("Delete Row", state="disabled")
        self.editmenu.entryconfig("Find...", state="disabled")
        self.editmenu.entryconfig("Resize Columns", state="disabled")
        self.editmenu.entryconfig("Undo", state="disabled")
        self.editmenu.entryconfig("Redo", state="disabled")
        self.context_menu.entryconfig("Export Node as CSV...", state="disabled")
        self.utilsmenu.entryconfig("Query Designer...", state="disabled")
        self.root.update_idletasks()

    def open_xml_file_threaded(self):
        filepath = filedialog.askopenfilename(title="Open XML File",
                                              filetypes=(("XML files", "*.xml"), ("All files", "*.*")))
        if not filepath:
            return
        self.filemenu.entryconfig("Open XML...", state="disabled")
        self._reset_ui_for_new_file()
        self.current_loaded_filepath = filepath
        threading.Thread(target=self._parse_and_populate_worker, args=(filepath,), daemon=True).start()

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
                    self.root.after(0, self.update_status,
                                    f"Loading {os.path.basename(filepath)} ({int(progress)}%)...", True, progress)
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

    def _finish_loading_success(self, loaded_filepath):
        self.filename_display_var.set(os.path.basename(loaded_filepath))
        self.filemenu.entryconfig("Save XML As...", state="normal")
        self.editmenu.entryconfig("Find...", state="normal")
        self.root.update_idletasks()
        self.update_status(f"Populating tree for {os.path.basename(loaded_filepath)}...", True, 0)
        self.populate_main_xml_treeview()
        self.update_status("Discovering tables...", True, 85)
        self.discover_potential_tables()
        self.populate_table_combobox()
        if self.potential_tables:
            self.utilsmenu.entryconfig("Query Designer...", state="normal")
        self.root.update_idletasks()
        self.update_status(f"Loaded: {os.path.basename(loaded_filepath)}", show_progress=False)

    def _finish_loading_error(self, error_message):
        self.filename_display_var.set("Load Error.")
        self.current_loaded_filepath = ""
        self.update_status(f"Error: {error_message[:100]}...", show_progress=False)
        messagebox.showerror("Load Error", error_message)
        self.tables_combobox.set('')
        self.tables_combobox['values'] = []
        self.tables_combobox.config(state="disabled")
        self.filemenu.entryconfig("Save XML As...", state="disabled")
        self.editmenu.entryconfig("Find...", state="disabled")
        self.utilsmenu.entryconfig("Query Designer...", state="disabled")
        self.root.update_idletasks()

    def clear_xml_treeview(self):
        children_to_delete = list(self.tree.get_children(""))
        for item in children_to_delete:
            try:
                self.tree.delete(item)
            except tk.TclError:
                pass
        self.tree_item_to_element.clear()

    def populate_main_xml_treeview(self):
        if self.xml_tree_root is None:
            self.update_status("XML root not found.", False)
            return
        self.root.after(20, self._add_node_to_main_xml_treeview_recursive, self.xml_tree_root, "", False, 0, True)

    def _add_node_to_main_xml_treeview_recursive(self, element, parent_item_id, open_node, depth, is_first_level_call):
        if depth > 250:
            node_text = f"<{element.tag}> (Children not displayed: depth > 250)"
            try:
                self.tree.insert(parent_item_id, 'end', text=node_text, open=False)
            except tk.TclError:
                pass
            return
        node_text = f"<{element.tag}>"
        if element.attrib:
            first_attr_key, first_attr_val = list(element.attrib.items())[0]
            val_str = str(first_attr_val)
            attr_preview = f'{first_attr_key}="{val_str[:15]}"'
            if len(val_str) > 15 or len(element.attrib) > 1:
                attr_preview += "..."
            node_text += f" [{attr_preview}]"
        if len(element) > 0:
            node_text += f" ({len(element)})"
        try:
            item_id = self.tree.insert(parent_item_id, 'end', text=node_text, open=open_node)
            self.tree_item_to_element[item_id] = element
        except tk.TclError:
            return
        children = list(element)
        total_children_at_this_level = len(children)
        if is_first_level_call and total_children_at_this_level == 0:
            self.update_status("Tree populated.", False)
        for i, child_element in enumerate(children):
            self._add_node_to_main_xml_treeview_recursive(child_element, item_id, False, depth + 1, False)
            if is_first_level_call and total_children_at_this_level > 0:
                if (i + 1) % 20 == 0 or (i + 1) == total_children_at_this_level:
                    progress = int(((i + 1) / total_children_at_this_level) * 80)
                    self.update_status(f"Populating tree ({progress}%)...", True, progress)
                    self.tree.update_idletasks()

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
                    display_name_candidate = most_common_tag.capitalize()
                    if parent_element.tag != self.xml_tree_root.tag or len(
                            Counter(c.tag for c in self.xml_tree_root if c.tag is not None).keys()) == 1:
                        display_name_candidate = parent_element.tag.capitalize()
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
            self.editmenu.entryconfig("Find...", state="normal")
        else:
            self.tables_combobox.config(state="disabled")
            self.editmenu.entryconfig("Find...", state="disabled")

    def show_find_dialog(self):
        try:
            if not self.potential_tables:
                messagebox.showwarning("Warning", "No tables available to search.")
                return
            find_dialog = tk.Toplevel(self.root)
            find_dialog.title("Find in Table")
            find_dialog.geometry("600x400")
            find_dialog.transient(self.root)
            find_dialog.grab_set()
            ttk.Label(find_dialog, text="Table:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            table_var = tk.StringVar()
            table_combobox = ttk.Combobox(find_dialog, textvariable=table_var, state="readonly")
            table_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
            table_names = list(self.table_combobox_map.keys())
            table_combobox['values'] = table_names
            current_table = self.selected_table_var.get()
            if current_table in table_names:
                table_var.set(current_table)
            elif table_names:
                table_var.set(table_names[0])
            ttk.Label(find_dialog, text="Field:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
            field_var = tk.StringVar()
            field_combobox = ttk.Combobox(find_dialog, textvariable=field_var, state="readonly")
            field_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

            def update_fields(*args):
                selected_table = table_var.get()
                internal_key = self.table_combobox_map.get(selected_table)
                if internal_key:
                    columns = self.potential_tables[internal_key]["columns"]
                    field_combobox['values'] = columns
                    if columns:
                        field_var.set(columns[0])
                    else:
                        field_var.set("")

            table_var.trace("w", update_fields)
            update_fields()
            ttk.Label(find_dialog, text="Value:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
            search_var = tk.StringVar()
            search_entry = ttk.Entry(find_dialog, textvariable=search_var)
            search_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
            search_entry.focus_set()
            match_case_var = tk.BooleanVar()
            ttk.Checkbutton(find_dialog, text="Match case", variable=match_case_var).grid(row=3, column=1, padx=5,
                                                                                          pady=5, sticky="w")
            results_frame = ttk.Frame(find_dialog)
            results_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
            results_tree = ttk.Treeview(results_frame, columns=("Row",), show="headings", height=5)
            results_tree.heading("Row", text="Matching Row")
            results_tree.column("Row", width=100, anchor="center")
            results_tree.pack(side="left", fill="both", expand=True)
            results_vsb = ttk.Scrollbar(results_frame, orient="vertical", command=results_tree.yview)
            results_vsb.pack(side="right", fill="y")
            results_tree.configure(yscrollcommand=results_vsb.set)
            nav_frame = ttk.Frame(find_dialog)
            nav_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
            next_button = ttk.Button(nav_frame, text="Next", state="disabled")
            next_button.pack(side="left", padx=5)
            prev_button = ttk.Button(nav_frame, text="Previous", state="disabled")
            prev_button.pack(side="left", padx=5)
            current_match_index = [0]
            matches = []

            def perform_search():
                for item in results_tree.get_children():
                    results_tree.delete(item)
                matches.clear()
                current_match_index[0] = 0
                selected_table = table_var.get()
                internal_key = self.table_combobox_map.get(selected_table)
                if not internal_key: return
                table_info = self.potential_tables[internal_key]
                row_elements = table_info["parent_element"].findall(table_info["row_tag"])
                field = field_var.get()
                if not field: return
                search_value = search_var.get()
                if not search_value: return
                match_case = match_case_var.get()
                for i, element in enumerate(row_elements, 1):
                    col_element = element.find(field)
                    cell_text = col_element.text.strip() if col_element is not None and col_element.text else ""
                    compare_text = cell_text if match_case else cell_text.lower()
                    compare_search = search_value if match_case else search_value.lower()
                    if compare_search in compare_text:
                        matches.append((i, element))
                if not matches:
                    messagebox.showinfo("Info", "No matches found.")
                    return
                for i, (row_num, _) in enumerate(matches):
                    item_id = results_tree.insert("", "end", values=(f"Row {row_num}",))
                    results_tree.item(item_id, tags=(f"match_{i}",))
                    results_tree.tag_bind(f"match_{i}", "<Button-1>", lambda e, rn=row_num: jump_to_row(rn))
                if len(matches) > 1:
                    next_button.config(state="normal")
                    prev_button.config(state="normal")
                highlight_match(0)

            def highlight_match(index):
                for item in results_tree.get_children():
                    results_tree.item(item, tags=())
                if matches:
                    item_id = results_tree.get_children()[index]
                    results_tree.item(item_id, tags=("selected",))
                    results_tree.tag_configure("selected", background="lightblue")
                    results_tree.see(item_id)

            def jump_to_row(row_num):
                selected_table = table_var.get()
                internal_key = self.table_combobox_map.get(selected_table)
                if internal_key and (
                        self.current_right_panel_view != "table" or self.selected_table_var.get() != selected_table):
                    self.selected_table_var.set(selected_table)
                    self.on_table_combobox_select()
                for item_id in self.table_treeview.get_children():
                    values = self.table_treeview.item(item_id, "values")
                    if values and values[0] == str(row_num):
                        self.table_treeview.selection_set(item_id)
                        self.table_treeview.see(item_id)
                        break

            def next_match():
                if matches:
                    current_match_index[0] = (current_match_index[0] + 1) % len(matches)
                    highlight_match(current_match_index[0])
                    jump_to_row(matches[current_match_index[0]][0])

            def prev_match():
                if matches:
                    current_match_index[0] = (current_match_index[0] - 1) % len(matches)
                    highlight_match(current_match_index[0])
                    jump_to_row(matches[current_match_index[0]][0])

            next_button.config(command=next_match)
            prev_button.config(command=prev_match)
            search_button = ttk.Button(find_dialog, text="Find", command=perform_search)
            search_button.grid(row=6, column=0, columnspan=2, pady=10)
            search_entry.bind("<Return>", lambda e: perform_search())
            find_dialog.columnconfigure(1, weight=1)
            find_dialog.rowconfigure(4, weight=1)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to show find dialog: {str(e)}\n{traceback.format_exc()[:200]}")

    def on_table_combobox_select(self, event=None):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            self.deselect_row()
            selected_display_name = self.selected_table_var.get()
            if selected_display_name:
                internal_key = self.table_combobox_map.get(selected_display_name)
                if internal_key:
                    self.switch_to_table_view()
                    self.display_table_view_data(internal_key)
                    self.filemenu.entryconfig("Save Table as CSV...", state="normal")
                    self.editmenu.entryconfig("Find...", state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to select table: {str(e)}\n{traceback.format_exc()[:200]}")

    def on_xml_tree_node_select(self, event=None):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            self.deselect_row()
            if self.current_right_panel_view == "table":
                self.selected_table_var.set('')
            selected_item_id = self.tree.focus()
            element = self.tree_item_to_element.get(selected_item_id) if selected_item_id else None
            self.selected_element_for_context_menu = element
            self.switch_to_details_view(element)
            self._update_context_menu_state(element)
            self.filemenu.entryconfig("Save Table as CSV...", state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to select XML node: {str(e)}\n{traceback.format_exc()[:200]}")

    def switch_to_details_view(self, element):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            self.deselect_row()
            if self.current_right_panel_view == "table":
                self.table_treeview.pack_forget()
                self.table_treeview_vsb.pack_forget()
                self.table_treeview_hsb.pack_forget()
                self.detail_tree_vsb.pack(side='right', fill='y')
                self.detail_tree.pack(side='left', expand=True, fill='both')
                self.current_right_panel_view = "details"
            self.right_panel_outer.config(text="Node Details" if element else "Content")
            self.update_node_detail_panel(element)
            self.root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to switch to details view: {str(e)}\n{traceback.format_exc()[:200]}")

    def switch_to_table_view(self):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            self.deselect_row()
            if self.current_right_panel_view == "details":
                self.detail_tree.pack_forget()
                self.detail_tree_vsb.pack_forget()
                self.table_treeview_hsb.pack(side='bottom', fill='x')
                self.table_treeview_vsb.pack(side='right', fill='y')
                self.table_treeview.pack(side='left', expand=True, fill='both')
                self.current_right_panel_view = "table"
            if self.tree.selection():
                self.tree.selection_remove(self.tree.selection())
            self.root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to switch to table view: {str(e)}\n{traceback.format_exc()[:200]}")

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
            messagebox.showerror("Error", f"Failed to update detail panel: {str(e)}\n{traceback.format_exc()[:200]}")

    def display_table_view_data(self, internal_table_key):
        try:
            for item in self.table_treeview.get_children():
                self.table_treeview.delete(item)
            self.table_treeview["columns"] = []
            self.table_sort_column_id = None
            self.deselect_row()

            table_info = self.potential_tables.get(internal_table_key)
            if not table_info:
                messagebox.showerror("Error", f"Table data for '{internal_table_key}' not found.")
                return

            actual_combobox_label = self.selected_table_var.get()
            columns = table_info["columns"]
            data_to_populate = []

            # --- CACHE LOGIC ---
            if internal_table_key in self.table_data_cache:
                data_to_populate = self.table_data_cache[internal_table_key]
                self.right_panel_outer.config(
                    text=f"Table: {actual_combobox_label} ({len(data_to_populate)} rows) [Cached]")
            else:
                # Data not in cache, so parse from XML and populate cache
                parent_element = table_info["parent_element"]
                row_tag = table_info["row_tag"]
                row_elements = parent_element.findall(row_tag)
                self.right_panel_outer.config(text=f"Table: {actual_combobox_label} ({len(row_elements)} rows)")

                for i, row_element in enumerate(row_elements):
                    row_data = {"#": str(i + 1), "_element": row_element}
                    for col_name in columns:
                        col_data_element = row_element.find(col_name)
                        cell_text = col_data_element.text.strip() if col_data_element is not None and col_data_element.text else ""
                        row_data[col_name] = cell_text
                    data_to_populate.append(row_data)

                self.table_data_cache[internal_table_key] = data_to_populate
            # --- END CACHE LOGIC ---

            display_columns = ["#"] + columns
            self.table_treeview["columns"] = display_columns
            self.table_treeview.column("#0", width=0, stretch=tk.NO)
            self.table_treeview.column("#", width=50, stretch=tk.NO, anchor='e')
            self.table_treeview.heading("#", text="#", anchor='e')
            for col_name in columns:
                self.table_treeview.column(col_name, width=100, anchor='w')
                self.table_treeview.heading(col_name, text=col_name, anchor='w')

            self._populate_table_rows(data_to_populate, columns, actual_combobox_label)

            self.resize_columns()
            self.editmenu.entryconfig("Resize Columns", state="normal")
            self.table_treeview_hsb.pack(side='bottom', fill='x')
            self.root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display table: {str(e)}\n{traceback.format_exc()[:200]}")

    def _populate_table_rows(self, data_to_populate, data_columns, table_label):
        try:
            for item in self.table_treeview.get_children():
                self.table_treeview.delete(item)

            self.update_status(f"Loading table '{table_label}' ({len(data_to_populate)} rows)...", True, 0)
            all_display_columns = ["#"] + data_columns
            for i, row_data in enumerate(data_to_populate):
                display_values = [row_data.get(col_id, "") for col_id in all_display_columns]
                item_id = self.table_treeview.insert("", "end", values=display_values)
                if "_element" in row_data:
                    self.table_treeview.item(item_id, tags=(str(id(row_data["_element"])),))
                if (i + 1) % 50 == 0 or (i + 1) == len(data_to_populate):
                    progress = ((i + 1) / len(data_to_populate)) * 100 if len(data_to_populate) > 0 else 100
                    self.update_status(f"Loading table '{table_label}' ({int(progress)}%)...", show_progress=True,
                                       progress_value=progress)
            self.update_status(f"Table '{table_label}' loaded.", show_progress=False)
            if self.selected_table_row_id:
                for item_id in self.table_treeview.get_children():
                    if item_id == self.selected_table_row_id:
                        self.table_treeview.item(item_id, tags=('selected',))
                        break
        except Exception as e:
            messagebox.showerror("Error", f"Failed to populate table rows: {str(e)}\n{traceback.format_exc()[:200]}")

    def on_table_tree_click(self, event):
        try:
            if self.active_cell_editor:
                if event.widget != self.active_cell_editor:
                    self._finish_cell_edit()
            region = self.table_treeview.identify("region", event.x, event.y)
            if region == "heading":
                self.on_table_header_click_for_sort(event)
        except Exception as e:
            messagebox.showerror("Error", f"Table click failed: {str(e)}\n{traceback.format_exc()[:200]}")

    def on_table_cell_or_header_double_click(self, event):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            if self.current_right_panel_view != "table":
                return
            region = self.table_treeview.identify("region", event.x, event.y)
            if region == "heading":
                self.resize_columns()
                return
            item_id = self.table_treeview.identify_row(event.y)
            column_id_str = self.table_treeview.identify_column(event.x)
            if not item_id or not column_id_str:
                return
            columns = self.table_treeview["columns"]
            actual_column_name = self.table_treeview.column(column_id_str, "id")
            if actual_column_name == "#":
                self.select_row(item_id)
                return
            col_index = columns.index(actual_column_name)
            bbox = self.table_treeview.bbox(item_id, column_id_str)
            if not bbox:
                return
            x, y, width, height = bbox
            current_row_values = self.table_treeview.item(item_id, "values")
            original_value = current_row_values[col_index]
            self.active_cell_editor = ttk.Entry(self.table_treeview)
            self.active_cell_editor.insert(0, original_value)
            self.active_cell_editor.place(x=x, y=y, width=width, height=height, anchor='nw')
            self.active_cell_editor.focus_set()
            self.active_cell_editor.select_range(0, 'end')
            self.active_cell_editor.editing_item_id = item_id
            self.active_cell_editor.editing_col_index = col_index
            self.active_cell_editor.editing_column_name = actual_column_name
            tags = self.table_treeview.item(item_id, "tags")
            if tags:
                element_id = int(tags[0])
                selected_table = self.selected_table_var.get()
                internal_key = self.table_combobox_map.get(selected_table)
                if internal_key:
                    self.active_cell_editor.internal_key = internal_key  # Store for undo/cache
                    table_info = self.potential_tables.get(internal_key)
                    if table_info:
                        row_elements = table_info["parent_element"].findall(table_info["row_tag"])
                        for element in row_elements:
                            if id(element) == element_id:
                                self.active_cell_editor.element = element
                                break
            self.active_cell_editor.bind("<Return>", lambda e: self._finish_cell_edit(commit=True))
            self.active_cell_editor.bind("<Escape>", lambda e: self._finish_cell_edit(commit=False))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to handle double-click: {str(e)}\n{traceback.format_exc()[:200]}")

    def select_row(self, item_id):
        try:
            if self.selected_table_row_id:
                for item in self.table_treeview.get_children():
                    tags = list(self.table_treeview.item(item, "tags"))
                    if 'selected' in tags:
                        tags.remove('selected')
                        self.table_treeview.item(item, tags=tags)
            self.selected_table_row_id = item_id
            current_tags = list(self.table_treeview.item(item_id, "tags"))
            current_tags.append('selected')
            self.table_treeview.item(item_id, tags=tuple(current_tags))
            self.editmenu.entryconfig("Cut Row", state="normal")
            self.editmenu.entryconfig("Copy Row", state="normal")
            self.editmenu.entryconfig("Delete Row", state="normal")
            self.table_treeview.see(item_id)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to select row: {str(e)}\n{traceback.format_exc()[:200]}")

    def deselect_row(self):
        try:
            if self.selected_table_row_id:
                for item in self.table_treeview.get_children():
                    tags = list(self.table_treeview.item(item, "tags"))
                    if 'selected' in tags:
                        tags.remove('selected')
                        self.table_treeview.item(item, tags=tuple(tags))
                self.selected_table_row_id = None
                self.editmenu.entryconfig("Cut Row", state="disabled")
                self.editmenu.entryconfig("Copy Row", state="disabled")
                self.editmenu.entryconfig("Delete Row", state="disabled")
            if not self.clipboard_row_data:
                self.editmenu.entryconfig("Paste Row", state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to deselect row: {str(e)}\n{traceback.format_exc()[:200]}")

    # --- NEW --- Helper method to invalidate cache for a specific table
    def _invalidate_table_cache(self, internal_key):
        """Removes a table's data from the cache, forcing a reload from XML."""
        if internal_key in self.table_data_cache:
            del self.table_data_cache[internal_key]
            print(f"Cache invalidated for key: {internal_key}")  # For debugging

    def copy_row(self):
        try:
            if not self.selected_table_row_id: return
            values = self.table_treeview.item(self.selected_table_row_id, "values")
            tags = self.table_treeview.item(self.selected_table_row_id, "tags")
            if not tags: return
            element_id = int(tags[0])
            selected_table = self.selected_table_var.get()
            internal_key = self.table_combobox_map.get(selected_table)
            if not internal_key: return
            table_info = self.potential_tables.get(internal_key)
            row_elements = table_info["parent_element"].findall(table_info["row_tag"])
            element = None
            for el in row_elements:
                if id(el) == element_id:
                    element = el
                    break
            if not element: return
            self.clipboard_row_data = {
                "values": list(values), "element": element,
                "row_tag": element.tag, "internal_key": internal_key
            }
            self.editmenu.entryconfig("Paste Row", state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy row: {str(e)}\n{traceback.format_exc()[:200]}")

    def cut_row(self):
        try:
            if not self.selected_table_row_id: return
            values = self.table_treeview.item(self.selected_table_row_id, "values")
            tags = self.table_treeview.item(self.selected_table_row_id, "tags")
            if not tags: return
            element_id = int(tags[0])
            selected_table = self.selected_table_var.get()
            internal_key = self.table_combobox_map.get(selected_table)
            if not internal_key: return
            table_info = self.potential_tables.get(internal_key)
            row_elements = table_info["parent_element"].findall(table_info["row_tag"])
            element = None
            for el in row_elements:
                if id(el) == element_id:
                    element = el
                    break
            if not element: return
            parent_element = table_info["parent_element"]
            index = list(parent_element).index(element)
            self.clipboard_row_data = {
                "values": list(values), "element": element,
                "row_tag": element.tag, "internal_key": internal_key
            }
            self.undo_stack.append({
                "action": "cut", "item_id": self.selected_table_row_id, "element": element,
                "parent_element": parent_element, "index": index, "values": values, "internal_key": internal_key
            })
            self.redo_stack.clear()
            self._invalidate_table_cache(internal_key)
            parent_element.remove(element)
            self.table_treeview.delete(self.selected_table_row_id)
            self.selected_table_row_id = None
            self.editmenu.entryconfig("Cut Row", state="disabled")
            self.editmenu.entryconfig("Copy Row", state="disabled")
            self.editmenu.entryconfig("Delete Row", state="disabled")
            self.editmenu.entryconfig("Paste Row", state="normal")
            self.editmenu.entryconfig("Undo", state="normal")
            self.editmenu.entryconfig("Redo", state="disabled")
            self._refresh_table()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to cut row: {str(e)}\n{traceback.format_exc()[:200]}")

    def paste_row(self):
        try:
            if not self.clipboard_row_data: return
            selected_table = self.selected_table_var.get()
            if not selected_table: return
            internal_key = self.table_combobox_map.get(selected_table)
            if not internal_key: return
            if internal_key == self.clipboard_row_data.get("internal_key"):
                table_info = self.potential_tables.get(internal_key)
                parent_element = table_info["parent_element"]
                row_tag = table_info["row_tag"]
                columns = table_info["columns"]
                new_element = ET.Element(row_tag)
                for i, col in enumerate(columns):
                    col_element = ET.SubElement(new_element, col)
                    col_element.text = self.clipboard_row_data["values"][i + 1]
                parent_element.append(new_element)
                self.undo_stack.append({"action": "paste", "element": new_element, "parent_element": parent_element,
                                        "internal_key": internal_key})
                self.redo_stack.clear()
                self._invalidate_table_cache(internal_key)
                self.editmenu.entryconfig("Undo", state="normal")
                self.editmenu.entryconfig("Redo", state="disabled")
                self._refresh_table()
            else:
                messagebox.showerror("Paste Row", "Cannot paste row data into a different table.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste row: {str(e)}\n{traceback.format_exc()[:200]}")

    def delete_row(self):
        try:
            if not self.selected_table_row_id: return
            values = self.table_treeview.item(self.selected_table_row_id, "values")
            tags = self.table_treeview.item(self.selected_table_row_id, "tags")
            if not tags: return
            element_id = int(tags[0])
            selected_table = self.selected_table_var.get()
            internal_key = self.table_combobox_map.get(selected_table)
            if not internal_key: return
            table_info = self.potential_tables.get(internal_key)
            row_elements = table_info["parent_element"].findall(table_info["row_tag"])
            element = None
            for el in row_elements:
                if id(el) == element_id:
                    element = el
                    break
            if not element: return
            parent_element = table_info["parent_element"]
            index = list(parent_element).index(element)
            self.undo_stack.append({
                "action": "delete", "item_id": self.selected_table_row_id, "element": element,
                "parent_element": parent_element, "index": index, "values": values, "internal_key": internal_key
            })
            self.redo_stack.clear()
            self._invalidate_table_cache(internal_key)
            parent_element.remove(element)
            self.table_treeview.delete(self.selected_table_row_id)
            self.selected_table_row_id = None
            self.editmenu.entryconfig("Cut Row", state="disabled")
            self.editmenu.entryconfig("Copy Row", state="disabled")
            self.editmenu.entryconfig("Delete Row", state="disabled")
            self.editmenu.entryconfig("Undo", state="normal")
            self.editmenu.entryconfig("Redo", state="disabled")
            self._refresh_table()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete row: {str(e)}\n{traceback.format_exc()[:200]}")

    def undo_action(self):
        try:
            if not self.undo_stack: return
            action = self.undo_stack.pop()
            internal_key = action.get("internal_key")
            if internal_key:
                self._invalidate_table_cache(internal_key)

            if action["action"] == "edit":
                element = action["element"]
                column = action["column"]
                old_value = action["old_value"]
                col_data = element.find(column)
                if col_data is not None:
                    col_data.text = old_value
                self.redo_stack.append(action)
            elif action["action"] in ["cut", "delete"]:
                parent_element = action["parent_element"]
                element = action["element"]
                index = action["index"]
                parent_element.insert(index, element)
                self.redo_stack.append(action)
            elif action["action"] == "paste":
                parent_element = action["parent_element"]
                element = action["element"]
                parent_element.remove(element)
                self.redo_stack.append(action)

            self._refresh_table()
            self.editmenu.entryconfig("Redo", state="normal")
            if not self.undo_stack:
                self.editmenu.entryconfig("Undo", state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to undo: {str(e)}\n{traceback.format_exc()[:200]}")

    def redo_action(self):
        try:
            if not self.redo_stack: return
            action = self.redo_stack.pop()
            internal_key = action.get("internal_key")
            if internal_key:
                self._invalidate_table_cache(internal_key)

            if action["action"] == "edit":
                element = action["element"]
                column = action["column"]
                new_value = action["new_value"]
                col_data = element.find(column)
                if col_data is not None:
                    col_data.text = new_value
                self.undo_stack.append(action)
            elif action["action"] in ["cut", "delete"]:
                parent_element = action["parent_element"]
                element = action["element"]
                parent_element.remove(element)
                self.undo_stack.append(action)
            elif action["action"] == "paste":
                parent_element = action["parent_element"]
                element = action["element"]
                parent_element.append(element)
                self.undo_stack.append(action)

            self._refresh_table()
            self.editmenu.entryconfig("Undo", state="normal")
            if not self.redo_stack:
                self.editmenu.entryconfig("Redo", state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to redo: {str(e)}\n{traceback.format_exc()[:200]}")

    def _refresh_table(self):
        try:
            selected_table = self.selected_table_var.get()
            if selected_table:
                internal_key = self.table_combobox_map.get(selected_table)
                if internal_key:
                    self.display_table_view_data(internal_key)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh table: {str(e)}\n{traceback.format_exc()[:200]}")

    # --- CORRECTED --- To use cached data for sorting
    def on_table_header_click_for_sort(self, event):
        try:
            column_id_str = self.table_treeview.identify_column(event.x)
            column_id = self.table_treeview.column(column_id_str, "id")
            if column_id == "#": return

            selected_table = self.selected_table_var.get()
            internal_key = self.table_combobox_map.get(selected_table)
            if not internal_key or internal_key not in self.table_data_cache:
                return  # Can't sort if data isn't loaded/cached

            table_info = self.potential_tables.get(internal_key)
            columns = table_info["columns"]

            # Get data from cache instead of re-parsing XML
            data_to_sort = self.table_data_cache[internal_key]

            if self.table_sort_column_id == column_id:
                self.table_sort_direction_is_asc = not self.table_sort_direction_is_asc
            else:
                self.table_sort_direction_is_asc = True
            self.table_sort_column_id = column_id

            # Use sorted() to create a new sorted list
            sorted_data = sorted(
                data_to_sort,
                key=lambda x: str(x.get(column_id, "")).lower(),
                reverse=not self.table_sort_direction_is_asc
            )

            # Repopulate the treeview with the sorted data
            self._populate_table_rows(sorted_data, columns, selected_table)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to sort table: {str(e)}\n{traceback.format_exc()[:200]}")

    def _finish_cell_edit(self, commit=False):
        try:
            if not self.active_cell_editor: return
            editor = self.active_cell_editor
            self.active_cell_editor = None

            item_id = editor.editing_item_id
            col_index = editor.editing_col_index
            column = editor.editing_column_name
            element = getattr(editor, "element", None)
            internal_key = getattr(editor, "internal_key", None)

            if commit and element and column and internal_key:
                new_value = editor.get()
                current_values = list(self.table_treeview.item(item_id, "values"))
                old_value = current_values[col_index]

                if new_value != old_value:
                    if col_index < len(current_values):
                        current_values[col_index] = new_value
                        self.table_treeview.item(item_id, values=current_values)

                    col_data = element.find(column)
                    if col_data is not None:
                        col_data.text = new_value
                    else:  # Handle case where the tag doesn't exist yet
                        col_data = ET.SubElement(element, column)
                        col_data.text = new_value

                    # Invalidate cache and set up undo/redo
                    self._invalidate_table_cache(internal_key)
                    self.undo_stack.append({
                        "action": "edit", "element": element, "internal_key": internal_key,
                        "column": column, "old_value": old_value, "new_value": new_value
                    })
                    self.redo_stack.clear()
                    self.editmenu.entryconfig("Undo", state="normal")
                    self.editmenu.entryconfig("Redo", state="disabled")

            editor.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to finish cell edit: {str(e)}\n{traceback.format_exc()[:200]}")

    def _update_context_menu_state(self, element):
        try:
            if element and len(element) > 0 and element[0].tag is not None:
                if all(child.tag == element[0].tag for child in element):
                    self.context_menu.entryconfig("Export Node as CSV...", state="normal")
                else:
                    self.context_menu.entryconfig("Export Node as CSV...", state="disabled")
            else:
                self.context_menu.entryconfig("Export Node as CSV...", state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update context menu: {str(e)}\n{traceback.format_exc()[:200]}")

    def show_context_menu(self, event):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            item_id = self.tree.identify_row(event.y)
            if item_id:
                if self.tree.selection() != (item_id,):
                    self.tree.selection_set(item_id)
                    self.tree.focus(item_id)
                self._update_context_menu_state(self.tree_item_to_element.get(item_id))
                self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to show context menu: {str(e)}\n{traceback.format_exc()[:200]}")

    def save_xml_as(self):
        try:
            if not self.xml_tree_root: return
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xml",
                filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
                title="Save XML File"
            )
            if not filepath: return
            self.update_status(f"Saving XML to {os.path.basename(filepath)}...", show_progress=True, progress_value=0)
            try:
                tree = ET.ElementTree(self.xml_tree_root)
                tree.write(filepath, encoding='utf-8', xml_declaration=True)
                self.update_status(f"Success: saved to {os.path.basename(filepath)}", show_progress=False)
                messagebox.showinfo("Success", f"Successfully saved to {filepath}")
                self.current_loaded_filepath = filepath
                self.filename_display_var.set(os.path.basename(filepath))
            except Exception as e:
                self.update_status(f"Error saving XML: {str(e)[:50]}...", show_progress=False)
                messagebox.showerror("Error", f"Failed to save XML: {str(e)}\n{traceback.format_exc()[:200]}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save XML: {str(e)}\n{traceback.format_exc()[:200]}")

    def save_current_table_as_csv(self):
        try:
            if self.active_cell_editor:
                self._finish_cell_edit()
            if self.current_right_panel_view != "table":
                messagebox.showerror("Error", "No table displayed")
                return
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Select CSV"
            )
            if not filepath: return
            self.update_status(f"Saving to {os.path.basename(filepath)}...", show_progress=True, progress_value=0)
            try:
                columns = self.table_treeview["columns"]
                with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(columns)
                    items = list(self.table_treeview.get_children())
                    for i, item in enumerate(items):
                        values = self.table_treeview.item(item, "values")
                        writer.writerow(values)
                        progress = ((i + 1) / len(items)) * 100 if len(items) > 0 else 100
                        self.update_status(f"Saving to CSV: ({int(progress)}%)...", True, progress)
                self.update_status(f"Success: saved to {os.path.basename(filepath)}", show_progress=False)
                messagebox.showinfo("Success", f"Successfully saved to {filepath}")
            except Exception as e:
                self.update_status(f"Error saving CSV: {str(e)[:50]}...", show_progress=False)
                messagebox.showerror("Error", f"Failed to save CSV: {str(e)}\n{traceback.format_exc()[:200]}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV: {str(e)}\n{traceback.format_exc()[:200]}")
        finally:
            self.update_status("Ready", show_progress=False)

    def export_node_as_csv(self):
        try:
            selected_element = self.selected_element_for_context_menu
            if not selected_element or len(selected_element) == 0: return
            child_elements = list(selected_element)
            first_child_tag = child_elements[0].tag
            if not all(child.tag == first_child_tag for child in child_elements):
                messagebox.showwarning("Warning", "All child elements must have the same tag.")
                return
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Node as CSV"
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
                    messagebox.showerror("Error", "No valid headers found for CSV export.")
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
                messagebox.showinfo("Success", f"Successfully exported node to {filepath}")
            except Exception as e:
                self.update_status(f"Error exporting CSV: {str(e)[:50]}...", show_progress=False)
                messagebox.showerror("Error", f"Failed to export CSV: {str(e)}\n{traceback.format_exc()[:200]}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export CSV: {str(e)}\n{traceback.format_exc()[:200]}")


if __name__ == "__main__":
    root = tk.Tk()
    app = XMLNotepad(root)
    root.mainloop()
