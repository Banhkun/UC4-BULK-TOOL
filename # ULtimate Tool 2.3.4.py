import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, font, filedialog
import threading
import base64
import automic_rest as automic
import copy
import re
import time
import os
import json
import sys
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
from xml.etree import ElementTree as ET
from datetime import date
from google import genai
from google.genai import types
import logging
from typing import List, Tuple
import tkinter as tk
import keyring
from tkinter import ttk, messagebox, Toplevel
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from requests.exceptions import ConnectionError, HTTPError
from threading import Lock, Thread
# Set environment variables for Gemini API


logging.basicConfig(level=logging.INFO, filename="findbulk.log", format="%(asctime)s - %(levelname)s - %(message)s")
def sanitize_string(s):
    if not isinstance(s, str):
        return ""
    s = re.sub(r'(?i)(^UC4_)|(_UC4$)', '', s)
    s = re.sub(r'(?i)(_UC4_)', '_', s)
    s = re.sub(r'^[^0-9A-Za-z_]+', '', s)
    s = re.sub(r'[^0-9A-Za-z_]', '_', s)  # Replace non-alphanumeric characters with '_'
    s = re.sub(r'_+', '_', s)  # Collapse multiple underscores into a single underscore
    return s
class ToolTip:
    """A class to create tooltips for tkinter widgets."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert") if hasattr(self.widget, "bbox") else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Remove window decorations
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("tahoma", "8", "normal"),
        )
        label.pack(ipadx=1)
        
    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

class JobCreatorApp:
    MAX_WORKERS = 10  # Adjust based on API rate limits
    def __init__(self, parent, env_var, client_var, entries, client_map):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.client_map = client_map
        self.jobs_list = []
        self.jobps_list = []
        self.undo_stack = []
        self.redo_stack = []
        self.connection_lock = Lock()
        self.stop_check = False  # Flag to stop the check process
        self.rv_checked = False  # Flag to track if report/variant pairs have been checked        
        self.logger = logging.getLogger(__name__)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )        
        self.failed_pairs = []  # Store non-existent (program, variant) pairs
        self.current_entry = None  # Track the current Entry widget

        self.build_ui()

    def load_config(self, config):
        """Load tab-specific configuration into UI fields."""
        if 'ARMT_NO' in config:
            self.entries['ARMT_NO'].insert(0, config['ARMT_NO'])
        if 'template_job_armt' in config:
            self.template_job_armt.insert(0, config['template_job_armt'])
        if 'template_joplan_armt' in config:
            self.template_joplan_armt.insert(0, config['template_joplan_armt'])
    def on_action_selected(self, event):
        # Extract only the code (e.g., 'S') from selection like 'S - Skip'
        selected = self.action_combobox.get()
        self.action_var.set(selected[0])  # Just set 'S', 'H', 'X', or 'A'

    def save_config(self):

        return {
            'ARMT_NO': self.entries['ARMT_NO'].get(),
            'template_job_armt': self.template_job_armt.get(),
            'template_joplan_armt': self.template_joplan_armt.get(),
        }

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='ARMT No.:').grid(row=0, column=0, sticky='w')
        self.entries['ARMT_NO'] = ttk.Entry(frm)
        self.entries['ARMT_NO'].grid(row=0, column=1, sticky='ew', padx=5)

        ttk.Label(frm, text='Jobplan Template:').grid(row=1, column=0, sticky='w')
        self.template_joplan_armt = ttk.Entry(frm)
        self.template_joplan_armt.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(frm, text='Jobs Template:').grid(row=2, column=0, sticky='w')
        self.template_job_armt = ttk.Entry(frm)
        self.template_job_armt.grid(row=2, column=1, sticky='ew', padx=5, pady=2)

        self.create_main_var = tk.BooleanVar()
        chk = ttk.Checkbutton(frm, text='Create Main Jobplan', variable=self.create_main_var, command=self.toggle_main_fields)
                # Add tooltip to "Create Main Jobplan" checkbox
        ToolTip(
            chk,
            "Enables creation of a main jobplan that can contain other jobplans or jobs, with optional sequential predecessors."
        )
        chk.grid(row=0, column=2, sticky='w')
        self.is_predecessor_var = tk.BooleanVar()
        self.predecessor_chk = ttk.Checkbutton(frm, text='Use Sequential Predecessors', variable=self.is_predecessor_var, command=self.toggle_main_fields)
        self.predecessor_chk.grid(row=1, column=2, columnspan=2, sticky='w')
        ToolTip(
            self.predecessor_chk,
            "When checked, jobs/jobplans in the main jobplan will run sequentially, with each depending on the previous one."
        )
        self.main_label = ttk.Label(frm, text='Main JOBP Name:')
        self.main_entry = ttk.Entry(frm)
        self.jobp_main_entry = self.main_entry
        self.main_label.grid(row=2, column=2, sticky='w')
        self.main_entry.grid(row=2, column=3, sticky='ew', padx=5)
        self.is_main_jobp_var = tk.BooleanVar()
        self.main_jobp_chk = ttk.Checkbutton(frm, text='Main Contains Jobplans', variable=self.is_main_jobp_var)
        self.main_jobp_chk.grid(row=0, column=3, sticky='w')
        ToolTip(
            self.main_jobp_chk,
            "If checked, the main jobplan will contain other jobplans; otherwise, it will contain jobs."
        )
        

        # Frame to hold both dropdowns in column 3, row 1
        self.dropdown_frame = ttk.Frame(frm)
        self.dropdown_frame.grid(row=1, column=3, sticky='ew', padx=5)
        
        # Adding Condition Type Dropdown
        self.condition_type_var = tk.StringVar()
        self.condition_type_combobox = ttk.Combobox(self.dropdown_frame, textvariable=self.condition_type_var, 
                                                  values=('', 'ENDED_OK', 'ANY_OK', 'ANY_ABEND', 'USER1'), 
                                                  state='readonly', width=5)
        self.condition_type_combobox.set('')
        self.condition_type_combobox.grid(row=0, column=1, sticky='ew', padx=(0, 2))
        # Optional: Add tooltip to Condition Type Combobox
        ToolTip(
            self.condition_type_combobox,
            "Select the condition type for sequential predecessors (e.g., 'ENDED_OK' means the predecessor must complete successfully)."
        )
        # Adding Action Dropdown
        self.action_var = tk.StringVar()
        # Show friendly descriptions in dropdown
        display_values = ['S - Skip', 'H - Block', 'X - Block & Abort', 'A - Abort']
        self.action_combobox = ttk.Combobox(self.dropdown_frame, 
                                    values=display_values,
                                    state='readonly', width=18)
        self.action_combobox.current(0)  # default to 'S - Skip'
        self.action_combobox.grid(row=0, column=3, sticky='ew', padx=(2, 0))
        ToolTip(
            self.action_combobox,
            "Defines the action to take if a predecessor's condition fails (e.g., 'Skip', 'Block', or 'Abort')."
        )
        # Sync dropdown selection to only the code (e.g., 'S')
        self.action_combobox.bind("<<ComboboxSelected>>", self.on_action_selected)

        # Manually trigger for default value
        self.action_var.set(display_values[0][0])

        # Configure dropdown frame columns to split evenly
        self.dropdown_frame.columnconfigure(1, weight=1)
        self.dropdown_frame.columnconfigure(3, weight=1)


        self.tree_frame = ttk.Frame(frm)
        self.tree_frame.grid(row=3, column=0, columnspan=4, sticky='nsew', padx=5)
        self.pairs_tree = ttk.Treeview(self.tree_frame, columns=('jobname', 'program', 'variant', 'login', 'language', 'extra', 'list_t','start_time'), show='headings', height=6)
        for col in ('jobname', 'program', 'variant', 'login', 'language', 'extra', 'list_t','start_time'):
            self.pairs_tree.heading(col, text=col.capitalize())
            self.pairs_tree.column(col, width=150, stretch=True)
        self.pairs_tree.grid(row=0, column=0, sticky='nsew')
        scroll_y = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.pairs_tree.yview)
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.pairs_tree.xview)
        scroll_x.grid(row=1, column=0, sticky='ew')
        self.pairs_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Initialize Treeview with one empty row
        self.ensure_empty_row()
        self.pairs_tree.bind("<Double-1>", self.on_double_click)
        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)
        self.pairs_tree.bind("<Button-3>", self.show_paste_menu)
        self.pairs_tree.bind("<Control-a>", self.select_all)
        self.pairs_tree.bind("<Delete>", self.delete_selected)
        self.pairs_tree.bind("<BackSpace>", self.delete_selected)
        self.pairs_tree.bind("<Control-z>", self.undo)
        self.pairs_tree.bind("<Control-y>", self.redo)
        self.pairs_tree.bind("<Control-c>", self.copy_selected)  # New binding for Ctrl+C
        self.pairs_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='jobname'))  # New binding for Ctrl+V
        self.pairs_tree.bind("<Tab>", self.on_tab_press)  # New binding for Tab key
        self.pairs_tree.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, direction="backward"))
       # Button frame for Run and Check R&V buttons
        # Row 4 container (holds buttons & RV frame)
        self.row4_frame = ttk.Frame(frm)
        self.row4_frame.grid(row=4, column=0, columnspan=4, sticky='ew', pady=12)
        self.row4_frame.grid_columnconfigure(0, weight=1)
        self.row4_frame.grid_columnconfigure(1, weight=0)

        # Left-aligned buttons (Run, SID, CLIENT, LOGIN)
        self.button_frame = ttk.Frame(self.row4_frame)
        self.button_frame.grid(row=0, column=0, sticky='w')

        # Right-aligned CHECK/STOP buttons
        self.rv_button_frame = ttk.Frame(self.row4_frame)
        self.rv_button_frame.grid(row=0, column=1, sticky='w')

        # self.run_btn = ttk.Button(self.button_frame, text='Create Jobs', command=self.start)
        # self.run_btn.grid(row=0, column=0, padx=5, sticky='w')
        small_font = ("Arial",10 )

        self.run_btn = ttk.Button(self.button_frame, text='Create Jobs', command=self.start)
        self.run_btn.grid(row=0, column=0, padx=5, sticky='w')

        self.sid_label = ttk.Label(self.button_frame, text='SID:')
        self.sid_entry = ttk.Entry(self.button_frame, width=5, font=small_font)
        self.sid_label.grid(row=0, column=1, padx=2)
        self.sid_entry.grid(row=0, column=2, padx=2)

        self.client_label = ttk.Label(self.button_frame, text='CLIENT:')
        self.client_entry = ttk.Entry(self.button_frame, width=5, font=small_font)
        self.client_label.grid(row=0, column=3, padx=2)
        self.client_entry.grid(row=0, column=4, padx=2)

        self.login_label = ttk.Label(self.button_frame, text='LOGIN:')
        self.login_entry = ttk.Entry(self.button_frame, width=10, font=small_font)
        self.login_label.grid(row=0, column=5, padx=2)
        self.login_entry.grid(row=0, column=6, padx=2)

        # RV Buttons
        self.check_rv_btn = ttk.Button(self.rv_button_frame, text='CHECK', command=self.check_report_variant)
        self.check_rv_btn.configure(style='Small.TButton')
        self.check_rv_btn.grid(row=0, column=0, padx=(0, 2))

        self.stop_check_rv_btn = ttk.Button(self.rv_button_frame, text='STOP', command=self.stop_check_report_variant)
        self.stop_check_rv_btn.configure(style='Small.TButton')
        self.stop_check_rv_btn.grid(row=0, column=0)
        
        # self.transform_btn = ttk.Button(self.rv_button_frame, text='🪄', command=self.transform_schedules)
        # self.transform_btn.grid(row=0, column=1, padx=(10, 0))

        # Define a smaller style for the buttons
        style = ttk.Style()
        medium_font = ("Arial",12 )

        style.configure('Small.TButton', font=medium_font, padding=(1, 1))
        # Initially hide SID label, entry, and button
        self.sid_label.grid_remove()
        self.sid_entry.grid_remove()
        self.client_label.grid_remove()
        self.client_entry.grid_remove()
        self.login_label.grid_remove()
        self.login_entry.grid_remove()
        self.check_rv_btn.grid_remove()
        self.stop_check_rv_btn.grid_remove()

        # Copy Failed R&V button (initially hidden)
        self.copy_failed_btn = ttk.Button(frm, text='Copy Failed R&V', command=self.copy_failed_pairs)
        self.copy_failed_btn.grid(row=6, column=3, sticky='nsew', padx=5)
        self.copy_failed_btn.grid_remove()


        # ToolTip(
        #     self.transform_btn,
        #     "Transforms schedules in the start_time column using the Gemini API."
        # )
        ttk.Label(frm, text='Output:').grid(row=5, column=0, sticky='nw')
        self.log_box = scrolledtext.ScrolledText(frm, height=10, state='disabled')
        self.log_box.grid(row=5, column=1, columnspan=3, sticky='ew', padx=5)

        self.copy_jobs_btn = ttk.Button(frm, text='Copy JOBS List', command=self.copy_jobs_list)
        self.copy_jobs_btn.grid(row=6, column=1, sticky='nsew',padx = 5)
        self.copy_jobps_btn = ttk.Button(frm, text='Copy JOBP List', command=self.copy_jobps_list)
        self.copy_jobps_btn.grid(row=6, column=2, sticky='nsew',padx = 5)

        frm.rowconfigure(3, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(2, weight=1)
        frm.columnconfigure(3, weight=1)
        self.toggle_main_fields()
        # Bind Treeview update to check for program/variant data
        self.pairs_tree.bind('<<TreeviewSelect>>', self.update_rv_button_visibility)
        self.pairs_tree.bind('<<TreeviewOpen>>', self.update_rv_button_visibility)
        self.pairs_tree.bind('<KeyRelease>', self.update_rv_button_visibility)
    def copy_selected(self, event=None):
        """Copy selected Treeview rows to the clipboard as tab-separated values."""
        selected = self.pairs_tree.selection()
        if not selected:
            self.log("No rows selected to copy.")
            return 'break'

        rows_data = []
        for row_id in selected:
            values = self.pairs_tree.item(row_id)['values']
            # Only copy non-empty rows (skip if all values are empty)
            if any(val.strip() for val in values):
                rows_data.append('\t'.join(str(val) for val in values))

        if rows_data:
            clipboard_text = '\n'.join(rows_data)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(clipboard_text)
            self.parent.update()
            self.log(f"Copied {len(rows_data)} row(s) to clipboard.")
        else:
            self.log("No non-empty rows selected to copy.")

        return 'break'
    def on_double_click(self, event):
        row_id = self.pairs_tree.identify_row(event.y)
        if not row_id:
            return
        column = self.pairs_tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.pairs_tree['columns'][column_num]
        self.edit_cell(row_id, column_name, column_num)

    def edit_cell(self, row_id, column_name, column_num):
        """Open an Entry widget for editing a specific cell."""
        if self.current_entry:
            self.current_entry.destroy()
        x, y, width, height = self.pairs_tree.bbox(row_id, f"#{column_num + 1}")
        self.current_entry = ttk.Entry(self.tree_frame)
        self.current_entry.place(x=x, y=y, width=width, height=height)
        current_value = self.pairs_tree.item(row_id)['values'][column_num]
        self.current_entry.insert(0, current_value)
        self.current_entry.select_range(0, tk.END)
        self.current_entry.focus()
        self.current_entry.bind("<Return>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<FocusOut>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<Escape>", lambda e: self.current_entry.destroy())
        self.current_entry.bind("<Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="forward"))
        self.current_entry.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="backward"))
    def on_tab_press(self, event, current_row_id=None, current_col_num=None, direction="forward"):
        """Handle Tab (forward) and Shift+Tab (backward) key presses to move to the next or previous cell."""
        if current_row_id is None or current_col_num is None:
            # Tab or Shift+Tab pressed directly in Treeview
            selected = self.pairs_tree.selection()
            if not selected:
                return 'break'
            current_row_id = selected[0]
            current_col_num = 0 if direction == "forward" else len(self.pairs_tree['columns']) - 1
        else:
            # Tab or Shift+Tab pressed in Entry widget, save current edit
            self.save_edit(self.current_entry, current_row_id, self.pairs_tree['columns'][current_col_num])
        
        columns = self.pairs_tree['columns']
        children = self.pairs_tree.get_children()
        current_row_index = children.index(current_row_id)
        
        if direction == "forward":
            # Move to next column or next row
            if current_col_num < len(columns) - 1:
                next_col_num = current_col_num + 1
                next_row_id = current_row_id
            else:
                # Move to the first column of the next row
                if current_row_index < len(children) - 1:
                    next_row_id = children[current_row_index + 1]
                    next_col_num = 0
                else:
                    return 'break'  # Stay at last cell
        else:  # direction == "backward"
            # Move to previous column or previous row
            if current_col_num > 0:
                next_col_num = current_col_num - 1
                next_row_id = current_row_id
            else:
                # Move to the last column of the previous row
                if current_row_index > 0:
                    next_row_id = children[current_row_index - 1]
                    next_col_num = len(columns) - 1
                else:
                    return 'break'  # Stay at first cell
        
        next_col_name = columns[next_col_num]
        self.edit_cell(next_row_id, next_col_name, next_col_num)
        return 'break'
    def ensure_empty_row(self):
        """Ensure there is always one empty row at the end of the Treeview."""
        columns = self.pairs_tree['columns']
        children = self.pairs_tree.get_children()
        # Check if the last row is empty (all values are empty strings)
        if children:
            last_row_values = self.pairs_tree.item(children[-1])['values']
            last_row_empty = all(val == '' for val in last_row_values)
            if last_row_empty:
                return  # Last row is already empty, no need to add another
        # Add a new empty row
        empty_row = [''] * len(columns)
        self.pairs_tree.insert('', 'end', values=empty_row)



    def update_rv_button_visibility(self, event=None):
        """Show/hide SID textbox and CHECK R&V button based on program and variant columns."""
        has_program_variant = False
        for row_id in self.pairs_tree.get_children():
            values = self.pairs_tree.item(row_id)['values']
            program = values[1].strip() if len(values) > 1 else ''
            if program:  # Program column has data
                has_program_variant = True
                break
        if has_program_variant:
            self.sid_label.grid(row=0, column=1, padx=2, sticky='e')
            self.sid_entry.grid(row=0, column=2, padx=2, sticky='e')
            self.client_label.grid(row=1, column=1, padx=2, sticky='e')
            self.client_entry.grid(row=1, column=2, padx=0, sticky='e')
            self.login_label.grid(row=1, column=3, padx=0, sticky='w')
            self.login_entry.grid(row=1, column=4, padx=0, sticky='w')
            self.check_rv_btn.grid(row=0, column=3, padx=0,ipady = 8, sticky='w')
            self.stop_check_rv_btn.grid_remove()  # Ensure stop button is hidden initially
        else:
            self.sid_label.grid_remove()
            self.sid_entry.grid_remove()
            self.client_label.grid_remove()
            self.client_entry.grid_remove()
            self.login_label.grid_remove()
            self.login_entry.grid_remove()
            self.check_rv_btn.grid_remove()
            self.stop_check_rv_btn.grid_remove()
            self.copy_failed_btn.grid_remove()
            self.rv_checked = False  # Reset check status when no program/variant data
            self.failed_pairs = []  # Clear failed pairs
            # Reset any red coloring
            for row_id in self.pairs_tree.get_children():
                self.pairs_tree.item(row_id, tags=())

    def stop_check_report_variant(self):
        """Set flag to stop the check process."""
        self.stop_check = True
        self.log("Stopping CHECK R&V process...")

    def check_report_variant(self):
        """Check existence of (program, variant) pairs and color non-existent rows red."""
        self.run_btn.config(state='disabled')
        self.check_rv_btn.grid_remove()
        self.stop_check_rv_btn.grid(row=0, column=3, padx=5, sticky='w')
        self.stop_check = False  # Reset stop flag
        Thread(target=self._check_report_variant_thread, daemon=True).start()

    def _check_report_variant_thread(self):
        """Thread to check (program, variant) pairs existence."""
        try:
            env = self.env_var.get().strip()
            try:
                cid = int(self.client_var.get().strip())
            except ValueError:
                self.log("Error: Invalid Client ID")
                self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
                return

            user = self.entries['USERID'].get().strip()
            pwd = self.entries['PASSWORD'].get().strip()
            sid = self.sid_entry.get().strip()
            client = self.client_entry.get().strip()
            login = self.login_entry.get().strip()

            api_url = f'https://rb-{env}-api.bosch.com'

            if not user or not pwd or not sid or not client:
                self.log("Error: User ID, Password, and SID, CLIENT are required")
                self.parent.after(0, lambda: messagebox.showerror("Error", "Please provide User ID, Password, and SID"))
                return

            auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
            try:
                with self.connection_lock:
                    automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
            except requests.exceptions.HTTPError as e:
                self.log(f"Authentication failed: {str(e)}")
                self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
                return

            # Load test_variants_bulk.json
            try:
                with open('_internal/test_variants_bulk.json', 'r') as file:
                    loaded_data = json.load(file)
            except FileNotFoundError:
                self.log("Error: 'test_variants_bulk.json' not found!")
                self.parent.after(0, lambda: messagebox.showerror("Error", "'test_variants_bulk.json' not found!"))
                return
            except json.JSONDecodeError:
                self.log("Error: Invalid JSON format in 'test_variants_bulk.json'!")
                self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid JSON format in 'test_variants_bulk.json'!"))
                return

            # Validate JSON structure
            try:
                process_list = loaded_data['data']['jobs']['scripts'][0]['process']
                if not isinstance(process_list, list):
                    self.log("Error: 'process' field is not a list!")
                    self.parent.after(0, lambda: messagebox.showerror("Error", "'process' field is not a list!"))
                    return
            except KeyError:
                self.log("Error: Invalid JSON structure! Expected 'data.jobs.scripts[0].process'.")
                self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid JSON structure! Expected 'data.jobs.scripts[0].process'."))
                return

            # Modify process_list[1] to use SID
            try:
                if len(process_list) > 1:
                    process_list[1] = f':put_att HOST      ="SAP{sid}"'
                    if cid == 1111:
                        edit_login = login if login else "SY-BATCH"
                        process_list[0] = f':put_att LOGIN      ="LOGIN_R3_{client}_{edit_login}"'
                    else:
                        edit_login = login if login else "UC4CPIC"
                        process_list[0] = f':put_att LOGIN      ="LOGIN_R3_{client}_{edit_login}"'

                else:
                    self.log("Error: process_list does not have enough entries to modify index 1!")
                    self.parent.after(0, lambda: messagebox.showerror("Error", "process_list does not have enough entries to modify index 1!"))
                    return
            except IndexError:
                self.log("Error: process_list index 1 is out of range!")
                self.parent.after(0, lambda: messagebox.showerror("Error", "process_list index 1 is out of range!"))
                return

            # Collect program and variant pairs from Treeview
            pairs_data = []
            row_ids = []
            for row_id in self.pairs_tree.get_children():
                values = self.pairs_tree.item(row_id)['values']
                program = values[1].strip() if len(values) > 1 else ''
                variant = values[2].strip() if len(values) > 2 else ''
                if program:  # Only include rows with a program
                    pairs_data.append((program, variant))
                    row_ids.append(row_id)

            if not pairs_data:
                self.log("No program/variant pairs to check.")
                self.parent.after(0, lambda: messagebox.showinfo("Info", "No program/variant pairs to check."))
                return

            # Append unique REP values to process_list
            rep_values = []
            seen_reps = set()
            for rep, _ in pairs_data:
                if rep not in seen_reps:
                    rep_values.append(rep)
                    seen_reps.add(rep)
            for rep in rep_values:
                new_line = f'R3_GET_VARIANTS REP = "{rep}",ERROR="IGNORE"'
                process_list.append(new_line)

            # Add client field to loaded_data
            loaded_data['client'] = cid

            # Post the updated JSON
            try:
                with self.connection_lock:
                    res = automic.postObjects(client_id=cid, body=loaded_data, query="overwrite_existing_objects=true")
                self.log(f"Post API Response: {res.status}")
            except Exception as e:
                self.log(f"Error posting data to Automic API: {str(e)}")
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Error posting data to Automic API: {str(e)}"))
                return

            # Execute the job
            body = {
                "object_name": "TEST_VARIANTS_BUILK",
                "execution_option": "execute",
            }
            try:
                with self.connection_lock:
                    res = automic.executeObject(client_id=cid, body=body)
                runid = res.response['run_id']
                self.log(f"Job executed with RunID: {runid}")
            except Exception as e:
                self.log(f"Error executing object: {str(e)}")
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Error executing object: {str(e)}"))
                return

            # Wait until execution is complete (status 1900)
            def get_uc_status(client, runid):
                try:
                    res = automic.getExecution(client_id=client, run_id=runid)
                    return res.response['status']
                except Exception as e:
                    self.log(f"Error checking execution status: {str(e)}")
                    return None

            start_time = time.time()
            timeout = 60  # 60 seconds timeout
            while get_uc_status(cid, runid) != 1900:
                if self.stop_check:
                    self.log("CHECK R&V process stopped by user.")
                    self.parent.after(0, lambda: messagebox.showinfo("Info", "CHECK R&V process was stopped."))
                    return
                if time.time() - start_time > timeout:
                    self.log("Error: Timeout waiting for job execution to complete (status 1900).")
                    self.parent.after(0, lambda: messagebox.showerror("Error", "Timeout waiting for job execution to complete."))
                    return
                time.sleep(3)

            # Retrieve report content
            try:
                with self.connection_lock:
                    res = automic.listReportContent(client_id=cid, run_id=runid, report_type='PLOG', query="max_results=5")
                report_content = res.response['data'][0]['content']
                self.log("Report retrieved successfully")
            except Exception as e:
                self.log(f"Error retrieving report content: {str(e)}")
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Error retrieving report content: {str(e)}"))
                return

            # Check existence of program/variant pairs
            self.failed_pairs = []
            results = []
            seen_pairs = set()
            for (rep, var), row_id in zip(pairs_data, row_ids):
                if (rep, var) not in seen_pairs:
                    seen_pairs.add((rep, var))
                    rep_escaped = re.escape(rep)
                    var_escaped = re.escape(var)
                    pattern = f';{rep_escaped};{var_escaped}\\b'
                    exists = bool(re.search(pattern, report_content))
                    results.append({
                        'REP': rep,
                        'VAR': var,
                        'Status': 'exists' if exists else 'non-exist',
                        'row_id': row_id
                    })
                    if not exists:
                        self.failed_pairs.append(f"{rep}\t{var}")
                        self.pairs_tree.item(row_id, tags=('non_exist',))
                    else:
                        self.pairs_tree.item(row_id, tags=())

            # Define red tag for non-existent pairs
            self.pairs_tree.tag_configure('non_exist', background='red')

            # Show Copy Failed R&V button if there are failed pairs
            if self.failed_pairs:
                self.copy_failed_btn.grid(row=6, column=3, sticky='nsew', padx=5)
            else:
                self.copy_failed_btn.grid_remove()

            # Log results
            self.log("REP and VAR pair existence check results:")
            for result in results:
                self.log(f"REP: {result['REP']}, VAR: {result['VAR']}, Status: {result['Status']}")
            # Set check status to True if check completes successfully
            self.rv_checked = True

        except Exception as e:
            self.log(f"Unexpected error in check_report_variant: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Error", f"Unexpected error: {str(e)}"))
        finally:
            self.parent.after(0, lambda: self.run_btn.config(state='normal'))
            self.parent.after(0, lambda: self.check_rv_btn.grid(row=0, column=3, padx=5, sticky='e'))
            self.parent.after(0, lambda: self.stop_check_rv_btn.grid_remove())

    def copy_failed_pairs(self):
        """Copy non-existent program/variant pairs to clipboard."""
        if not self.failed_pairs:
            self.log("No failed program/variant pairs to copy.")
            self.parent.after(0, lambda: messagebox.showinfo("Info", "No failed program/variant pairs to copy."))
            return
        pairs_text = "\n".join(self.failed_pairs)
        self.parent.clipboard_clear()
        self.parent.clipboard_append(pairs_text)
        self.parent.update()
        self.log("Copied failed program/variant pairs to clipboard.")

    def toggle_main_fields(self):
        if self.create_main_var.get():
            self.main_label.grid()
            self.main_entry.grid()
            self.predecessor_chk.grid()
            self.main_jobp_chk.grid()
            if self.is_predecessor_var.get():
                self.condition_type_combobox.grid()
                self.action_combobox.grid()
            else:
                self.condition_type_combobox.grid_remove()
                self.action_combobox.grid_remove()
        else:
            self.main_label.grid_remove()
            self.main_entry.grid_remove()
            self.predecessor_chk.grid_remove()
            self.main_jobp_chk.grid_remove()
            self.condition_type_combobox.grid_remove()
            self.action_combobox.grid_remove()
            self.is_main_jobp_var.set(True)

    def save_schedule_state(self):
        """Save the current state of the start_time column for undo."""
        return [(child, self.pairs_tree.item(child)['values'][7]) for child in self.pairs_tree.get_children()]


    def save_edit(self, entry, row_id, column_name):
        new_value = entry.get().strip()
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.pairs_tree.set(row_id, column_name, new_value)
        # Reset rv_checked if program or variant column is edited
        if column_name in ('program', 'variant'):
            self.rv_checked = False
        entry.destroy()
        self.update_rv_button_visibility()
        self.ensure_empty_row()  # Ensure an empty row after pasting

    def show_paste_menu(self, event):
        self.clicked_column = self.pairs_tree.identify_column(event.x)[1:]
        self.paste_menu.tk_popup(event.x_root, event.y_root)

    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()

        columns = self.pairs_tree['columns']
        if start_column is None:
            if hasattr(self, 'clicked_column') and self.clicked_column:
                col_index = int(self.clicked_column) - 1
                start_column = columns[col_index]
            else:
                start_column = 'jobname'

        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return

        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return

            existing_rows = list(self.pairs_tree.get_children())
            max_columns = len(columns) - start_col_index

            for i, line in enumerate(lines):
                clipboard_values = line.split('\t') if '\t' in line else [line]
                clipboard_values = [val.strip() for val in clipboard_values[:max_columns]]
                while len(clipboard_values) < max_columns:
                    clipboard_values.append('')
                
                if i < len(existing_rows):
                    row_id = existing_rows[i]
                    current_values = list(self.pairs_tree.item(row_id)['values'])
                    for j, value in enumerate(clipboard_values):
                        current_values[start_col_index + j] = value
                    self.pairs_tree.item(row_id, values=current_values)
                else:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(clipboard_values):
                        new_values[start_col_index + j] = value
                    self.pairs_tree.insert('', 'end', values=new_values)
            # Reset rv_checked if pasting into program or variant columns
            if start_column in ('program', 'variant') or start_col_index <= columns.index('variant'):
                self.rv_checked = False
            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
            self.update_rv_button_visibility()
            self.ensure_empty_row()  # Ensure an empty row after pasting
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()

    def select_all(self, event):
        self.pairs_tree.selection_set(self.pairs_tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.pairs_tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.pairs_tree.delete(item)
            self.rv_checked = False
            self.update_rv_button_visibility()
            self.ensure_empty_row()  # Ensure an empty row after deletion
        return 'break'

    def save_state(self):
        return [self.pairs_tree.item(child)['values'] for child in self.pairs_tree.get_children()]

    def restore_state(self, state):
        self.pairs_tree.delete(*self.pairs_tree.get_children())
        for values in state:
            self.pairs_tree.insert('', 'end', values=values)
        # Reset rv_checked since table data has changed
        self.rv_checked = False
        self.update_rv_button_visibility()
        self.ensure_empty_row()  # Ensure an empty row after restoring state

    def undo(self, event):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'



    def copy_jobs_list(self):
        if not self.jobs_list:
            self.log("No jobs list available to copy.")
            messagebox.showinfo("Info", "No jobs list available to copy.")
            return
        jobs_text = "\n".join(self.jobs_list)
        self.parent.clipboard_clear()
        self.parent.clipboard_append(jobs_text)
        self.parent.update()
        self.log("Copied JOBS list to clipboard.")

    def copy_jobps_list(self):
        if not self.jobps_list:
            self.log("No job plans list available to copy.")
            messagebox.showinfo("Info", "No job plans list available to copy.")
            return
        jobps_text = "\n".join(self.jobps_list)
        self.parent.clipboard_clear()
        self.parent.clipboard_append(jobps_text)
        self.parent.update()
        self.log("Copied JOBP list to clipboard.")

    def log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
        self.parent.update_idletasks()

    def start(self):
        """Check if report/variant pairs have been validated before creating jobs."""
        has_program_variant = False
        for row_id in self.pairs_tree.get_children():
            values = self.pairs_tree.item(row_id)['values']
            program = values[1].strip() if len(values) > 1 else ''
            if program:  # Program column has data
                has_program_variant = True
                break

        # If program/variant data exists and hasn't been checked, show popup
        if has_program_variant and not self.rv_checked:
            dialog = tk.Toplevel(self.parent)
            dialog.title("Validation Required")
            dialog.geometry("300x150")
            dialog.transient(self.parent)
            dialog.grab_set()

            ttk.Label(dialog, text="Hey handsome guy! \n Pls check the Program and Variant first").pack(pady=20)
            
            btn_frame = ttk.Frame(dialog)
            btn_frame.pack(pady=10)
            
            ttk.Button(btn_frame, text="Close", command=dialog.destroy).pack(side='left', padx=5)
            ttk.Button(btn_frame, text="Create anyway", command=lambda: [dialog.destroy(), self.execute_thread()]).pack(side='left', padx=5)
            
            # Center dialog relative to parent
            dialog.update_idletasks()
            x = self.parent.winfo_rootx() + (self.parent.winfo_width() - dialog.winfo_width()) // 2
            y = self.parent.winfo_rooty() + (self.parent.winfo_height() - dialog.winfo_height()) // 2

            dialog.geometry(f"+{x}+{y}")
        else:
            self.execute_thread()

    def execute_thread(self):
        """Start the job creation process in a separate thread."""
        self.run_btn.config(state='disabled')
        Thread(target=self.execute, daemon=True).start()    
    def create_job(self, p, cid, user, armt, tmpl_jobs, base_jobs, default_login):
        """Create a single job and return its name and status."""
        jn = p['jobname'].upper()
        name_jobs = f"{base_jobs}_{jn}"
        login_val = f"{'_'.join(default_login.split('_')[:3])}_{p['login']}" if p.get('login') else default_login

        if cid == 1111:
            script = (
                [f":INC BSH_XXXX_INC_MIGRATION_SIMULATION WAIT_TIME = \"<Random number of seconds between 1 to 60>\" ,NOFOUND=IGNORE"]
                + ([f':PUT_ATT SAP_LANG="{p["language"]}"'] if p.get('language') else [])
                + [
                    f":PUT_ATT JOB_NAME= \"{jn}\"",
                    f":PUT_ATT LOGIN='{login_val}'",
                    f"R3_ACTIVATE_REPORT REPORT='{p['program']}',VARIANT='{p['variant']}',COPIES=1,EXPIR=8,LINE_COUNT=65,LINE_SIZE=80,LAYOUT=X_FORMAT,DATA_SET=LIST1S,TYPE=TEXT"
                ]
            )
        else:
            script = (
                ([f':PUT_ATT JOB_NAME= "{jn}"'] if not p.get('isBSH') else [])
                + ([f':PUT_ATT SAP_LANG="{p["language"]}"'] if p.get('language') else [])
                + [f"R3_ACTIVATE_REPORT REPORT='{p['program']}',VARIANT='{p['variant']}'"]
            )
        if p.get('extra'):
            script[-1] = script[-1] + "," + p['extra']
        if p.get('list_t'):
            script[-1] = script[-1] + "," + "LIST_T=" + p['list_t']
        
        nj = copy.deepcopy(tmpl_jobs)
        nj['general_attributes']['name'] = name_jobs
        for proc in nj.get('scripts', []):
            if 'process' in proc:
                proc['process'] = script
        
        try:
            with self.connection_lock:
                res_j = automic.postObjects(
                    client_id=cid,
                    body={'total': 1, 'data': {'jobs': nj}, 'path': f'AUTOMATION_JOBS/{user}/{armt}', 'client': cid, 'hasmore': False},
                )
            status = None if res_j.status is None else res_j.status
            self.logger.info(f"Created job {name_jobs}, status: {status}")
            return name_jobs, status, p.get('start_time_info', {})
        except requests.exceptions.HTTPError as e:
            self.logger.error(f"HTTP error creating job {name_jobs}: {str(e)}")
            return name_jobs, str(e), p.get('start_time_info', {})

    def create_jobplan(self, p, cid, user, armt, tmpl_jobp, base_jobp, tmpl_jobs, base_jobs):
        """Create a single jobplan and return its name and status."""
        jn = p['jobname'].upper()
        name_jobp = f"{base_jobp}_{jn}"
        njp = copy.deepcopy(tmpl_jobp)
        njp['general_attributes']['name'] = name_jobp
        for wf in njp.get('workflow_definitions', []):
            if wf.get('object_name') == tmpl_jobs['general_attributes']['name']:
                wf['object_name'] = f"{base_jobs}_{jn}"
        
        try:
            with self.connection_lock:
                res_p = automic.postObjects(
                    client_id=cid,
                    body={'total': 1, 'data': {'jobp': njp}, 'path': f'AUTOMATION_JOBS/{user}/{armt}', 'client': cid, 'hasmore': False},
                )
            status = None if res_p.status is None else res_p.status
            self.logger.info(f"Created jobplan {name_jobp}, status: {status}")
            return name_jobp, status, p.get('start_time_info', {})
        except requests.exceptions.HTTPError as e:
            self.logger.error(f"HTTP error creating jobplan {name_jobp}: {str(e)}")
            return name_jobp, str(e), p.get('start_time_info', {})
    def execute(self):
        try:
            self.jobs_list = []
            self.jobps_list = []
            env = self.env_var.get().strip()
            try:
                cid = int(self.client_var.get().strip())
            except ValueError:
                self.log("Error: Invalid Client ID")
                messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value.")
                return

            user = self.entries['USERID'].get().strip()
            pwd = self.entries['PASSWORD'].get().strip()
            armt = self.entries['ARMT_NO'].get().strip()
            api_url = f'https://rb-{env}-api.bosch.com'
            t_job = self.template_job_armt.get().strip().upper()
            t_joplan = self.template_joplan_armt.get().strip().upper()
            rows = [self.pairs_tree.item(child)['values'] for child in self.pairs_tree.get_children()]
            rows = [row for row in rows if any(val.strip() for val in row)]
            create_main = self.create_main_var.get()
            main_name = self.jobp_main_entry.get().strip()

            is_main_jobp = self.is_main_jobp_var.get()

            if not user or not pwd:
                self.parent.after(0, lambda: self.log("Error: User ID and Password are required"))
                self.parent.after(0, lambda: messagebox.showerror("Error", "Please provide both User ID and Password"))
                return

            auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
            try:
                with self.connection_lock:
                    automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
            except requests.exceptions.HTTPError as e:
                self.parent.after(0, lambda: self.log(f"Authentication failed: {str(e)}"))
                self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
                return

            tmpl_jobp = None
            base_jobp = ''
            if t_joplan:
                self.parent.after(0, lambda: self.log(f"Fetching jobplan {t_joplan}"))
                try:
                    rp = automic.getObjects(client_id=cid, object_name=t_joplan)
                    if rp.status != 200:
                        self.parent.after(0, lambda: self.log(f"Failed to fetch jobplan {t_joplan}: {rp.status}"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", f"Failed to fetch jobplan {t_joplan}: {rp.status}"))
                        return
                    if 'data' not in rp.response or 'jobp' not in rp.response['data']:
                        self.parent.after(0, lambda: self.log(f"Error: Jobplan {t_joplan} not found or invalid response"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", f"Jobplan {t_joplan} not found or invalid response from server"))
                        return
                    tmpl_jobp = rp.response['data']['jobp']
                    if cid == 1111:
                        base_jobp = tmpl_jobp['general_attributes']['name'][:31]
                    else:
                        name_parts = tmpl_jobp['general_attributes']['name'].split('_')
                        base_jobp = '_'.join(name_parts[:5])
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error fetching jobplan {t_joplan}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to fetch jobplan {t_joplan}: {str(e)}"))
                    return
            tmpl_jobs = None
            base_jobs = ''
            if t_job:
                try:
                    rj = automic.getObjects(client_id=cid, object_name=t_job)
                    if rj.status != 200:
                        self.parent.after(0, lambda: self.log(f"Failed to fetch job {t_job}: {rj.status}"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", f"Failed to fetch job {t_job}: {rj.status}"))
                        return
                    if 'data' not in rj.response or 'jobs' not in rj.response['data']:
                        self.parent.after(0, lambda: self.log(f"Error: Job {t_job} not found or invalid response"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", f"Job {t_job} not found or invalid response from server"))
                        return
                    tmpl_jobs = rj.response['data']['jobs']
                    if cid == 1111:
                        base_jobs = tmpl_jobs['general_attributes']['name'][:21]
                    else:
                        name_parts = tmpl_jobs['general_attributes']['name'].split('_')
                        base_jobs = '_'.join(name_parts[:3])
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error fetching job {t_job}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to fetch job {t_job}: {str(e)}"))
                    return
                # default_login = extract_default_login(tmpl_jobs)

            
            # Initialize dictionary to store start times
            instances = []  # List of dicts: {'raw_name': name, 'object_name': object_name, 'start_time_info': start_time_info}
            # Refined parsing logic
            pairs = []
            for row in rows:
                values = [val.strip() for val in row]
                pair = {}
                if values[0] and all(not v for v in values[1:7]):
                    pair = {"jobp": values[0]}

                elif not values[0] and values[1] and values[2]:
                    program = values[1]
                    variant = values[2]
                    if cid == 1111:
                        jobname = f"C_{sanitize_string(program)}_{sanitize_string(variant)}"
                    else:
                        jobname = f"{sanitize_string(program)}_{sanitize_string(variant)}"
                    pair = {"jobname": jobname, "program": program, "variant": variant, "isBSH": True}
                    if len(values) > 7 and values[7]:
                        pair["start_time"] = values[7]
                elif values[0] and values[1] and values[2]:
                    pair = {"jobname": values[0], "program": values[1], "variant": values[2]}
                    if len(values) > 3 and values[3]:
                        pair["login"] = values[3]
                    if len(values) > 4 and values[4]:
                        pair["language"] = values[4]
                    if len(values) > 5 and values[5]:
                        pair["extra"] = values[5]
                    if len(values) > 6 and values[6]:
                        pair["list_t"] = values[6]
                start_time_info = {}
                if len(values) > 7 and values[7]:
                    start_time_str = values[7]
                    parts = [p.strip() for p in start_time_str.split(',')]
                    if len(parts) == 1:
                        start_time_info = {"start_time": parts[0]}
                    elif len(parts) == 2:
                        start_time_info = {
                            "calendar": parts[0],
                            "calendar_event": parts[1],
                        }
                    elif len(parts) == 3:
                        start_time_info = {
                            "calendar": parts[0],
                            "calendar_event": parts[1],
                            "start_time": parts[2]
                        }
                    elif len(parts) == 4:
                        start_time_info = {
                            "calendar": parts[0],
                            "calendar_event": parts[1],
                            "start_time": parts[2],
                            "timezone": parts[3]
                        }
                    else:
                        self.log(f"Warning: Invalid start_time format: {start_time_str}")
                pair["start_time_info"] = start_time_info
                pairs.append(pair)
            if pairs and pairs[0].get("jobp"):
                t_joplan = pairs[0]["jobp"]
                self.parent.after(0, lambda: self.log(f"Fetching {t_joplan}"))
                try:
                    rp = automic.getObjects(client_id=cid, object_name=t_joplan)
                    if rp.status != 200:
                        self.parent.after(0, lambda: self.log(f"Failed to fetch {t_joplan}: {rp.status}"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", f"Failed to fetch jobplan {t_joplan}: {rp.status}"))
                        return
                    if 'data' not in rp.response:
                        self.parent.after(0, lambda: self.log(f"Error: Jobplan {t_joplan} not found or invalid response"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", f"Jobplan {t_joplan} not found or invalid response from server"))
                        return
                    if 'jobp' in rp.response['data']:
                        tmpl_jobp = rp.response['data']['jobp']
                        for p in pairs:
                            jobp_name = p['jobp']
                            self.jobps_list.append(jobp_name)
                            instances.append({
                                'raw_name': jobp_name,
                                'object_name': jobp_name,
                                'start_time_info': p.get('start_time_info', {})
                            })
                    else:
                        for p in pairs:
                            job_name = p['jobp']
                            self.jobs_list.append(job_name)
                            instances.append({
                                'raw_name': job_name,
                                'object_name': job_name,
                                'start_time_info': p.get('start_time_info', {})
                            })
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error fetching jobplan {t_joplan}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to fetch jobplan {t_joplan}: {str(e)}"))
                    return
            else:
                # Parallel processing for job and jobplan creation
                default_login = extract_default_login(tmpl_jobs) if tmpl_jobs else ''
                with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
                    future_to_obj = {}
                    for p in pairs:
                        jn = p['jobname'].upper()
                        future = executor.submit(self.create_job, p, cid, user, armt, tmpl_jobs, base_jobs, default_login)
                        future_to_obj[future] = (p, 'job')
                        if not create_main:
                            future = executor.submit(self.create_jobplan, p, cid, user, armt, tmpl_jobp, base_jobp, tmpl_jobs, base_jobs)
                            future_to_obj[future] = (p, 'jobplan')
                        elif tmpl_jobp and is_main_jobp:
                            future = executor.submit(self.create_jobplan, p, cid, user, armt, tmpl_jobp, base_jobp, tmpl_jobs, base_jobs)
                            future_to_obj[future] = (p, 'jobplan')


                    for future in as_completed(future_to_obj):
                        p, obj_type = future_to_obj[future]
                        try:
                            name, status, start_time_info = future.result()
                            if obj_type == 'jobplan':
                                self.jobps_list.append(name)
                                self.parent.after(0, lambda n=name, s=status: self.log(f"JOBP: {n}" if s is None else f"FAIL JOBP: {n} ({s})"))
                                instances.append({
                                    'raw_name': p['jobname'].upper(),
                                    'object_name': name,
                                    'start_time_info': start_time_info
                                })
                            else:
                                self.jobs_list.append(name)
                                self.parent.after(0, lambda n=name, s=status: self.log(f"JOBS: {n}" if s is None else f"FAIL JOBS: {n} ({s})"))
                                if not is_main_jobp:
                                    instances.append({
                                        'raw_name': p['jobname'].upper(),
                                        'object_name': name,
                                        'start_time_info': start_time_info
                                    })
                        except Exception as e:
                            self.logger.error(f"Error processing {obj_type} {p['jobname']}: {str(e)}")
                            self.parent.after(0, lambda: self.log(f"Error processing {obj_type} {p['jobname']}: {str(e)}"))

            # Main jobplan creation (sequential, as it depends on jobs/jobplans)
            is_predecessor_var = self.is_predecessor_var.get()
            action_var = self.action_var.get()
            condition_type_var = self.condition_type_var.get()
            if create_main and main_name and tmpl_jobp:
                data = copy.deepcopy(tmpl_jobp)
                start_node = next(obj for obj in data['workflow_definitions'] if obj['object_type'] == '<START>')
                end_node = next(obj for obj in data['workflow_definitions'] if obj['object_type'] == '<END>')
                new_defs = [start_node]
                new_calendar_conditions = []
                line_no = 2
                if is_main_jobp:
                    for key in ['condition_values', 'condition_parameters', 'conditions']:
                        data.pop(key, None)
                    for instance in instances:
                        new_node = {
                            'line_number': line_no,
                            'object_type': 'JOBP',
                            'object_name': instance['object_name'],
                            'precondition_error_action': 'H',
                            'active': 1,
                            'mrt_time': '000000',
                            'childflags': '0000000000000000',
                            'rollback_enabled': 1
                        }
                        start_info = instance['start_time_info']
                        if start_info and "start_time" in start_info:
                            new_node['earliest_start_time'] = start_info["start_time"]
                        if start_info and "timezone" in start_info:
                            new_node['timezone_ert'] = start_info["timezone"]
                        if start_info and "calendar" in start_info and "calendar_event" in start_info:
                            new_node['calendar_condition_type'] = '2'
                            new_cal_cond = {
                                'workflow_line_number': line_no,
                                'line_number': 1,
                                'calendar': start_info["calendar"],
                                'calendar_event': start_info["calendar_event"]
                            }
                            new_calendar_conditions.append(new_cal_cond)
                        if is_predecessor_var:
                            new_node['predecessors'] = 1
                            new_node['row'] = 1
                            new_node['column'] = line_no
                            if action_var:
                                new_node['precondition_error_action'] = action_var
                        else:
                            new_node.pop('predecessors', None)
                            new_node['row'] = line_no - 1
                            new_node['column'] = 2
                        new_defs.append(new_node)
                        line_no += 1
                else:
                    try:
                        jobs_template_node = next(obj for obj in data['workflow_definitions'] if obj['object_type'] == 'JOBS')
                    except StopIteration:
                        raise ValueError("JOBS node not found in workflow_definitions when is_main_jobp is False")
                    for instance in instances:
                        new_node = copy.deepcopy(jobs_template_node)
                        new_node.update({
                            'line_number': line_no,
                            'object_name': instance['object_name'],
                            'row': 1 if is_predecessor_var else line_no - 1,
                            'column': line_no if is_predecessor_var else 2
                        })
                        start_info = instance['start_time_info']
                        if start_info and "start_time" in start_info:
                            new_node['earliest_start_time'] = start_info["start_time"]
                        if start_info and "timezone" in start_info:
                            new_node['timezone_ert'] = start_info["timezone"]
                        if start_info and "calendar" in start_info and "calendar_event" in start_info:
                            new_node['calendar_condition_type'] = '2'
                            new_cal_cond = {
                                'workflow_line_number': line_no,
                                'line_number': 1,
                                'calendar': start_info["calendar"],
                                'calendar_event': start_info["calendar_event"]
                            }
                            new_calendar_conditions.append(new_cal_cond)
                        if is_predecessor_var:
                            new_node['predecessors'] = 1
                            if action_var:
                                new_node['precondition_error_action'] = action_var
                        else:
                            new_node.pop('predecessors', None)
                        new_defs.append(new_node)
                        line_no += 1
                    condition_template = [cond for cond in data['conditions'] if cond['workflow_line_number'] == 2]
                    if not condition_template:
                        self.parent.after(0, lambda: self.log("Error: No conditions found for workflow_line_number 2"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", "No conditions found for workflow_line_number 2"))
                        return                    
                    new_conditions = []
                    for wf_line in range(3, line_no):
                        for cond in condition_template:
                            new_cond = copy.deepcopy(cond)
                            new_cond['workflow_line_number'] = wf_line
                            new_conditions.append(new_cond)
                    data['conditions'].extend(new_conditions)
                    condition_value_template = [cond for cond in data['condition_values'] if cond['workflow_line_number'] == 2]
                    if not condition_value_template:
                        self.parent.after(0, lambda: self.log("Error: No condition_values found for workflow_line_number 2"))
                        self.parent.after(0, lambda: messagebox.showerror("Error", "No condition_values found for workflow_line_number 2"))
                        return
                    new_conditions_value = []
                    for wf_line in range(3, line_no):
                        for cond in condition_value_template:
                            new_cond = copy.deepcopy(cond)
                            new_cond['workflow_line_number'] = wf_line
                            new_conditions_value.append(new_cond)
                    data['condition_values'].extend(new_conditions_value)
                    if cid == 1111:
                        condition_parameter_template = [cond for cond in data['condition_parameters'] if cond['workflow_line_number'] == 2]
                        if not condition_parameter_template:
                            self.parent.after(0, lambda: self.log("Error: No condition_parameters found for workflow_line_number 2"))
                            self.parent.after(0, lambda: messagebox.showerror("Error", "No condition_parameters found for workflow_line_number 2"))
                            return
                        new_conditions_parameter = []
                        for wf_line in range(3, line_no):
                            for cond in condition_parameter_template:
                                new_cond = copy.deepcopy(cond)
                                new_cond['workflow_line_number'] = wf_line
                                new_conditions_parameter.append(new_cond)
                        data['condition_parameters'].extend(new_conditions_parameter)

                end_node['line_number'] = line_no
                end_node['row'] = 1
                if is_predecessor_var:
                    end_node['column'] = line_no
                    end_node['predecessors'] = 1
                else:
                    end_node['column'] = 3
                    end_node.pop('predecessors', None)
                new_defs.append(end_node)
                data['workflow_definitions'] = new_defs
                if new_calendar_conditions:
                    data['calendar_conditions'] = new_calendar_conditions
                def gen_conditions(defs):
                    line_conds = []
                    if is_predecessor_var:
                        for node in defs:
                            ln = node['line_number']
                            if 'predecessors' in node:
                                preds = node.get('predecessors', [])
                                preds = [ln - 1]
                                for idx, p in enumerate(preds, 1):
                                    line_condition = {'workflow_line_number': ln, 'line_number': idx, 'predecessor_line_number': p}
                                    if condition_type_var:
                                        line_condition["ok_status"] = condition_type_var
                                    line_conds.append(line_condition)
                    return line_conds

                data['line_conditions'] = gen_conditions(new_defs)
                data['general_attributes']['name'] = main_name
                data['general_attributes']['workflow_children'] = line_no

                body = {'total': 1, 'data': {'jobp': data}, 'path': f'AUTOMATION_JOBS/{user}/{armt}', 'client': cid, 'hasmore': False}
                # print(body)
                try:
                    with self.connection_lock:
                        resp_main = automic.postObjects(client_id=cid, body=body)
                    self.parent.after(0, lambda: self.log(f"MAIN JOBP: {main_name}" if resp_main.status is None else f"FAIL MAIN JOBP: {main_name} ({resp_main.status})"))
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error creating main jobplan {main_name}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to create main jobplan {main_name}: {str(e)}"))

            self.parent.after(0, lambda: self.log("All done."))

        except Exception as e:
            error_msg = str(e)
            self.parent.after(0, lambda: self.log(f"Unexpected error: {error_msg}"))
            self.parent.after(0, lambda: messagebox.showerror("Error", f"An unexpected error occurred: {error_msg}"))
        finally:
            self.parent.after(0, lambda: self.run_btn.config(state='normal'))

    def transform_schedules(self):
        """Transform schedules in the start_time column using Gemini API and update the Treeview."""
        self.undo_stack.append(self.save_schedule_state())
        self.redo_stack.clear()
        os.environ['GEMINI_API_KEY'] = ''
        os.environ['HTTPS_PROXY'] = 'http://127.0.0.1:8008'
        os.environ['HTTP_PROXY'] = 'http://127.0.0.1:8008'
        try:
            # Initialize Gemini client
            client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
            model = "gemini-2.0-flash"

            # Collect and clean schedules from start_time column
            schedules = []
            row_ids = []
            for row_id in self.pairs_tree.get_children():
                values = self.pairs_tree.item(row_id)['values']
                start_time = values[7] if len(values) > 7 else ""
                # Clean non-breaking spaces, normalize whitespace, and handle trailing commas
                start_time = start_time.replace('\xa0', ' ').strip().rstrip(',')
                schedules.append(start_time)
                row_ids.append(row_id)

            if not schedules:
                self.log("No schedules to transform.")
                messagebox.showinfo("Info", "No schedules found in start_time column.")
                return

            # Log input schedules for debugging
            self.log(f"Input schedules count: {len(schedules)}")
            self.log(f"Input schedules: {schedules}")

            # Construct the prompt with indexed schedules
            indexed_schedules = [f"[{i}] {schedule}" for i, schedule in enumerate(schedules)]
            joined_schedules = '\n'.join(indexed_schedules)
            prompt = f"""You are a schedule transformer. Given a list of {len(schedules)} schedule descriptions, each prefixed with an index [0] to [{len(schedules)-1}], transform each one into a structured format or an empty string based on the following rules:

    - The calendar is always 'C_STANDARD_CALENDAR'.
    - If the input matches the pattern 'Every X Min' or 'every X hour' (where X is a number), output an empty string ("").
    - For other inputs, determine the frequency, time, and timezone as follows:
    - Frequency:
        - 'DAILY' for inputs starting with 'Every Day', 'Daily', or 'Daily: Mo to Su'.
        - 'MONDAY_TO_FRIDAY' for inputs starting with 'MO-FR', 'Daily:MON to FRI', or 'Daily: MON to FRI'.
        - 'MONDAY_TO_SATURDAY' for inputs like 'Daily, From Monday until Saturday'.
        - 'FRIDAY' for inputs starting with 'Every Friday'.
        - 'SUNDAY' for inputs like 'weekly Sundays'.
        - 'LAST_DAY_OF_MONTH' for inputs starting with 'Month end'.
        - 'DAY_OF_MONTH_05' for inputs like 'monthly – each 5th day of a month'.
    - Time (if present):
        - Extract from the input (e.g., '10:00 pm', '7 am', '00:01 am', '19:00') and convert to 24-hour 00HHMM format (e.g., '10:00 pm' -> '002200', '7 am' -> '000700', '19:00' -> '001900').
        - If no time is provided, omit the time field in the output.
    - Timezone (if present):
        - Extract the timezone abbreviation (e.g., 'EET', 'CET') and format as 'TZ#' followed by the abbreviation (e.g., 'TZ#EET').
        - If no timezone is provided, omit the timezone field; this is valid (e.g., 'Month end, 00:01 am' outputs 'C_STANDARD_CALENDAR, LAST_DAY_OF_MONTH, 000001').
    - Output format:
    - If time and timezone are present: C_STANDARD_CALENDAR, frequency, time, timezone
    - If time is present but no timezone: C_STANDARD_CALENDAR, frequency, time
    - If only frequency is present: C_STANDARD_CALENDAR, frequency
    - If input is 'Every X Min', 'every X hour', or empty (''): ""
    - CRITICAL: Output EXACTLY {len(schedules)} lines, one per input schedule, in the same order as the input list (from [0] to [{len(schedules)-1}]). For invalid, untransformable, or unrecognized inputs (e.g., malformed time formats), output an empty string (""). Do NOT include code block markers (```), extra newlines, explanations, or any additional content beyond the {len(schedules)} transformed schedule lines.
    - Each output line must be a valid transformation or an empty string. Ensure no trailing or leading newlines, no duplicate lines, and no formatting markers.

    Examples:
    - Input: [0] Every Friday (11PM CET)
    Output: C_STANDARD_CALENDAR, FRIDAY, 002300, TZ#CET
    - Input: [1] Every 15 Min
    Output: ""
    - Input: [2] Every Day (01:00AM CET)
    Output: C_STANDARD_CALENDAR, DAILY, 000100, TZ#CET
    - Input: [3] MO-FR, 10:00 pm, EET
    Output: C_STANDARD_CALENDAR, MONDAY_TO_FRIDAY, 002200, TZ#EET
    - Input: [4] Month end, 00:01 am, EET
    Output: C_STANDARD_CALENDAR, LAST_DAY_OF_MONTH, 000001, TZ#EET
    - Input: [5] Month end, 00:01 am
    Output: C_STANDARD_CALENDAR, LAST_DAY_OF_MONTH, 000001
    - Input: [6] Daily
    Output: C_STANDARD_CALENDAR, DAILY
    - Input: [7] Daily:MON to FRI
    Output: C_STANDARD_CALENDAR, MONDAY_TO_FRIDAY
    - Input: [8] Daily, From Monday until Saturday
    Output: C_STANDARD_CALENDAR, MONDAY_TO_SATURDAY
    - Input: [9] weekly Sundays
    Output: C_STANDARD_CALENDAR, SUNDAY
    - Input: [10] monthly – each 5th day of a month
    Output: C_STANDARD_CALENDAR, DAY_OF_MONTH_05
    - Input: [11] every 1 hour, 7 am, EET
    Output: ""
    - Input: [12] Daily:MON to FRI, 7 am, EET
    Output: C_STANDARD_CALENDAR, MONDAY_TO_FRIDAY, 000700, TZ#EET
    - Input: [13] Daily 7 am, EET
    Output: C_STANDARD_CALENDAR, DAILY, 000700, TZ#EET
    - Input: [14] Invalid Schedule
    Output: ""
    - Input: [15] Daily:MON to FRI,19:00 EET
    Output: C_STANDARD_CALENDAR, MONDAY_TO_FRIDAY, 001900, TZ#EET
    - Input: [16] 
    Output: ""

    Transform the following schedules, provided as a list with indices, and output EXACTLY {len(schedules)} lines, one per input, in the same order. Return "" for invalid inputs, 'Every X Min'/'every X hour', or empty inputs (''):

    Schedules:
    {joined_schedules}

    Output EXACTLY {len(schedules)} lines, one per input schedule, in the same order, with no code block markers (```) or extra formatting.
    """

            # Set up the content for the API
            contents = [
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_text(text=prompt),
                    ],
                ),
            ]

            # Configuration for content generation
            generate_content_config = types.GenerateContentConfig(
                temperature=0.7,
                top_p=0.95,
                top_k=64,
                max_output_tokens=65536,
                response_mime_type="text/plain",
            )

            # Call the Gemini API
            response_stream = client.models.generate_content_stream(
                model=model,
                contents=contents,
                config=generate_content_config,
            )

            # Collect the response
            response_text = ""
            for chunk in response_stream:
                response_text += chunk.text

            # Remove any trailing or leading newlines and split into lines
            response_text = response_text.strip()
            transformed_schedules = response_text.splitlines()

            # Remove code block markers if present
            transformed_schedules = [line for line in transformed_schedules if line not in ['```', '```plaintext']]

            # Log for debugging
            self.log(f"Input schedules count: {len(schedules)}, Output schedules count: {len(transformed_schedules)}")
            self.log(f"Input schedules: {schedules}")
            self.log(f"Output schedules: {transformed_schedules}")

            # Validate output formats
            for i, (transformed, original) in enumerate(zip(transformed_schedules, schedules)):
                if transformed and transformed != '""':
                    parts = transformed.split(', ')
                    if len(parts) < 2 or parts[0] != 'C_STANDARD_CALENDAR':
                        self.log(f"Invalid output format for schedule '{original}' at index {i}: '{transformed}'")
                        transformed_schedules[i] = '""'
                    elif len(parts) >= 3 and not parts[2].startswith('00') and not parts[2].isdigit():
                        self.log(f"Invalid time format for schedule '{original}' at index {i}: '{transformed}'")
                        transformed_schedules[i] = '""'

            # Handle mismatched output
            if len(transformed_schedules) != len(schedules):
                self.log(f"Warning: Expected {len(schedules)} transformed schedules, but received {len(transformed_schedules)}.")
                messagebox.showwarning(
                    "Warning",
                    f"Expected {len(schedules)} transformed schedules, but received {len(transformed_schedules)}. "
                    "Truncating or padding output to match input count."
                )
                # Log unexpected output for debugging
                if len(transformed_schedules) > len(schedules):
                    self.log(f"Extra outputs: {transformed_schedules[len(schedules):]}")
                elif len(transformed_schedules) < len(schedules):
                    missing_schedules = schedules[len(transformed_schedules):]
                    self.log(f"Missing outputs for schedules: {missing_schedules}")
                # Truncate or pad to match input count
                transformed_schedules = transformed_schedules[:len(schedules)] + [""] * (len(schedules) - len(transformed_schedules))

            # Update Treeview with transformed schedules
            for i, (row_id, transformed, original) in enumerate(zip(row_ids, transformed_schedules, schedules)):
                transformed = transformed.strip()  # Remove any whitespace
                # Convert API's '""' to empty string for Treeview
                if transformed == '""':
                    transformed = ""
                values = list(self.pairs_tree.item(row_id)['values'])
                values[7] = transformed  # Use empty string if API returned empty
                self.pairs_tree.item(row_id, values=values)
                if not transformed:
                    self.log(f"Row {row_id} (index {i}): Applied empty string for schedule: '{original}'")
                else:
                    self.log(f"Row {row_id} (index {i}): Successfully transformed to: '{transformed}'")

            self.log("Transformed schedules in start_time column using Gemini API.")

        except Exception as e:
            self.log(f"Error transforming schedules: {str(e)}")
            messagebox.showerror("Error", f"Failed to transform schedules: {str(e)}")
            self.undo_stack.pop()  # Remove the undo state if transformation fails
            
def extract_default_login(template_jobs):
    for proc in template_jobs.get('scripts', []):
        if 'process' in proc:
            for line in proc['process']:
                if line:
                    m = re.match(r":PUT_ATT\s+LOGIN\s*=\s*'([^']+)'", line)
                    if m:
                        return m.group(1)
    return 'LOGIN_R3_060_SY-BATCH'

# class AttributeUpdaterApp:
#     def __init__(self, parent, env_var, client_var, entries):
#         self.parent = parent
#         self.env_var = env_var
#         self.client_var = client_var
#         self.entries = entries
#         self.undo_stack = []
#         self.redo_stack = []
#         self.failed_jobs = []
#         self.build_ui()
#         self.current_entry = None

#     def build_ui(self):
#         frm = ttk.Frame(self.parent, padding=15)
#         frm.pack(fill='both', expand=True)

#         columns = ('object_name', 'MAIL_ADDRESS', 'AGGREG_LEVEL', 'IT_PRODUCT', 'RUNTIMELIMIT_RECIPIENT', 'ALERT_TYPE')
#         self.attributes_tree = ttk.Treeview(frm, columns=columns, show='headings')
#         for col in columns:
#             self.attributes_tree.heading(col, text=col)
#             self.attributes_tree.column(col, width=150, stretch=True)
#         self.attributes_tree.grid(row=0, column=0, columnspan=4, sticky='nsew')

#         self.paste_menu = tk.Menu(frm, tearoff=0)
#         self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)

#         self.attributes_tree.bind("<Button-3>", self.show_paste_menu)
#         self.attributes_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='object_name'))
#         self.attributes_tree.bind("<Double-1>", self.on_double_click)
#         self.attributes_tree.bind("<Control-z>", self.undo)
#         self.attributes_tree.bind("<Control-y>", self.redo)
#         self.attributes_tree.bind("<Control-a>", self.select_all)
#         self.attributes_tree.bind("<Delete>", self.delete_selected)
#         self.attributes_tree.bind("<BackSpace>", self.delete_selected)
#         self.attributes_tree.bind("<Control-c>", self.copy_selected)  # New binding for Ctrl+C
#         # self.parent.bind("<Button-1>", self.close_popup)
#         self.attributes_tree.bind("<Tab>", lambda e: self.on_tab_press(e, direction="forward"))
#         self.attributes_tree.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, direction="backward"))

#         ttk.Label(frm, text='Update message:').grid(row=5, column=0, sticky='w')
#         self.update_message_entry = ttk.Entry(frm)
#         self.update_message_entry.grid(row=5, column=1, sticky='ew', padx=5)

#         # Add ARMT_NO field for AttributeUpdaterApp
#         ttk.Label(frm, text="ARMT No:").grid(row=6, column=0, sticky="w", padx=5)
#         self.armt_no_entry = ttk.Entry(frm)  # Separate entry for AttributeUpdaterApp
#         self.armt_no_entry.grid(row=6, column=1, padx=5, sticky="ew")

#         self.update_btn = ttk.Button(frm, text='Update Attributes', command=self.start_update)
#         self.update_btn.grid(row=1, column=0, columnspan=3, pady=12)

#         ttk.Label(frm, text='Update Log:').grid(row=2, column=0, sticky='nw')
#         self.log_box = tk.Text(frm, height=10, state='disabled')
#         self.log_box.grid(row=3, column=0, columnspan=3, sticky='ew', padx=5)

#         self.copy_failed_btn = ttk.Button(frm, text='Copy Failed Jobs', command=self.copy_failed_jobs, state='disabled')
#         self.copy_failed_btn.grid(row=4, column=0, columnspan=3, pady=12)

#         ttk.Label(frm, text='Update message:').grid(row=5, column=0, sticky='w')
#         self.update_message_entry = ttk.Entry(frm)
#         self.update_message_entry.grid(row=5, column=1, sticky='ew', padx=5)

#         # Add ARMT_NO field for AttributeUpdaterApp
#         ttk.Label(frm, text="ARMT No:").grid(row=6, column=0, sticky="w", padx=5)
#         self.armt_no_entry = ttk.Entry(frm)  # Separate entry for AttributeUpdaterApp
#         self.armt_no_entry.grid(row=6, column=1, padx=5, sticky="ew")

#         frm.grid_rowconfigure(0, weight=1)
#         frm.grid_columnconfigure(1, weight=1)
#         self.ensure_empty_row()

#     def load_config(self, config):
#         if 'ARMT_NO' in config:
#             self.armt_no_entry.insert(0, config['ARMT_NO'])

#     def save_config(self):
#         return {'ARMT_NO': self.armt_no_entry.get()}

#     def show_paste_menu(self, event):
#         self.clicked_column = self.attributes_tree.identify_column(event.x)[1:]
#         self.paste_menu.tk_popup(event.x_root, event.y_root)

#     def paste_from_clipboard(self, start_column=None):
#         self.undo_stack.append(self.save_state())
#         self.redo_stack.clear()

#         columns = self.attributes_tree['columns']
#         if start_column is None:
#             if hasattr(self, 'clicked_column') and self.clicked_column:
#                 col_index = int(self.clicked_column) - 1
#                 start_column = columns[col_index]
#             else:
#                 start_column = 'object_name'
#         try:
#             start_col_index = columns.index(start_column)
#         except ValueError:
#             messagebox.showerror("Error", f"Invalid column: {start_column}")
#             self.undo_stack.pop()
#             return

#         try:
#             clipboard_data = self.parent.clipboard_get()
#             lines = clipboard_data.strip().splitlines()
#             if not lines:
#                 messagebox.showerror("Error", "Clipboard is empty.")
#                 self.undo_stack.pop()
#                 return

#             existing_rows = list(self.attributes_tree.get_children())
#             max_columns = len(columns) - start_col_index

#             for i, line in enumerate(lines):
#                 clipboard_values = line.split('\t') if '\t' in line else [line]
#                 clipboard_values = clipboard_values[:max_columns]
#                 while len(clipboard_values) < max_columns:
#                     clipboard_values.append('')

#                 if i < len(existing_rows):
#                     row_id = existing_rows[i]
#                     current_values = list(self.attributes_tree.item(row_id)['values'])
#                     for j, value in enumerate(clipboard_values):
#                         current_values[start_col_index + j] = value.strip()
#                     self.attributes_tree.item(row_id, values=current_values)
#                 else:
#                     new_values = [''] * len(columns)
#                     for j, value in enumerate(clipboard_values):
#                         new_values[start_col_index + j] = value.strip()
#                     self.attributes_tree.insert('', 'end', values=new_values)

#             self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
#         except tk.TclError:
#             messagebox.showerror("Error", "Clipboard contains invalid data.")
#             self.undo_stack.pop()
#         except Exception as e:
#             messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
#             self.undo_stack.pop()
#         self.ensure_empty_row()

#     def copy_selected(self, event=None):
#         """Copy selected Treeview rows to the clipboard as tab-separated values."""
#         selected = self.attributes_tree.selection()
#         if not selected:
#             self.log("No rows selected to copy.", tag='info')
#             return 'break'

#         rows_data = []
#         for row_id in selected:
#             values = self.attributes_tree.item(row_id)['values']
#             # Only copy non-empty rows (skip if all values are empty)
#             if any(val.strip() for val in values):
#                 rows_data.append('\t'.join(str(val) for val in values))

#         if rows_data:
#             clipboard_text = '\n'.join(rows_data)
#             self.parent.clipboard_clear()
#             self.parent.clipboard_append(clipboard_text)
#             self.parent.update()
#             self.log(f"Copied {len(rows_data)} row(s) to clipboard.", tag='info')
#         else:
#             self.log("No non-empty rows selected to copy.", tag='info')

#         return 'break'
#     def ensure_empty_row(self):
#         """Ensure there is always one empty row at the end of the Treeview."""
#         columns = self.attributes_tree['columns']
#         children = self.attributes_tree.get_children()
#         if children:
#             last_row_values = self.attributes_tree.item(children[-1])['values']
#             last_row_empty = all(val == '' for val in last_row_values)
#             if last_row_empty:
#                 return
#         empty_row = [''] * len(columns)
#         self.attributes_tree.insert('', 'end', values=empty_row)
#     def on_double_click(self, event):
#         row_id = self.attributes_tree.identify_row(event.y)
#         if not row_id:
#             return
#         column = self.attributes_tree.identify_column(event.x)
#         column_num = int(column[1:]) - 1
#         column_name = self.attributes_tree['columns'][column_num]
#         self.edit_cell(row_id, column_name, column_num)

#     def edit_cell(self, row_id, column_name, column_num):
#         """Open an Entry widget for editing a specific cell."""
#         if self.current_entry:
#             self.current_entry.destroy()
#         x, y, width, height = self.attributes_tree.bbox(row_id, f"#{column_num + 1}")
#         self.current_entry = ttk.Entry(self.parent)
#         self.current_entry.place(x=x, y=y, width=width, height=height)
#         current_value = self.attributes_tree.item(row_id)['values'][column_num]
#         self.current_entry.insert(0, current_value)
#         self.current_entry.select_range(0, tk.END)
#         self.current_entry.focus()
#         self.current_entry.bind("<Return>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
#         self.current_entry.bind("<FocusOut>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
#         self.current_entry.bind("<Escape>", lambda e: self.current_entry.destroy())
#         self.current_entry.bind("<Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="forward"))
#         self.current_entry.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="backward"))

#     def on_tab_press(self, event, current_row_id=None, current_col_num=None, direction="forward"):
#         """Handle Tab (forward) and Shift+Tab (backward) key presses to move to the next or previous cell."""
#         if current_row_id is None or current_col_num is None:
#             # Tab or Shift+Tab pressed directly in Treeview
#             selected = self.attributes_tree.selection()
#             if not selected:
#                 return 'break'
#             current_row_id = selected[0]
#             current_col_num = 0 if direction == "forward" else len(self.attributes_tree['columns']) - 1
#         else:
#             # Tab or Shift+Tab pressed in Entry widget, save current edit
#             self.save_edit(self.current_entry, current_row_id, self.attributes_tree['columns'][current_col_num])
        
#         columns = self.attributes_tree['columns']
#         children = self.attributes_tree.get_children()
#         current_row_index = children.index(current_row_id)
        
#         # Check if the last row is empty to limit navigation
#         last_row_empty = all(val == '' for val in self.attributes_tree.item(children[-1])['values'])
        
#         if direction == "forward":
#             # Move to next column or next row
#             if current_col_num < len(columns) - 1:
#                 next_col_num = current_col_num + 1
#                 next_row_id = current_row_id
#             else:
#                 # Move to the first column of the next row
#                 if current_row_index < len(children) - 1:
#                     next_row_id = children[current_row_index + 1]
#                     next_col_num = 0
#                 elif current_row_index == len(children) - 1 and not last_row_empty:
#                     next_row_id = current_row_id
#                     next_col_num = current_col_num
#                 else:
#                     return 'break'  # Stay at last cell if last row is empty
#         else:  # direction == "backward"
#             # Move to previous column or previous row
#             if current_col_num > 0:
#                 next_col_num = current_col_num - 1
#                 next_row_id = current_row_id
#             else:
#                 # Move to the last column of the previous row
#                 if current_row_index > 0:
#                     next_row_id = children[current_row_index - 1]
#                     next_col_num = len(columns) - 1
#                 else:
#                     return 'break'  # Stay at first cell
        
#         next_col_name = columns[next_col_num]
#         self.edit_cell(next_row_id, next_col_name, next_col_num)
#         return 'break'


#     def save_edit(self, entry, row_id, column_name):
#         new_value = entry.get().strip()
#         self.undo_stack.append(self.save_state())
#         self.redo_stack.clear()
#         self.attributes_tree.set(row_id, column_name, new_value)
#         entry.destroy()
#         self.ensure_empty_row()

#     def save_state(self):
#         return [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]

#     def restore_state(self, state):
#         self.attributes_tree.delete(*self.attributes_tree.get_children())
#         for values in state:
#             self.attributes_tree.insert('', 'end', values=values)

#     def undo(self, event):
#         if self.undo_stack:
#             state = self.undo_stack.pop()
#             self.redo_stack.append(self.save_state())
#             self.restore_state(state)
#         return 'break'

#     def redo(self, event):
#         if self.redo_stack:
#             state = self.redo_stack.pop()
#             self.undo_stack.append(self.save_state())
#             self.restore_state(state)
#         return 'break'

#     def select_all(self, event):
#         self.attributes_tree.selection_set(self.attributes_tree.get_children())
#         return 'break'

#     def delete_selected(self, event):
#         selected = self.attributes_tree.selection()
#         if selected:
#             self.undo_stack.append(self.save_state())
#             self.redo_stack.clear()
#             for item in selected:
#                 self.attributes_tree.delete(item)
#             self.ensure_empty_row()

#         return 'break'

#     def start_update(self):
#         self.update_btn.config(state='disabled')
#         threading.Thread(target=self.execute_update, daemon=True).start()

#     def execute_update(self):
#         self.failed_jobs = []
#         rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
#         rows = [row for row in rows if any(val.strip() for val in row)]
#         if not rows:
#             self.log("No data to update.")
#             self.parent.after(0, lambda: self.update_btn.config(state='normal'))
#             return

#         env = self.env_var.get().strip()
#         try:
#             cid = int(self.client_var.get().strip())
#         except ValueError:
#             self.log("Error: Invalid Client ID")
#             self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
#             self.parent.after(0, lambda: self.update_btn.config(state='normal'))
#             return

#         user = self.entries['USERID'].get().strip()
#         pwd = self.entries['PASSWORD'].get().strip()
#         api_url = f'https://rb-{env}-api.bosch.com'
#         name = self.entries['NAME'].get().strip()

#         auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
#         try:
#             automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
#         except requests.exceptions.HTTPError as e:
#             self.log(f"Authentication failed: {str(e)}")
#             self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
#             self.parent.after(0, lambda: self.update_btn.config(state='normal'))
#             return

#         update_message = self.update_message_entry.get().strip()
#         if not update_message:
#             self.log("No update message provided.")
#             self.parent.after(0, lambda: self.update_btn.config(state='normal'))
#             return
#         armt_no = self.armt_no_entry.get().strip()  # Use the tab-specific ARMT_NO
#         current_date = date.today().strftime("%d/%m/%Y")
#         doku_entry = f"{armt_no}, {name}, {current_date}, {update_message}"
#         updates = {}
#         for row in rows:
#             object_name = row[0].strip()
#             mail_address = row[1].strip()
#             aggreg_level = str(row[2]).strip()
#             it_product = row[3].strip()
#             runtimelimit_recipient = row[4].strip()
#             alert_type = row[5].strip()
#             if not object_name:
#                 self.log("Skipping row with empty object_name")
#                 continue
#             if object_name not in updates:
#                 updates[object_name] = {}
#             if mail_address:
#                 updates[object_name]['MAIL_ADDRESS'] = mail_address
#             if aggreg_level:
#                 updates[object_name]['AGGREG_LEVEL'] = aggreg_level
#             if it_product:
#                 updates[object_name]['IT_PRODUCT'] = it_product
#             if runtimelimit_recipient:
#                 updates[object_name]['RUNTIMELIMIT_RECIPIENT'] = runtimelimit_recipient
#             if alert_type:
#                 updates[object_name]['ALERT_TYPE'] = alert_type

#         for object_name, attrs in updates.items():
#             self.log(f"Processing {object_name}")
#             try:
#                 resp = automic.getObjects(client_id=cid, object_name=object_name)
#                 if resp.status != 200:
#                     self.log(f"Failed to fetch {object_name}: {resp.status}")
#                     self.failed_jobs.append(object_name)
#                     continue
#                 if 'data' not in resp.response:
#                     self.log(f"Error: Object {object_name} response missing 'data' key")
#                     self.failed_jobs.append(object_name)
#                     continue
#                 if 'jobs' in resp.response['data']:
#                     updating_obj = resp.response["data"]["jobs"]
#                 elif 'jobp' in resp.response['data']:
#                     updating_obj = resp.response["data"]["jobp"]
#                 else:
#                     self.log(f"Error: Object {object_name} not found or not a job/jobp")
#                     self.failed_jobs.append(object_name)
#                     continue
#                 updated = False
#                 doku_found = False
#                 for doc in updating_obj.get('documentation', []):
#                     if '_BSH' in doc:
#                         bsh_list = doc['_BSH']
#                         if not isinstance(bsh_list, list) or not bsh_list:
#                             self.log(f"Error: Invalid _BSH data for {object_name}")
#                             self.failed_jobs.append(object_name)
#                             continue
#                         content_str = None
#                         content_index = None
#                         for i, item in enumerate(bsh_list):
#                             if item.strip().startswith('<Content'):
#                                 content_str = item
#                                 content_index = i
#                                 break
#                         if content_str is None:
#                             self.log(f"Error: No <Content> element found in _BSH for {object_name}")
#                             self.failed_jobs.append(object_name)
#                             continue
#                         if not content_str.strip():
#                             self.log(f"Error: Empty content_str for {object_name}")
#                             self.failed_jobs.append(object_name)
#                             continue
#                         try:
#                             content_element = ET.fromstring(content_str)
#                             for attr_name, new_value in attrs.items():
#                                 content_element.set(attr_name, str(new_value.strip()))
#                                 self.log(f"Updated {attr_name} to {new_value} for {object_name}")
#                             updated_content_str = ET.tostring(content_element, encoding='unicode')
#                             bsh_list[content_index] = updated_content_str
#                             updated = True
#                         except ET.ParseError as e:
#                             self.log(f"Error parsing XML for {object_name}: {str(e)}")
#                             self.failed_jobs.append(object_name)
#                             continue

#                     if doc.get('Doku'):
#                         doku_list = doc['Doku']
#                         if isinstance(doku_list, list):
#                             doku_list.append(doku_entry)
#                             doku_found = True
#                             break



#                 if not updated:
#                     self.log(f"No _BSH documentation found for {object_name}")
#                     self.failed_jobs.append(object_name)
#                     continue

#                 # if not doku_found:
#                 #     jobs.setdefault('documentation', []).append({'Doku': [doku_entry]})

#                 resp_update = automic.postObjects(client_id=cid, body=resp.response, query="overwrite_existing_objects=true")
#                 if resp_update.status is None:
#                     self.log(f"Successfully updated {object_name}")
#                 else:
#                     self.log(f"Failed to update {object_name}: {resp_update.status}")
#                     self.failed_jobs.append(object_name)
#             except Exception as e:
#                 self.log(f"Error updating {object_name}: {str(e)}")
#                 self.failed_jobs.append(object_name)

#         self.log("All updates completed.")
#         self.parent.after(0, lambda: self.update_btn.config(state='normal'))
#         self.parent.after(0, lambda: self.copy_failed_btn.config(state='normal' if self.failed_jobs else 'disabled'))

#     def copy_failed_jobs(self):
#         if self.failed_jobs:
#             failed_list = "\n".join(self.failed_jobs)
#             self.parent.clipboard_clear()
#             self.parent.clipboard_append(failed_list)
#             self.log("Copied failed jobs to clipboard.")
#         else:
#             messagebox.showinfo("Info", "No failed jobs to copy.")

#     def log(self, msg):
#         self.parent.after(0, lambda: self._log(msg))

#     def _log(self, msg):
#         self.log_box.config(state='normal')
#         self.log_box.insert('end', msg + '\n')
#         self.log_box.see('end')
#         self.log_box.config(state='disabled')

class PvlChecker:
    CLIENT_MAP = {
        'eup6': ['1001', '1111'],
        'eup7': ['1101', '1301', '1401', '7101','7001']
    }
    TIMEOUT = 30
    MAX_WORKERS = 35  # Adjust based on API rate limits

    def __init__(self, parent, env_var, client_var, entries):
            # Setup logging
        self.logger = logging.getLogger(__name__)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )        
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.undo_stack = []
        self.redo_stack = []
        self.failed_jobs = []
        self.failed_logins = []

        self.connection_lock = Lock()  # Lock for thread-safe connection
        self.stop_flag = False

        # Encode credentials once
        self.auth = None
        self.job_names_map = {}  # Store multiple job names per row_id
        self.login_names_map = {}
        self.active_popup = None  # Track active Listbox popup
        self.current_entry = None
        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        columns = ('object_name', 'Program', 'Variant', 'Login')
        self.attributes_tree = ttk.Treeview(frm, columns=columns, show='headings')
        for col in columns:
            self.attributes_tree.heading(col, text=col)
            self.attributes_tree.column(col, width=150, stretch=True)
            self.attributes_tree.heading(col, command=lambda c=col: self.on_column_click(c))
        self.attributes_tree.grid(row=0, column=0, columnspan=4, sticky='nsew')
        # Configure row-level tags
        self.attributes_tree.tag_configure('multiple_jobs', background='#FFFFCC')  # Yellow
        self.attributes_tree.tag_configure('multiple_logins', background='#E6E6FA')  # Lavender

        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)

        self.attributes_tree.bind("<Button-3>", self.show_paste_menu)
        self.attributes_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='object_name'))
        self.attributes_tree.bind("<Control-c>", self.copy_selected)  # New binding for Ctrl+C
        self.attributes_tree.bind("<Double-1>", self.on_double_click)
        self.attributes_tree.bind("<Button-1>", self.on_single_click)  # New binding for popup
        self.attributes_tree.bind("<Control-z>", self.undo)
        self.attributes_tree.bind("<Control-y>", self.redo)
        self.attributes_tree.bind("<Control-a>", self.select_all)
        self.attributes_tree.bind("<Delete>", self.delete_selected)
        self.attributes_tree.bind("<BackSpace>", self.delete_selected)
        # self.parent.bind("<Button-1>", self.close_popup)
        self.attributes_tree.bind("<Tab>", lambda e: self.on_tab_press(e, direction="forward"))
        self.attributes_tree.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, direction="backward"))

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=12, sticky='ew')

        self.update_btn = ttk.Button(btn_frame, text='Check', command=self.start_check)
        self.update_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text='Stop Check', command=self.stop_check, state='disabled')
        self.stop_btn.grid(row=0, column=1, padx=5)

        self.export_btn = ttk.Button(btn_frame, text='Export to Excel', command=self.export_to_excel)
        self.export_btn.grid(row=0, column=2, padx=5)

        ttk.Label(frm, text='Searching Log:').grid(row=2, column=0, sticky='nw')
        self.log_box = tk.Text(frm, height=10, state='disabled')
        self.log_box.grid(row=3, column=0, columnspan=3, sticky='ew', padx=5)

        # Configure log box tags for color highlighting
        self.log_box.tag_configure('success', foreground='green')
        self.log_box.tag_configure('job_name', foreground='blue')
        self.log_box.tag_configure('login_name', foreground='purple')
        self.log_box.tag_configure('error', foreground='red')
        self.log_box.tag_configure('info', foreground='black')

        frm.grid_rowconfigure(0, weight=1)
        frm.grid_columnconfigure(1, weight=1)
        self.ensure_empty_row()  # Initialize with an empty row
    def copy_selected(self, event=None):
        """Copy selected Treeview rows to the clipboard as tab-separated values."""
        selected = self.attributes_tree.selection()
        if not selected:
            self.log("No rows selected to copy.", tag='info')
            return 'break'

        rows_data = []
        for row_id in selected:
            values = self.attributes_tree.item(row_id)['values']
            # Only copy non-empty rows (skip if all values are empty)
            if any(val.strip() for val in values):
                rows_data.append('\t'.join(str(val) for val in values))

        if rows_data:
            clipboard_text = '\n'.join(rows_data)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(clipboard_text)
            self.parent.update()
            self.log(f"Copied {len(rows_data)} row(s) to clipboard.", tag='info')
        else:
            self.log("No non-empty rows selected to copy.", tag='info')

        return 'break'
    def ensure_empty_row(self):
        """Ensure there is always one empty row at the end of the Treeview."""
        columns = self.attributes_tree['columns']
        children = self.attributes_tree.get_children()
        if children:
            last_row_values = self.attributes_tree.item(children[-1])['values']
            last_row_empty = all(val == '' for val in last_row_values)
            if last_row_empty:
                return
        empty_row = [''] * len(columns)
        self.attributes_tree.insert('', 'end', values=empty_row)
    def on_double_click(self, event):
        row_id = self.attributes_tree.identify_row(event.y)
        if not row_id:
            return
        column = self.attributes_tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.attributes_tree['columns'][column_num]
        self.edit_cell(row_id, column_name, column_num)

    def edit_cell(self, row_id, column_name, column_num):
        """Open an Entry widget for editing a specific cell."""
        if self.current_entry:
            self.current_entry.destroy()
        x, y, width, height = self.attributes_tree.bbox(row_id, f"#{column_num + 1}")
        self.current_entry = ttk.Entry(self.parent)
        self.current_entry.place(x=x, y=y, width=width, height=height)
        current_value = self.attributes_tree.item(row_id)['values'][column_num]
        self.current_entry.insert(0, current_value)
        self.current_entry.select_range(0, tk.END)
        self.current_entry.focus()
        self.current_entry.bind("<Return>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<FocusOut>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<Escape>", lambda e: self.current_entry.destroy())
        self.current_entry.bind("<Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="forward"))
        self.current_entry.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="backward"))

    def on_tab_press(self, event, current_row_id=None, current_col_num=None, direction="forward"):
        """Handle Tab (forward) and Shift+Tab (backward) key presses to move to the next or previous cell."""
        if current_row_id is None or current_col_num is None:
            # Tab or Shift+Tab pressed directly in Treeview
            selected = self.attributes_tree.selection()
            if not selected:
                return 'break'
            current_row_id = selected[0]
            current_col_num = 0 if direction == "forward" else len(self.attributes_tree['columns']) - 1
        else:
            # Tab or Shift+Tab pressed in Entry widget, save current edit
            self.save_edit(self.current_entry, current_row_id, self.attributes_tree['columns'][current_col_num])
        
        columns = self.attributes_tree['columns']
        children = self.attributes_tree.get_children()
        current_row_index = children.index(current_row_id)
        
        # Check if the last row is empty to limit navigation
        last_row_empty = all(val == '' for val in self.attributes_tree.item(children[-1])['values'])
        
        if direction == "forward":
            # Move to next column or next row
            if current_col_num < len(columns) - 1:
                next_col_num = current_col_num + 1
                next_row_id = current_row_id
            else:
                # Move to the first column of the next row
                if current_row_index < len(children) - 1:
                    next_row_id = children[current_row_index + 1]
                    next_col_num = 0
                elif current_row_index == len(children) - 1 and not last_row_empty:
                    next_row_id = current_row_id
                    next_col_num = current_col_num
                else:
                    return 'break'  # Stay at last cell if last row is empty
        else:  # direction == "backward"
            # Move to previous column or previous row
            if current_col_num > 0:
                next_col_num = current_col_num - 1
                next_row_id = current_row_id
            else:
                # Move to the last column of the previous row
                if current_row_index > 0:
                    next_row_id = children[current_row_index - 1]
                    next_col_num = len(columns) - 1
                else:
                    return 'break'  # Stay at first cell
        
        next_col_name = columns[next_col_num]
        self.edit_cell(next_row_id, next_col_name, next_col_num)
        return 'break'

    def save_edit(self, entry, row_id, column_name):
        new_value = entry.get().strip()
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.attributes_tree.set(row_id, column_name, new_value)
        entry.destroy()
        self.current_entry = None
        self.ensure_empty_row()  # Ensure empty row after edit
    def on_column_click(self, col_name):
        col_index = self.attributes_tree['columns'].index(col_name)
        values = [self.attributes_tree.item(item)['values'][col_index] for item in self.attributes_tree.get_children()]
        text = "\n".join(str(v) for v in values if v)
        if text:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(text)
            self.parent.update()
            self.log(f"Copied column '{col_name}' to clipboard.", tag='info')

    def export_to_excel(self):
        """Export treeview data to an Excel file."""
        if not self.attributes_tree.get_children():
            messagebox.showinfo("Export Info", "No data to export. Please add or check data first.")
            self.log("No data to export.", tag='info')
            return

        try:
            # Prepare data for DataFrame
            columns = self.attributes_tree['columns']
            data = []
            for item in self.attributes_tree.get_children():
                row = {col: self.attributes_tree.item(item)['values'][i] for i, col in enumerate(columns)}
                data.append(row)

            # Create DataFrame
            df = pd.DataFrame(data, columns=columns)
            # Generate timestamp for filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"pvl_checker_data_{timestamp}.xlsx"
            # Save to Excel
            df.to_excel(filename, index=False)
            self.log(f"Exported data to {filename}", tag='info')
            messagebox.showinfo("Export Success", f"Successfully exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to Excel: {str(e)}", tag='error')
            messagebox.showerror("Export Error", f"Failed to export to Excel: {str(e)}")

    def on_single_click(self, event):
        # Close any existing popup
        if self.active_popup:
            self.active_popup.destroy()
            self.active_popup = None

        row_id = self.attributes_tree.identify_row(event.y)
        column = self.attributes_tree.identify_column(event.x)
        if not row_id:
            return
        if column == '#1' and row_id in self.job_names_map and len(self.job_names_map[row_id]) > 1:
            self.show_job_names_popup(row_id, event.x, event.y)
            return 'break'  # Stop event propagation
        elif column == '#4' and row_id in self.login_names_map and len(self.login_names_map[row_id]) > 1:
            self.show_login_names_popup(row_id, event.x, event.y)
            return 'break'  # Stop event propagation

    def show_job_names_popup(self, row_id, x, y):
        # Create a Listbox at the click location to show multiple job names
        listbox = tk.Listbox(self.parent, font=('Montserrat', 10), height=min(len(self.job_names_map[row_id]), 5))
        listbox.place(x=x, y=y, width=500)  # Fixed width, adjust as needed
        listbox.focus()

        # Populate Listbox with job names
        for env, client_id, _, _, job_name in self.job_names_map[row_id]:
            listbox.insert(tk.END, f"({env}/{client_id}) {job_name} ")

        # Create context menu for copying
        context_menu = tk.Menu(listbox, tearoff=0)
        context_menu.add_command(label="Copy to Clipboard", command=lambda: self.copy_job_names(row_id))
        listbox.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))

        # Dismiss Listbox on click outside or Escape
        listbox.bind("<FocusOut>", lambda e: listbox.destroy())
        listbox.bind("<Escape>", lambda e: listbox.destroy())
    def show_login_names_popup(self, row_id, x, y):
        listbox = tk.Listbox(self.parent, font=('Montserrat', 10), height=min(len(self.login_names_map[row_id]), 5))
        listbox_width = 200
        listbox_height = min(len(self.login_names_map[row_id]), 5) * 20
        self.parent.update()  # Ensure window dimensions are up-to-date
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        x_new = max(0, min(x, parent_width - listbox_width))
        y_new = max(0, min(y, parent_height - listbox_height))
        listbox.place(x=x_new, y=y_new, width=listbox_width)
        listbox.focus_set()  # Explicitly set focus
        self.active_popup = listbox  # Track active popup

        for _, login_name in self.login_names_map[row_id]:
            listbox.insert(tk.END, login_name)

        context_menu = tk.Menu(listbox, tearoff=0)
        context_menu.add_command(label="Copy to Clipboard", command=lambda: self.copy_login_names(row_id))
        listbox.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))
        listbox.bind("<FocusOut>", lambda e: listbox.destroy())
        listbox.bind("<Escape>", lambda e: listbox.destroy())

    def copy_job_names(self, row_id):
        job_names = [f"{job_name} ({env}/{client_id})" for env, client_id, _, _, job_name in self.job_names_map[row_id]]
        self.parent.clipboard_clear()
        self.parent.clipboard_append("\n".join(job_names))
        self.log("Copied multiple job names to clipboard.", tag='info')
    def copy_login_names(self, row_id):
        login_names = [login_name for _, login_name in self.login_names_map[row_id]]
        self.parent.clipboard_clear()
        self.parent.clipboard_append("\n".join(login_names))
        self.log("Copied multiple login names to clipboard.", tag='info')

    def load_config(self, config):
        if 'ARMT_NO' in config:
            self.armt_no_entry.insert(0, config['ARMT_NO'])

    def save_config(self):
        return {'ARMT_NO': self.armt_no_entry.get()}
    def show_paste_menu(self, event):
        self.clicked_column = self.attributes_tree.identify_column(event.x)[1:]
        self.paste_menu.tk_popup(event.x_root, event.y_root)


    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()

        columns = self.attributes_tree['columns']
        if start_column is None:
            if hasattr(self, 'clicked_column') and self.clicked_column:
                col_index = int(self.clicked_column) - 1
                start_column = columns[col_index]
            else:
                start_column = 'object_name'
        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return

        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return

            existing_rows = list(self.attributes_tree.get_children())
            max_columns = len(columns) - start_col_index

            for i, line in enumerate(lines):
                clipboard_values = line.split('\t') if '\t' in line else [line]
                clipboard_values = clipboard_values[:max_columns]
                while len(clipboard_values) < max_columns:
                    clipboard_values.append('')

                if i < len(existing_rows):
                    row_id = existing_rows[i]
                    current_values = list(self.attributes_tree.item(row_id)['values'])
                    for j, value in enumerate(clipboard_values):
                        current_values[start_col_index + j] = value.strip()
                    self.attributes_tree.item(row_id, values=current_values)
                else:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(clipboard_values):
                        new_values[start_col_index + j] = value.strip()
                    self.attributes_tree.insert('', 'end', values=new_values)

            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
            self.ensure_empty_row()  # Ensure empty row after paste
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
            self.undo_stack.pop()



    def save_state(self):
        return [(child, self.attributes_tree.item(child)['values'], self.attributes_tree.item(child)['tags']) 
                for child in self.attributes_tree.get_children()]

    def restore_state(self, state):
        self.attributes_tree.delete(*self.attributes_tree.get_children())
        for child, values, tags in state:
            self.attributes_tree.insert('', 'end', values=values)

    def undo(self, event):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def select_all(self, event):
        self.attributes_tree.selection_set(self.attributes_tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.attributes_tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.job_names_map.pop(item, None)  # Remove job names for deleted rows
                self.login_names_map.pop(item, None)
                self.attributes_tree.delete(item)
            self.ensure_empty_row()  # Ensure empty row after deletion
        return 'break'

    def start_check(self):
        try:
            userid = self.entries['USERID'].get().strip()
            password = self.entries['PASSWORD'].get().strip()
            if not userid or not password:
                self.log("Please enter both User ID and Password.", tag='error')
                return
            self.auth = base64.b64encode(f"{userid}:{password}".encode()).decode()
            self.logger.debug(f"Encoded auth in start_check: {self.auth}")
        except KeyError as e:
            self.log(f"Configuration error: missing {e} field.", tag='error')
            return

        self.stop_flag = False
        self.job_names_map.clear()  # Clear previous job names
        self.login_names_map.clear()
        self.failed_logins.clear()

        self.update_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        threading.Thread(target=self.execute_check, daemon=True).start()
    def stop_check(self):
        self.stop_flag = True
        self.log("Stopping check process...", tag='info')
        self.update_btn.config(state='normal')
        self.stop_btn.config(state='disabled')

    def execute_check(self):
        self.log("Starting object check...", tag='info')

        rows = [(child, self.attributes_tree.item(child)['values']) 
                for child in self.attributes_tree.get_children()]
        
        env = self.env_var.get().strip()
        client_id = self.client_var.get().strip()
        if not env or not client_id:
            self.log("Environment or Client ID not specified for login check.", tag='error')
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            self.parent.after(0, lambda: self.stop_btn.config(state='disabled'))
            return

        tasks = []
        login_tasks = []
        for row_id, values in rows:
            # Ensure values has at least 4 elements
            values = list(values) + [''] * (4 - len(values))
            program, variant, login = values[1], values[2], values[3]

            # Skip completely empty rows
            if not program and not variant and not login:
                self.log(f"Skipping empty row {row_id}", tag='error')
                continue

            # Job checks for Program, Variant, or Program+Variant
            if program or variant:
                for env_key, client_ids in self.CLIENT_MAP.items():
                    for cid in client_ids:
                        if program and variant:
                            tasks.append(('program*variant', env_key, cid, program, variant, row_id))
                        elif program:
                            tasks.append(('program', env_key, cid, program, None, row_id))
                        elif variant:
                            tasks.append(('variant', env_key, cid, None, variant, row_id))

            # Login check
            if login:
                login_tasks.append((env, client_id, login, row_id))

        if not tasks and not login_tasks:
            self.log("No valid program, variant, or login entries to check.", tag='error')
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            self.parent.after(0, lambda: self.stop_btn.config(state='disabled'))
            return

        try:
            with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
                # Job checks
                future_to_task = {
                    executor.submit(self.search_objects, task_type, env, client_id, program, variant): (task_type, env, client_id, program, variant, row_id)
                    for task_type, env, client_id, program, variant, row_id in tasks
                }

                for future in as_completed(future_to_task):
                    if self.stop_flag:
                        self.log("Check process interrupted by user.", tag='info')
                        executor._threads.clear()
                        break

                    task_type, env, client_id, program, variant, row_id = future_to_task[future]
                    try:
                        results = future.result()
                        if results:
                            if row_id not in self.job_names_map:
                                self.job_names_map[row_id] = []
                            self.job_names_map[row_id].extend(results)
                            job_name = results[0][4]
                            # Get current tags and values
                            current_tags = list(self.attributes_tree.item(row_id)['tags'])
                            current_values = list(self.attributes_tree.item(row_id)['values'])
                            current_values[0] = f"({env}/{client_id}) {job_name}"
                            if len(self.job_names_map[row_id]) > 1 and 'multiple_jobs' not in current_tags:
                                current_tags.append('multiple_jobs')
                            self.parent.after(0, lambda: self.attributes_tree.item(row_id, values=current_values, tags=current_tags))
                            query = f"{program}*{variant}" if task_type == 'program*variant' else program or variant
                            self.log(f"{task_type.capitalize()} {query} in {env}/{client_id}: {len(results)} job(s) found", tag='success')
                            for _, _, _, _, job_name in results:
                                self.log(job_name, tag='job_name')
                        else:
                            query = f"{program}*{variant}" if task_type == 'program*variant' else program or variant
                            # self.failed_jobs.append(query)
                            # self.log(f"No jobs found for {task_type} {query} in {env}/{client_id}", tag='error')
                    except Exception as e:
                        query = f"{program}*{variant}" if task_type == 'program*variant' else program or variant
                        self.failed_jobs.append(query)
                        self.log(f"Error checking {task_type} {query} in {env}/{client_id}: {str(e)}", tag='error')

                # Login checks
                future_to_login_task = {
                    executor.submit(self.search_login_objects, env, client_id, login): (env, client_id, login, row_id)
                    for env, client_id, login, row_id in login_tasks
                }

                for future in as_completed(future_to_login_task):
                    if self.stop_flag:
                        self.log("Check process interrupted by user.", tag='info')
                        executor._threads.clear()
                        break

                    env, client_id, login, row_id = future_to_login_task[future]
                    try:
                        results = future.result()
                        if results:
                            if row_id not in self.login_names_map:
                                self.login_names_map[row_id] = []
                            self.login_names_map[row_id].extend(results)
                            login_name = results[0][1]
                            # Get current tags and values
                            current_tags = list(self.attributes_tree.item(row_id)['tags'])
                            current_values = list(self.attributes_tree.item(row_id)['values'])
                            current_values[3] = login_name
                            if len(self.login_names_map[row_id]) > 1 and 'multiple_logins' not in current_tags:
                                current_tags.append('multiple_logins')
                            self.parent.after(0, lambda: self.attributes_tree.item(row_id, values=current_values, tags=current_tags))
                            self.log(f"Login found for {login} in {env}/{client_id}: {login_name}", tag='success')
                            for _, login_name in results:
                                self.log(login_name, tag='login_name')
                        else:
                            self.failed_logins.append(login)
                            self.log(f"No login found for {login} in {env}/{client_id}", tag='error')
                    except Exception as e:
                        self.failed_logins.append(login)
                        self.log(f"Error checking login {login} in {env}/{client_id}: {str(e)}", tag='error')
        except Exception as e:
            self.log(f"Error during check process: {str(e)}", tag='error')

        if not self.stop_flag:
            self.log("Check completed.", tag='info')
            # if self.failed_jobs:
                # self.log(f"Failed jobs: {len(self.failed_jobs)}", tag='error')
            if self.failed_logins:
                self.log(f"Failed logins: {len(self.failed_logins)}", tag='error')
        self.parent.after(0, lambda: self.update_btn.config(state='normal'))
        self.parent.after(0, lambda: self.stop_btn.config(state='disabled'))

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        retry=retry_if_exception_type((ConnectionError, ConnectionResetError, HTTPError)),
        before_sleep=lambda retry_state: logging.info(f"Retrying {retry_state.fn.__name__} ({retry_state.attempt_number}/3) after {retry_state.next_action.sleep}s")
    )
    def search_objects(self, task_type: str, env: str, client_id: str, program: str, variant: str) -> List[Tuple[str, str, str, str, str]]:
        if self.stop_flag:
            raise RuntimeError("Check process stopped by user")
        
        api_url = f'https://rb-{env}-api.bosch.com'
        try:
            with self.connection_lock:
                automic.connection(
                    url=api_url,
                    auth=self.auth,
                    sslverify=False
                )
            body = {
                "filters": [
                    {"object_types": ["JOBS"], "filter_identifier": "object_type"}
                ],
                "max_results": 999
            }
            if task_type == 'program*variant':
                body["filters"].append({"query": f"REPORT*{program}*VARIANT*{variant}", "filter_identifier": "process"})
                body["filters"].append({"query": f"VARIANT*{variant}*REPORT*{program}", "filter_identifier": "process"})

            elif task_type == 'program':
                body["filters"].append({"query": program, "filter_identifier": "process"})
            elif task_type == 'variant':
                body["filters"].append({"query": variant, "filter_identifier": "process"})
            
            resp = automic.findObjects(client_id=client_id, body=body)
            self.logger.debug(f"API response for {env}/{client_id} {program}*{variant}: {resp.response}")
            return [(env, client_id, program, variant, job['name']) for job in resp.response['data']]
        except Exception as e:
            self.logger.error(f"Error processing {env}/{client_id} for {program}*{variant}: {str(e)}", exc_info=True)
            raise
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        retry=retry_if_exception_type((ConnectionError, ConnectionResetError, HTTPError)),
        before_sleep=lambda retry_state: logging.info(f"Retrying {retry_state.fn.__name__} ({retry_state.attempt_number}/3) after {retry_state.next_action.sleep}s")
    )
    def search_login_objects(self, env: str, client_id: str, login: str) -> List[Tuple[str, str]]:
        if self.stop_flag:
            raise RuntimeError("Check process stopped by user")
        api_url = f'https://rb-{env}-api.bosch.com'
        try:
            with self.connection_lock:
                automic.connection(
                    url=api_url,
                    auth=self.auth,
                    noproxy=True,
                    sslverify=False,
                    timeout=self.TIMEOUT
                )
            body = {
                "filters": [
                    {"object_name": f"{login}", "filter_identifier": "object_name"},
                    {"object_types": ["LOGIN"], "filter_identifier": "object_type"}
                ],
                "max_results": 999
            }
            resp = automic.findObjects(client_id=client_id, body=body)
            self.logger.debug(f"API response for {env}/{client_id} login {login}: {resp.response}")
            return [(login, job['name']) for job in resp.response['data']]            
        except Exception as e:
            self.logger.error(f"Error processing {env}/{client_id} for login {login}: {str(e)}", exc_info=True)
            raise    
    def copy_failed_jobs(self):
        failed_items = []
        if self.failed_jobs:
            failed_items.extend(self.failed_jobs)
        if self.failed_logins:
            failed_items.extend(self.failed_logins)
        if failed_items:
            failed_list = "\n".join(failed_items)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(failed_list)
            self.log("Copied failed jobs and logins to clipboard.", tag='info')
        else:
            messagebox.showinfo("Info", "No failed jobs or logins to copy.")

    def log(self, msg, tag='info'):
        self.parent.after(0, lambda: self._log(msg, tag))

    def _log(self, msg, tag):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n', tag)
        self.log_box.see('end')
        self.log_box.config(state='disabled')
    # def start_append_doku(self):
    #     self.append_doku_btn.config(state='disabled')
    #     threading.Thread(target=self.execute_append_doku, daemon=True).start()

    # def execute_append_doku(self):
    #     update_message = self.update_message_entry.get().strip()
    #     if not update_message:
    #         self.log("No update message provided.")
    #         self.parent.after(0, lambda: self.append_doku_btn.config(state='normal'))
    #         return

    #     rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
    #     rows = [row for row in rows if row[0].strip()]
    #     if not rows:
    #         self.log("No objects to update.")
    #         self.parent.after(0, lambda: self.append_doku_btn.config(state='normal'))
    #         return

    #     env = self.env_var.get().strip()
    #     try:
    #         cid = int(self.client_var.get().strip())
    #     except ValueError:
    #         self.log("Error: Invalid Client ID")
    #         self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
    #         self.parent.after(0, lambda: self.append_doku_btn.config(state='normal'))
    #         return

    #     user = self.entries['USERID'].get().strip()
    #     pwd = self.entries['PASSWORD'].get().strip()
    #     name = self.entries['NAME'].get().strip()
    #     armt_no = self.armt_no_entry.get().strip()  # Use the tab-specific ARMT_NO
    #     api_url = f'https://rb-{env}-api.bosch.com'
    #     auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
    #     try:
    #         automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
    #     except requests.exceptions.HTTPError as e:
    #         self.log(f"Authentication failed: {str(e)}")
    #         self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
    #         self.parent.after(0, lambda: self.append_doku_btn.config(state='normal'))
    #         return
    #     armt_no = self.armt_no_entry.get().strip()  # Use the tab-specific ARMT_NO

    #     current_date = date.today().strftime("%d/%m/%Y")
    #     doku_entry = f"{armt_no}, {name}, {current_date}, {update_message}"

    #     for row in rows:
    #         object_name = row[0].strip()
    #         self.log(f"Processing {object_name}")
    #         try:
    #             resp = automic.getObjects(client_id=cid, object_name=object_name)
    #             if resp.status != 200:
    #                 self.log(f"Failed to fetch {object_name}: {resp.status}")
    #                 continue
    #             if 'data' not in resp.response:
    #                 self.log(f"Error: Object {object_name} response missing 'data' key")
    #                 continue
    #             if 'jobs' in resp.response['data']:
    #                 jobs = resp.response["data"]["jobs"]
    #             elif 'jobp' in resp.response['data']:
    #                 jobs = resp.response["data"]["jobp"]
    #             else:
    #                 self.log(f"Error: Object {object_name} not found or not a job/jobp")
    #                 continue

    #             doku_found = False
    #             for doc in jobs.get('documentation', []):
    #                 if 'Doku' in doc:
    #                     doku_list = doc['Doku']
    #                     if isinstance(doku_list, list):
    #                         doku_list.append(doku_entry)
    #                         break
    #             if not doku_found:
    #                 jobs.setdefault('documentation', []).append({'Doku': [doku_entry]})

    #             resp_update = automic.postObjects(client_id=cid, body=resp.response, query="overwrite_existing_objects=true")
    #             if resp_update.status is None:
    #                 self.log(f"Successfully appended to Doku for {object_name}")
    #             else:
    #                 self.log(f"Failed to update {object_name}: {resp_update.status}")
    #         except Exception as e:
    #             self.log(f"Error updating {object_name}: {str(e)}")

    #     self.log("All Doku updates completed.")
    #     self.parent.after(0, lambda: self.append_doku_btn.config(state='normal'))




class JobsUpdater:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.undo_stack = []
        self.redo_stack = []
        self.failed_jobs = []
        self.job_processes = {}  # Store script['process'] for each job
        self.job_attributes = {}  # Store {jobname: {attr_name: value}}
        # Available options for dropdowns
        self.all_options = [''] + ['EXPIR', 'LIST_T', 'LINE_COUNT']
        self.current_entry = None

        self.column_configs = ['jobname', None, None, None, None, None]  # Track selected attributes
        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        # Frame for dropdown headers
        header_frame = ttk.Frame(frm)
        header_frame.grid(row=0, column=0, columnspan=5, sticky='nsew')

        # Define initial columns
        self.columns = ('jobname', 'attribute1', 'attribute2', 'attribute3', 'attribute4')
        self.attributes_tree = ttk.Treeview(frm, columns=self.columns, show='headings')
        self.attributes_tree.heading('jobname', text='Job Name')
        self.attributes_tree.column('jobname', width=200, stretch=True)
        for col in self.columns[1:]:
            self.attributes_tree.heading(col, text=col)
            self.attributes_tree.column(col, width=150, stretch=True)
        self.attributes_tree.grid(row=1, column=0, columnspan=5, sticky='nsew')

        # Combined options for all dropdowns
        # all_options = [''] + ['SAP_RECIPIENT', 'SAP_ADDRESSTYPE', 'JOB_NAME', 'LOGIN','SAP_LANG', 'EXPIR', 'LIST_T', 'LINE_COUNT']
        self.dropdowns = []
        ttk.Label(header_frame, text="Select Attributes:").grid(row=0, column=0, sticky='w', pady=5)
        for i, col in enumerate(self.columns, 0):
            var = tk.StringVar()
            if i == 0:  # jobname column
                menu = ttk.OptionMenu(header_frame, var, 'Job Name', 'Job Name')
                menu.config(state='disabled')  # Disable dropdown for jobname
            else:
                menu = ttk.OptionMenu(header_frame, var, None, *self.all_options, command=lambda val, idx=i: self.update_column(idx, val))
                self.dropdowns.append((var, menu))  # Store tuple of (var, menu) for updating later
            menu.grid(row=0, column=i, padx=2, sticky='we')
            # self.dropdowns.append(var)

        # Configure column widths and weights to match Treeview and allow expansion
        for i in range(5):  # Configure all 5 columns
            header_frame.grid_columnconfigure(i, weight=1, minsize=self.attributes_tree.column(self.columns[i], 'width'))
            frm.grid_columnconfigure(i, weight=1)
        frm.grid_rowconfigure(0, weight=0)  # Header row doesn't expand vertically
        frm.grid_rowconfigure(1, weight=1)  # Treeview row expands vertically

        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)

        self.attributes_tree.bind("<Button-3>", self.show_paste_menu)
        self.attributes_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='jobname'))
        self.attributes_tree.bind("<Double-1>", self.on_double_click)
        self.attributes_tree.bind("<Control-z>", self.undo)
        self.attributes_tree.bind("<Control-y>", self.redo)
        self.attributes_tree.bind("<Control-a>", self.select_all)
        self.attributes_tree.bind("<Delete>", self.delete_selected)
        self.attributes_tree.bind("<BackSpace>", self.delete_selected)
        self.attributes_tree.bind("<<TreeviewSelect>>", self.on_treeview_select)
        self.attributes_tree.bind("<Control-c>", self.copy_selected)  # New binding for Ctrl+C
        # self.parent.bind("<Button-1>", self.close_popup)
        self.attributes_tree.bind("<Tab>", lambda e: self.on_tab_press(e, direction="forward"))
        self.attributes_tree.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, direction="backward"))


        # Frame for inline buttons
        button_frame = ttk.Frame(frm)
        button_frame.grid(row=2, column=0, columnspan=5, sticky='ew', pady=5)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        self.update_btn = ttk.Button(button_frame, text='Update Jobs', command=self.start_update, width=10)
        self.update_btn.grid(row=0, column=0, padx=(0, 2), sticky='e')

        self.see_jobs_btn = ttk.Button(button_frame, text='See Jobs', command=self.start_fetch_jobs, width=10)
        self.see_jobs_btn.grid(row=0, column=1, padx=(2, 0), sticky='w')
        self.export_btn = ttk.Button(button_frame, text='Export to Excel', command=self.export_to_excel, width=12)
        self.export_btn.grid(row=0, column=2, padx=(5, 0), sticky='w')
        self.copy_failed_btn = ttk.Button(frm, text='Copy Failed', command=self.copy_failed_jobs, state='disabled', width=10)
        self.copy_failed_btn.grid(row=3, column=0, columnspan=5, sticky='ew', pady=5)

        ttk.Label(frm, text='Update Message:').grid(row=4, column=0, sticky='w')
        self.update_message_entry = ttk.Entry(frm)
        self.update_message_entry.grid(row=4, column=1, columnspan=4, sticky='ew', padx=5)

        ttk.Label(frm, text="ARMT No:").grid(row=5, column=0, sticky="w", padx=5)
        self.armt_no_entry = ttk.Entry(frm)
        self.armt_no_entry.grid(row=5, column=1, columnspan=4, sticky='ew', padx=5)

        # Frame for toggling between log and process views
        self.view_frame = ttk.Frame(frm)
        self.view_frame.grid(row=6, column=0, columnspan=5, sticky='nsew', pady=5)
        self.view_frame.grid_rowconfigure(1, weight=1)
        self.view_frame.grid_columnconfigure(0, weight=1)

        # Update Log view
        self.log_label = ttk.Label(self.view_frame, text='Update Log:')
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box = tk.Text(self.view_frame, height=10, state='disabled')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

        # Job Script Process view (initially hidden)
        self.process_label = ttk.Label(self.view_frame, text='Job Script Process:')
        self.process_text = tk.Text(self.view_frame, height=10, state='disabled')
        self.ensure_empty_row()
    def update_column(self, col_index, value):
        """Update the column header and configuration based on the selected attribute."""
        self.column_configs[col_index] = value if value else None
        col_name = self.columns[col_index]
        if value in ['EXPIR', 'LIST_T', 'LINE_COUNT']:
            self.attributes_tree.heading(col_name, text=f"{value} (R3)")
        elif value:
            self.attributes_tree.heading(col_name, text=value)
        else:
            self.attributes_tree.heading(col_name, text=col_name)
        self.log(f"Column {col_index} set to {value}")
    def update_dropdowns(self):
        """Update dropdown menus with the current all_options."""
        for i, (var, menu) in enumerate(self.dropdowns, 1):  # Skip jobname column
            menu['menu'].delete(0, 'end')  # Clear existing menu options
            for option in self.all_options:
                menu['menu'].add_command(label=option, command=tk._setit(var, option, lambda val, idx=i: self.update_column(idx, val)))

    def load_config(self, config):
        if 'ARMT_NO' in config:
            self.armt_no_entry.insert(0, config['ARMT_NO'])
        if 'COLUMN_CONFIGS' in config:
            for i, attr in enumerate(config['COLUMN_CONFIGS'][1:], 1):
                if attr in self.all_options:
                    self.dropdowns[i-1][0].set(attr)
                    self.update_column(i, attr)

    def save_config(self):
        return {
            'ARMT_NO': self.armt_no_entry.get(),
            'COLUMN_CONFIGS': self.column_configs
        }

    def show_paste_menu(self, event):
        self.clicked_column = self.attributes_tree.identify_column(event.x)[1:]
        self.paste_menu.tk_popup(event.x_root, event.y_root)

    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()

        columns = self.attributes_tree['columns']
        if start_column is None:
            if hasattr(self, 'clicked_column') and self.clicked_column:
                col_index = int(self.clicked_column) - 1
                start_column = columns[col_index]
            else:
                start_column = 'jobname'
        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return

        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return

            existing_rows = list(self.attributes_tree.get_children())
            max_columns = len(columns) - start_col_index

            for i, line in enumerate(lines):
                clipboard_values = line.split('\t') if '\t' in line else [line]
                clipboard_values = clipboard_values[:max_columns]
                while len(clipboard_values) < max_columns:
                    clipboard_values.append('')

                if i < len(existing_rows):
                    row_id = existing_rows[i]
                    current_values = list(self.attributes_tree.item(row_id)['values'])
                    for j, value in enumerate(clipboard_values):
                        current_values[start_col_index + j] = value.strip()
                    self.attributes_tree.item(row_id, values=current_values)
                else:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(clipboard_values):
                        new_values[start_col_index + j] = value.strip()
                    self.attributes_tree.insert('', 'end', values=new_values)

            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
            # Update job_processes: keep only jobs still in Treeview
            current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
            jobs_to_remove = set(self.job_processes.keys()) - current_jobs
            for job in jobs_to_remove:
                self.job_processes.pop(job, None)
            
            # Clear process text to reflect updated state
            self.process_text.config(state='normal')
            self.process_text.delete(1.0, tk.END)
            self.process_text.insert(tk.END, "No process data available. Click 'See Jobs' to fetch.")
            self.process_text.config(state='disabled') 
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
            self.undo_stack.pop()
        self.ensure_empty_row()

    def copy_selected(self, event=None):
        """Copy selected Treeview rows to the clipboard as tab-separated values."""
        selected = self.attributes_tree.selection()
        if not selected:
            self.log("No rows selected to copy.", tag='info')
            return 'break'

        rows_data = []
        for row_id in selected:
            values = self.attributes_tree.item(row_id)['values']
            # Only copy non-empty rows (skip if all values are empty)
            if any(val.strip() for val in values):
                rows_data.append('\t'.join(str(val) for val in values))

        if rows_data:
            clipboard_text = '\n'.join(rows_data)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(clipboard_text)
            self.parent.update()
            self.log(f"Copied {len(rows_data)} row(s) to clipboard.", tag='info')
        else:
            self.log("No non-empty rows selected to copy.", tag='info')

        return 'break'
    def ensure_empty_row(self):
        """Ensure there is always one empty row at the end of the Treeview."""
        columns = self.attributes_tree['columns']
        children = self.attributes_tree.get_children()
        if children:
            last_row_values = self.attributes_tree.item(children[-1])['values']
            last_row_empty = all(val == '' for val in last_row_values)
            if last_row_empty:
                return
        empty_row = [''] * len(columns)
        self.attributes_tree.insert('', 'end', values=empty_row)
    def on_double_click(self, event):
        row_id = self.attributes_tree.identify_row(event.y)
        if not row_id:
            return
        column = self.attributes_tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.attributes_tree['columns'][column_num]
        self.edit_cell(row_id, column_name, column_num)

    def edit_cell(self, row_id, column_name, column_num):
        """Open an Entry widget for editing a specific cell."""
        if self.current_entry:
            self.current_entry.destroy()
        x, y, width, height = self.attributes_tree.bbox(row_id, f"#{column_num + 1}")
        self.current_entry = ttk.Entry(self.parent)
        self.current_entry.place(x=x, y=y, width=width, height=height)
        current_value = self.attributes_tree.item(row_id)['values'][column_num]
        self.current_entry.insert(0, current_value)
        self.current_entry.select_range(0, tk.END)
        self.current_entry.focus()
        self.current_entry.bind("<Return>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<FocusOut>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<Escape>", lambda e: self.current_entry.destroy())
        self.current_entry.bind("<Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="forward"))
        self.current_entry.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="backward"))

    def on_tab_press(self, event, current_row_id=None, current_col_num=None, direction="forward"):
        """Handle Tab (forward) and Shift+Tab (backward) key presses to move to the next or previous cell."""
        if current_row_id is None or current_col_num is None:
            # Tab or Shift+Tab pressed directly in Treeview
            selected = self.attributes_tree.selection()
            if not selected:
                return 'break'
            current_row_id = selected[0]
            current_col_num = 0 if direction == "forward" else len(self.attributes_tree['columns']) - 1
        else:
            # Tab or Shift+Tab pressed in Entry widget, save current edit
            self.save_edit(self.current_entry, current_row_id, self.attributes_tree['columns'][current_col_num])
        
        columns = self.attributes_tree['columns']
        children = self.attributes_tree.get_children()
        current_row_index = children.index(current_row_id)
        
        # Check if the last row is empty to limit navigation
        last_row_empty = all(val == '' for val in self.attributes_tree.item(children[-1])['values'])
        
        if direction == "forward":
            # Move to next column or next row
            if current_col_num < len(columns) - 1:
                next_col_num = current_col_num + 1
                next_row_id = current_row_id
            else:
                # Move to the first column of the next row
                if current_row_index < len(children) - 1:
                    next_row_id = children[current_row_index + 1]
                    next_col_num = 0
                elif current_row_index == len(children) - 1 and not last_row_empty:
                    next_row_id = current_row_id
                    next_col_num = current_col_num
                else:
                    return 'break'  # Stay at last cell if last row is empty
        else:  # direction == "backward"
            # Move to previous column or previous row
            if current_col_num > 0:
                next_col_num = current_col_num - 1
                next_row_id = current_row_id
            else:
                # Move to the last column of the previous row
                if current_row_index > 0:
                    next_row_id = children[current_row_index - 1]
                    next_col_num = len(columns) - 1
                else:
                    return 'break'  # Stay at first cell
        
        next_col_name = columns[next_col_num]
        self.edit_cell(next_row_id, next_col_name, next_col_num)
        return 'break'


    def save_edit(self, entry, row_id, column_name):
        new_value = entry.get().strip()
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.attributes_tree.set(row_id, column_name, new_value)
                # If jobname was edited, update job_processes
        # if column_name == 'jobname':
        #     old_jobname = self.attributes_tree.item(row_id)['values'][0].strip()
        #     if old_jobname in self.job_processes and old_jobname != new_value:
        #         self.job_processes[new_value] = self.job_processes.pop(old_jobname)
        
        # # Update job_processes: keep only jobs still in Treeview
        # current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
        # jobs_to_remove = set(self.job_processes.keys()) - current_jobs
        # for job in jobs_to_remove:
        #     self.job_processes.pop(job, None)
        
        # Clear process text to reflect updated state
        self.process_text.config(state='normal')
        self.process_text.delete(1.0, tk.END)
        # self.process_text.insert(tk.END, "No process data available. Click 'See Jobs' to fetch.")
        self.process_text.config(state='disabled')
        entry.destroy()
        self.ensure_empty_row()


    def save_state(self):
        return [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]

    def restore_state(self, state):
        self.attributes_tree.delete(*self.attributes_tree.get_children())
        for values in state:
            self.attributes_tree.insert('', 'end', values=values)

    def undo(self, event):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def select_all(self, event):
        self.attributes_tree.selection_set(self.attributes_tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.attributes_tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.attributes_tree.delete(item)
            # Update job_processes: keep only jobs still in Treeview
            current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
            jobs_to_remove = set(self.job_processes.keys()) - current_jobs
            for job in jobs_to_remove:
                self.job_processes.pop(job, None)
            
            # Clear process text to reflect updated state
            self.process_text.config(state='normal')
            self.process_text.delete(1.0, tk.END)
            self.process_text.insert(tk.END, "No process data available. Click 'See Jobs' to fetch.")
            self.process_text.config(state='disabled')
            self.ensure_empty_row()

        return 'break'

    def show_log_view(self):
        """Show the Update Log view and hide the Job Script Process view."""
        self.process_label.grid_remove()
        self.process_text.grid_remove()
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

    def show_process_view(self):
        """Show the Job Script Process view and hide the Update Log view."""
        self.log_label.grid_remove()
        self.log_box.grid_remove()
        self.process_label.grid(row=0, column=0, sticky='nw')
        self.process_text.grid(row=1, column=0, sticky='nsew', padx=5)

    def start_update(self):
        self.update_btn.config(state='disabled')
                # Clear process text and job processes to avoid stale data
        self.process_text.config(state='normal')
        self.process_text.delete(1.0, tk.END)
        self.process_text.config(state='disabled')
        self.job_processes.clear()
        self.show_log_view()  # Show log view when updating
        threading.Thread(target=self.execute_update, daemon=True).start()

    def execute_update(self):
        self.failed_jobs = []
        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No data to update.")
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        name = self.entries['NAME'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
        except requests.exceptions.HTTPError as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        update_message = self.update_message_entry.get().strip()
        if update_message:

            armt_no = self.armt_no_entry.get().strip()
            current_date = date.today().strftime("%d/%m/%Y")
            doku_entry = f"{armt_no}, {name}, {current_date}, {update_message}"

        for row in rows:
            jobname = str(row[0]).strip()
            if not jobname:
                self.log("Skipping row with empty jobname")
                continue

            self.log(f"Processing {jobname}")
            try:
                resp = automic.getObjects(client_id=cid, object_name=jobname)
                if resp.status != 200:
                    self.log(f"Failed to fetch {jobname}: {resp.status}")
                    self.failed_jobs.append(jobname)
                    continue
                if 'data' not in resp.response or 'jobs' not in resp.response['data']:
                    self.log(f"Error: Object {jobname} response missing required keys")
                    self.failed_jobs.append(jobname)
                    continue
                jobs = resp.response["data"]["jobs"]
                updated = False

                for script in jobs.get('scripts', []):
                    if not script.get('process'):
                        continue
                    process_lines = script['process']
                    # Track which attributes need to be inserted
                    attrs_to_insert = []
                    # Track which attributes are already present
                    attrs_found = set()
                    for i, line in enumerate(process_lines):
                        if line is None or line.strip().startswith('!'):
                            continue
                        for col_idx in range(1, 5):
                            attr_value = str(row[col_idx]).strip()
                            attr_name = self.column_configs[col_idx]
                            if not attr_name or not attr_value:
                                continue
                            if attr_name in ['EXPIR', 'LIST_T', 'LINE_COUNT']:
                                if line.startswith("R3_ACTIVATE_REPORT"):
                                    parts = line.split(",")
                                    for j, part in enumerate(parts):
                                        if f"{attr_name}" in part:
                                            parts[j] = f'{attr_name}={attr_value}'
                                            process_lines[i] = ",".join(parts)
                                            self.log(f"Updated {attr_name} (R3) to {attr_value} for {jobname}")
                                            updated = True
                                            attrs_found.add(attr_name)
                                            break
                                    if attr_name not in attrs_found:
                                        parts.append(f'{attr_name}={attr_value}')
                                        process_lines[i] = ",".join(parts)
                                        self.log(f"Added {attr_name} (R3) with {attr_value} for {jobname}")

                                        attrs_found.add(attr_name)
                                        updated = True

                            else:
                                # if attr_name == 'LOGIN' and cid == 1111:
                                #     attr_value = f"LOGIN_R3_060_{attr_value}"
                                prefixes = ['PUT_ATT', 'SET', 'RSET', 'PSET']
                                for prefix in prefixes:
                                    pattern = re.compile(rf'(^:\s*{prefix})\s*{re.escape(attr_name)}\s*=')

                                    if re.match(pattern, line.strip()):
                                        attrs_found.add(attr_name)
                                        value_start = line.index('=') + 1
                                        new_line = f"{line[:value_start]}\"{attr_value}\""
                                        process_lines[i] = new_line
                                        self.log(f"Updated {attr_name} to {attr_value} for {jobname} with {prefix}")
                                        updated = True
                                        break
                                        # Insert new :PUT_ATT lines for attributes that weren't found
                    for col_idx in range(1, 5):
                        attr_value = str(row[col_idx]).strip()
                        attr_name = self.column_configs[col_idx]
                        if attr_value and attr_name and attr_name not in attrs_found:
                            prefix = ':SET' if (attr_name.startswith('&') or attr_name.startswith('#')) and attr_name.endswith('#') else ':PUT_ATT'
                            new_line = f"{prefix} {attr_name}=\"{attr_value}\""
                            # Insert the new line in the middle of process_lines
                            insert_index = len(process_lines) // 2 if process_lines else 0
                            process_lines.insert(insert_index, new_line)
                            self.log(f"Inserted {attr_name}={attr_value} for {jobname} with {prefix}")
                            updated = True
                doku_found = False
                if update_message:
                    for doc in jobs.get('documentation', []):
                        if cid ==1111 and 'Doku' in doc:
                            doku_list = doc['Doku']
                            if isinstance(doku_list, list):
                                doku_list.append(doku_entry)
                                doku_found = True
                                break
                        elif '_STRUKTUR' in doc:
                            struktur = doc['_STRUKTUR']
                            for i, line in enumerate(struktur):
                                if '</HINTS_CHARACTERISTICS>' in line:
                                    # Split before the closing tag
                                    content, tag = line.split('</HINTS_CHARACTERISTICS>')
                                    # print(content)
                                    # Append , and 'your_defined_string'
                                    new_content = f'{content.strip()}' 
                                    # Rebuild the line
                                    struktur[i] = new_content
                                    new_line =  f'{doku_entry}</HINTS_CHARACTERISTICS>{tag}'
                                    struktur.insert(i+1,new_line)
                                    break
                    # if not doku_found:
                    #     jobs.setdefault('documentation', []).append({'Doku': [doku_entry]})
                print(resp.response)
                if updated or update_message:
                    resp_update = automic.postObjects(client_id=cid, body=resp.response, query="overwrite_existing_objects=true")
                    if resp_update.status is None:
                        self.log(f"Successfully updated {jobname}")
                    else:
                        self.log(f"Failed to update {jobname}: {resp_update.status}")
                        self.failed_jobs.append(jobname)
                else:
                    self.log(f"No updates applied for {jobname}")
                    self.failed_jobs.append(jobname)

            except Exception as e:
                self.log(f"Error updating {jobname}: {str(e)}")
                self.failed_jobs.append(jobname)

        self.log("All updates applied.")
        self.parent.after(0, lambda: self.update_btn.config(state='normal'))
        self.parent.after(0, lambda: self.copy_failed_btn.config(state='normal' if self.failed_jobs else 'disabled'))

    def start_fetch_jobs(self):
        self.see_jobs_btn.config(state='disabled')
        self.show_process_view()  # Show process view when fetching jobs
        threading.Thread(target=self.fetch_jobs, daemon=True).start()
    def fetch_job_data(self,jobname, cid, api_url, auth):
        """Fetch and process data for a single job, returning results for the main thread."""
        try:
            resp = automic.getObjects(client_id=cid, object_name=jobname)
            if resp.status != 200:
                return jobname, None, None, None, f"Failed to fetch {jobname}: {resp.status}"
            if 'data' not in resp.response or 'jobs' not in resp.response['data']:
                return jobname, None, None, None, f"Error: Object {jobname} response missing required keys"
            jobs = resp.response["data"]["jobs"]
            process_lines = []
            job_attributes = {}
            dynamic_attributes = set()

            for script in jobs.get('scripts', []):
                if not script.get('process'):
                    continue
                process_lines = [line for line in script['process'] if line is not None]
                if process_lines:
                    for line in process_lines:
                        line = line.strip()
                        if line.startswith('!'):
                            continue
                        for prefix in ['PUT_ATT', 'SET', 'RSET', 'PSET']:
                            match = re.match(rf'^:\s*{prefix}\s*(\S+)\s*=\s*(?:"([^"]*)"|(\S+))', line, re.IGNORECASE)
                            if match:
                                attr_name, quoted_value, unquoted_value = match.groups()
                                value = quoted_value if quoted_value is not None else unquoted_value
                                dynamic_attributes.add(attr_name)
                                job_attributes[attr_name] = value
            if process_lines:
                job_attributes["R3_ACTIVATE_REPORT"] = process_lines[-1]
            return jobname, process_lines, job_attributes, dynamic_attributes, None
        except Exception as e:
            print(e)
            return jobname, None, None, None, f"Error fetching {jobname}: {str(e)}"
    def fetch_jobs(self):
        """Fetch job data in parallel using ThreadPoolExecutor and update UI when done."""
        self.job_processes.clear()
        self.job_attributes.clear()

        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No jobs to fetch.")
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
        except requests.exceptions.HTTPError as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials"))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)
            return

        jobnames = [str(row[0]).strip() for row in rows if str(row[0]).strip()]

        def background_task():
            """Run parallel job fetching in a background thread."""
            job_processes = {}
            job_attributes = {}
            dynamic_attributes = set()

            with ThreadPoolExecutor(max_workers=10) as executor:
                # Submit all jobs to the executor
                future_to_jobname = {
                    executor.submit(self.fetch_job_data, jobname, cid, api_url, auth): jobname
                    for jobname in jobnames
                }
                # Process results as they complete
                for future in as_completed(future_to_jobname):
                    jobname = future_to_jobname[future]
                    try:
                        jobname, process_lines, job_attrs, dyn_attrs, error = future.result()
                        if error:
                            self.log(error)
                        else:
                            if process_lines:
                                job_processes[jobname] = process_lines
                            if job_attrs:
                                job_attributes[jobname] = job_attrs
                            dynamic_attributes.update(dyn_attrs)
                    except Exception as e:
                        self.log(f"Error processing {jobname}: {str(e)}")

            # Update instance variables after all jobs are processed
            self.job_processes.update(job_processes)
            self.job_attributes.update(job_attributes)
            self.all_options = [''] + sorted(list(dynamic_attributes)) + ['EXPIR', 'LIST_T', 'LINE_COUNT']

            # Schedule UI updates on the main thread
            self.parent.after(0, self.update_dropdowns)
            self.parent.after(0, lambda: self.log(f"Updated dropdown options: {self.all_options}"))
            self.parent.after(0, lambda: self.log("All jobs fetched."))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)

        # Start the background task
        threading.Thread(target=background_task, daemon=True).start()
    def export_to_excel(self):
        """Export job attributes to an Excel file."""
        if not self.job_attributes:
            messagebox.showinfo("Export Info", "No job attribute data to export. Click 'See Jobs' to fetch data.")
            self.log("No job attribute data to export.")
            return

        try:
            # Prepare data for DataFrame
            job_names = sorted(self.job_attributes.keys())
            all_attrs = sorted(set(attr for job_attrs in self.job_attributes.values() for attr in job_attrs))
            data = []
            for job in job_names:
                row = {'Job Name': job}
                for attr in all_attrs:
                    row[attr] = self.job_attributes[job].get(attr, '')
                data.append(row)

            # Create DataFrame
            df = pd.DataFrame(data, columns=['Job Name'] + all_attrs)
            # Generate timestamp for filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"job_attributes_{timestamp}.xlsx"
            # Save to Excel
            df.to_excel(filename, index=False)
            self.log(f"Exported job attributes to {filename}")
            messagebox.showinfo("Export Success", f"Successfully exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to Excel: {str(e)}")
            messagebox.showerror("Export Error", f"Failed to export to Excel: {str(e)}")

    def update_process_text_for_selection(self):
            """Update process_text based on the current Treeview selection."""
            selected = self.attributes_tree.selection()
            self.process_text.config(state='normal')
            self.process_text.delete(1.0, tk.END)
            if not selected:
                self.process_text.config(state='disabled')
                return

            row_id = selected[0]
            jobname = self.attributes_tree.item(row_id)['values'][0].strip()
            if jobname in self.job_processes:
                process_lines = [str(line) for line in self.job_processes[jobname]]  # Ensure all are strings
                self.process_text.insert(tk.END, "\n".join(process_lines))
            else:
                self.process_text.insert(tk.END, "No process data available. Click 'See Jobs' to fetch.")
            self.process_text.config(state='disabled')
    def on_treeview_select(self, event):
        """Handle Treeview selection changes."""
        self.update_process_text_for_selection()

    def copy_failed_jobs(self):
        if self.failed_jobs:
            failed_list = "\n".join(self.failed_jobs)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(failed_list)
            self.log("Copied failed jobs to clipboard.")
        else:
            messagebox.showinfo("Info", "No failed jobs to copy.")

    def log(self, msg):
        self.parent.after(0, lambda: self._log(msg))

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')


class JobpUpdater:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.undo_stack = []
        self.redo_stack = []
        self.failed_jobs = []
        self.job_processes = {}  # Store script['process'] for each job
        self.job_attributes = {}  # Store {jobname: {attr_name: value}}

        # Available options for dropdowns
        self.all_options = [''] + ['SMT_ON', 'SMT_OFF', 'EMAIL_ON', 'EMAIL_OFF', 'EXTRA_LINE1', 'EXTRA_LINE2', 'EXTRA_LINE3']
        self.current_entry = None

        self.column_configs = ['jobname', None, None, None, None, None]  # Track selected attributes
        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        # Frame for dropdown headers
        header_frame = ttk.Frame(frm)
        header_frame.grid(row=0, column=0, columnspan=5, sticky='nsew')

        # Define initial columns
        self.columns = ('jobname', 'attribute1', 'attribute2', 'attribute3', 'attribute4','attribute5')
        self.attributes_tree = ttk.Treeview(frm, columns=self.columns, show='headings')
        self.attributes_tree.heading('jobname', text='Job Name')
        self.attributes_tree.column('jobname', width=200, stretch=True)
        for col in self.columns[1:]:
            self.attributes_tree.heading(col, text=col)
            self.attributes_tree.column(col, width=150, stretch=True)
        self.attributes_tree.grid(row=1, column=0, columnspan=6, sticky='nsew')

        # Combined options for all dropdowns
        # all_options = [''] + ['SMT_ON', 'SMT_OFF', 'EMAIL_ON', 'EMAIL_OFF', 'EXTRA_LINE1', 'EXTRA_LINE2', 'EXTRA_LINE3']
        self.dropdowns = []
        ttk.Label(header_frame, text="Select Attributes:").grid(row=0, column=0, sticky='w', pady=5)
        for i, col in enumerate(self.columns, 0):
            var = tk.StringVar()
            if i == 0:  # jobname column
                menu = ttk.OptionMenu(header_frame, var, 'Job Name', 'Job Name')
                menu.config(state='disabled')  # Disable dropdown for jobname
            else:
                menu = ttk.OptionMenu(header_frame, var, None, *self.all_options, command=lambda val, idx=i: self.update_column(idx, val))
                self.dropdowns.append((var, menu))  # Store tuple of (var, menu) for updating later
            menu.grid(row=0, column=i, padx=2, sticky='we')
            # self.dropdowns.append(var)

        # Configure column widths and weights to match Treeview and allow expansion
        for i in range(6):  # Configure all 5 columns
            header_frame.grid_columnconfigure(i, weight=1, minsize=self.attributes_tree.column(self.columns[i], 'width'))
            frm.grid_columnconfigure(i, weight=1)
        frm.grid_rowconfigure(0, weight=0)  # Header row doesn't expand vertically
        frm.grid_rowconfigure(1, weight=1)  # Treeview row expands vertically

        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)

        self.attributes_tree.bind("<Button-3>", self.show_paste_menu)
        self.attributes_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='jobname'))
        self.attributes_tree.bind("<Double-1>", self.on_double_click)
        self.attributes_tree.bind("<Control-z>", self.undo)
        self.attributes_tree.bind("<Control-y>", self.redo)
        self.attributes_tree.bind("<Control-a>", self.select_all)
        self.attributes_tree.bind("<Delete>", self.delete_selected)
        self.attributes_tree.bind("<BackSpace>", self.delete_selected)
        self.attributes_tree.bind("<<TreeviewSelect>>", self.on_treeview_select)
        self.attributes_tree.bind("<Control-c>", self.copy_selected)  # New binding for Ctrl+C
        # self.parent.bind("<Button-1>", self.close_popup)
        self.attributes_tree.bind("<Tab>", lambda e: self.on_tab_press(e, direction="forward"))
        self.attributes_tree.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, direction="backward"))

        # Frame for inline buttons
        button_frame = ttk.Frame(frm)
        button_frame.grid(row=2, column=0, columnspan=5, sticky='ew', pady=5)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        self.update_btn = ttk.Button(button_frame, text='Update Jobs', command=self.start_update, width=10)
        self.update_btn.grid(row=0, column=0, padx=(0, 2), sticky='e')

        self.see_jobs_btn = ttk.Button(button_frame, text='See Jobp', command=self.start_fetch_jobs, width=10)
        self.see_jobs_btn.grid(row=0, column=1, padx=(2, 0), sticky='w')
        self.export_btn = ttk.Button(button_frame, text='Export to Excel', command=self.export_to_excel, width=12)
        self.export_btn.grid(row=0, column=2, padx=(5, 0), sticky='w')

        self.copy_failed_btn = ttk.Button(frm, text='Copy Failed', command=self.copy_failed_jobs, state='disabled', width=10)
        self.copy_failed_btn.grid(row=3, column=0, columnspan=5, sticky='ew', pady=5)

        ttk.Label(frm, text='Update Message:').grid(row=4, column=0, sticky='w')
        self.update_message_entry = ttk.Entry(frm)
        self.update_message_entry.grid(row=4, column=1, columnspan=4, sticky='ew', padx=5)

        ttk.Label(frm, text="ARMT No:").grid(row=5, column=0, sticky="w", padx=5)
        self.armt_no_entry = ttk.Entry(frm)
        self.armt_no_entry.grid(row=5, column=1, columnspan=4, sticky='ew', padx=5)

        # Frame for toggling between log and process views
        self.view_frame = ttk.Frame(frm)
        self.view_frame.grid(row=6, column=0, columnspan=5, sticky='nsew', pady=5)
        self.view_frame.grid_rowconfigure(1, weight=1)
        self.view_frame.grid_columnconfigure(0, weight=1)

        # Update Log view
        self.log_label = ttk.Label(self.view_frame, text='Update Log:')
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box = tk.Text(self.view_frame, height=10, state='disabled')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

        # Job Script Process view (initially hidden)
        self.process_label = ttk.Label(self.view_frame, text='Job Script Process:')
        self.process_text = tk.Text(self.view_frame, height=10, state='disabled')
        self.ensure_empty_row()

    def update_column(self, col_index, value):
        """Update the column header and configuration based on the selected attribute."""
        self.column_configs[col_index] = value if value else None
        col_name = self.columns[col_index]  # Use the column name from the columns tuple
        if value in ['EXTRA_LINE1', 'EXTRA_LINE2', 'EXTRA_LINE3']:
            # Treat as r3_attribute type
            self.attributes_tree.heading(col_name, text=f"{value} (SCRIPT)")
        elif value:
            # Treat as regular attribute
            self.attributes_tree.heading(col_name, text=value)
        else:
            # Default or empty selection
            self.attributes_tree.heading(col_name, text=col_name)
        self.log(f"Column {col_index} set to {value}")
    def update_dropdowns(self):
        """Update dropdown menus with the current all_options."""
        for i, (var, menu) in enumerate(self.dropdowns, 1):  # Skip jobname column
            menu['menu'].delete(0, 'end')  # Clear existing menu options
            for option in self.all_options:
                menu['menu'].add_command(label=option, command=tk._setit(var, option, lambda val, idx=i: self.update_column(idx, val)))


    def load_config(self, config):
        if 'ARMT_NO' in config:
            self.armt_no_entry.insert(0, config['ARMT_NO'])
        if 'COLUMN_CONFIGS' in config:
            for i, attr in enumerate(config['COLUMN_CONFIGS'][1:], 1):
                if attr in self.all_options:
                    self.dropdowns[i-1][0].set(attr)
                    self.update_column(i, attr)

    def save_config(self):
        return {
            'ARMT_NO': self.armt_no_entry.get(),
            'COLUMN_CONFIGS': self.column_configs
        }

    def show_paste_menu(self, event):
        self.clicked_column = self.attributes_tree.identify_column(event.x)[1:]
        self.paste_menu.tk_popup(event.x_root, event.y_root)

    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()

        columns = self.attributes_tree['columns']
        if start_column is None:
            if hasattr(self, 'clicked_column') and self.clicked_column:
                col_index = int(self.clicked_column) - 1
                start_column = columns[col_index]
            else:
                start_column = 'jobname'
        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return

        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return

            existing_rows = list(self.attributes_tree.get_children())
            max_columns = len(columns) - start_col_index

            for i, line in enumerate(lines):
                clipboard_values = line.split('\t') if '\t' in line else [line]
                clipboard_values = clipboard_values[:max_columns]
                while len(clipboard_values) < max_columns:
                    clipboard_values.append('')

                if i < len(existing_rows):
                    row_id = existing_rows[i]
                    current_values = list(self.attributes_tree.item(row_id)['values'])
                    for j, value in enumerate(clipboard_values):
                        current_values[start_col_index + j] = value.strip()
                    self.attributes_tree.item(row_id, values=current_values)
                else:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(clipboard_values):
                        new_values[start_col_index + j] = value.strip()
                    self.attributes_tree.insert('', 'end', values=new_values)

            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
                        # Update job_processes: keep only jobs still in Treeview
            current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
            jobs_to_remove = set(self.job_processes.keys()) - current_jobs
            for job in jobs_to_remove:
                self.job_processes.pop(job, None)
            
            # Clear process text to reflect updated state
            self.process_text.config(state='normal')
            self.process_text.delete(1.0, tk.END)
            self.process_text.insert(tk.END, "No process data available. Click 'See Jobp' to fetch.")
            self.process_text.config(state='disabled') 
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
            self.undo_stack.pop()
        self.ensure_empty_row()

    def copy_selected(self, event=None):
        """Copy selected Treeview rows to the clipboard as tab-separated values."""
        selected = self.attributes_tree.selection()
        if not selected:
            self.log("No rows selected to copy.", tag='info')
            return 'break'

        rows_data = []
        for row_id in selected:
            values = self.attributes_tree.item(row_id)['values']
            # Only copy non-empty rows (skip if all values are empty)
            if any(val.strip() for val in values):
                rows_data.append('\t'.join(str(val) for val in values))

        if rows_data:
            clipboard_text = '\n'.join(rows_data)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(clipboard_text)
            self.parent.update()
            self.log(f"Copied {len(rows_data)} row(s) to clipboard.", tag='info')
        else:
            self.log("No non-empty rows selected to copy.", tag='info')

        return 'break'
    def ensure_empty_row(self):
        """Ensure there is always one empty row at the end of the Treeview."""
        columns = self.attributes_tree['columns']
        children = self.attributes_tree.get_children()
        if children:
            last_row_values = self.attributes_tree.item(children[-1])['values']
            last_row_empty = all(val == '' for val in last_row_values)
            if last_row_empty:
                return
        empty_row = [''] * len(columns)
        self.attributes_tree.insert('', 'end', values=empty_row)
    def on_double_click(self, event):
        row_id = self.attributes_tree.identify_row(event.y)
        if not row_id:
            return
        column = self.attributes_tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.attributes_tree['columns'][column_num]
        self.edit_cell(row_id, column_name, column_num)

    def edit_cell(self, row_id, column_name, column_num):
        """Open an Entry widget for editing a specific cell."""
        if self.current_entry:
            self.current_entry.destroy()
        x, y, width, height = self.attributes_tree.bbox(row_id, f"#{column_num + 1}")
        self.current_entry = ttk.Entry(self.parent)
        self.current_entry.place(x=x, y=y, width=width, height=height)
        current_value = self.attributes_tree.item(row_id)['values'][column_num]
        self.current_entry.insert(0, current_value)
        self.current_entry.select_range(0, tk.END)
        self.current_entry.focus()
        self.current_entry.bind("<Return>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<FocusOut>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<Escape>", lambda e: self.current_entry.destroy())
        self.current_entry.bind("<Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="forward"))
        self.current_entry.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="backward"))

    def on_tab_press(self, event, current_row_id=None, current_col_num=None, direction="forward"):
        """Handle Tab (forward) and Shift+Tab (backward) key presses to move to the next or previous cell."""
        if current_row_id is None or current_col_num is None:
            # Tab or Shift+Tab pressed directly in Treeview
            selected = self.attributes_tree.selection()
            if not selected:
                return 'break'
            current_row_id = selected[0]
            current_col_num = 0 if direction == "forward" else len(self.attributes_tree['columns']) - 1
        else:
            # Tab or Shift+Tab pressed in Entry widget, save current edit
            self.save_edit(self.current_entry, current_row_id, self.attributes_tree['columns'][current_col_num])
        
        columns = self.attributes_tree['columns']
        children = self.attributes_tree.get_children()
        current_row_index = children.index(current_row_id)
        
        # Check if the last row is empty to limit navigation
        last_row_empty = all(val == '' for val in self.attributes_tree.item(children[-1])['values'])
        
        if direction == "forward":
            # Move to next column or next row
            if current_col_num < len(columns) - 1:
                next_col_num = current_col_num + 1
                next_row_id = current_row_id
            else:
                # Move to the first column of the next row
                if current_row_index < len(children) - 1:
                    next_row_id = children[current_row_index + 1]
                    next_col_num = 0
                elif current_row_index == len(children) - 1 and not last_row_empty:
                    next_row_id = current_row_id
                    next_col_num = current_col_num
                else:
                    return 'break'  # Stay at last cell if last row is empty
        else:  # direction == "backward"
            # Move to previous column or previous row
            if current_col_num > 0:
                next_col_num = current_col_num - 1
                next_row_id = current_row_id
            else:
                # Move to the last column of the previous row
                if current_row_index > 0:
                    next_row_id = children[current_row_index - 1]
                    next_col_num = len(columns) - 1
                else:
                    return 'break'  # Stay at first cell
        
        next_col_name = columns[next_col_num]
        self.edit_cell(next_row_id, next_col_name, next_col_num)
        return 'break'

    def _strip_quotes(self, value):
        """Strip all surrounding single quotes from a string, returning the clean value."""
        if isinstance(value, str):
            while value.startswith("'") and value.endswith("'") and len(value) >= 2:
                value = value[1:-1]
        return str(value)  # Ensure the result is a string
    def save_edit(self, entry, row_id, column_name):
        new_value = str(entry.get().strip())  # Get input as string
        quoted_value = f"{new_value}"  # Wrap in single quotes
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.attributes_tree.set(row_id, column_name, quoted_value)
        # Debug: Verify stored and displayed values
        saved_value = self.attributes_tree.item(row_id, 'values')[self.attributes_tree['columns'].index(column_name)]
        # self.log(f"Saved value: {saved_value}, Displayed: {self._strip_quotes(saved_value)}")
        # Clear process text
        self.process_text.config(state='normal')
        self.process_text.delete(1.0, tk.END)
        self.process_text.insert(tk.END, "No process data available. Click 'See Jobp' to fetch.")
        self.process_text.config(state='disabled')
        entry.destroy()
        self.ensure_empty_row()

    def save_state(self):
        return [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]

    def restore_state(self, state):
        self.attributes_tree.delete(*self.attributes_tree.get_children())
        for values in state:
            self.attributes_tree.insert('', 'end', values=values)

    def undo(self, event):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def select_all(self, event):
        self.attributes_tree.selection_set(self.attributes_tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.attributes_tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.attributes_tree.delete(item)
                            # Update job_processes: keep only jobs still in Treeview
            current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
            jobs_to_remove = set(self.job_processes.keys()) - current_jobs
            for job in jobs_to_remove:
                self.job_processes.pop(job, None)
            
            # Clear process text to reflect updated state
            self.process_text.config(state='normal')
            self.process_text.delete(1.0, tk.END)
            self.process_text.insert(tk.END, "No process data available. Click 'See Jobs' to fetch.")
            self.process_text.config(state='disabled')
            self.ensure_empty_row()

        return 'break'
    def show_log_view(self):
        """Show the Update Log view and hide the Job Script Process view."""
        self.process_label.grid_remove()
        self.process_text.grid_remove()
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

    def show_process_view(self):
        """Show the Job Script Process view and hide the Update Log view."""
        self.log_label.grid_remove()
        self.log_box.grid_remove()
        self.process_label.grid(row=0, column=0, sticky='nw')
        self.process_text.grid(row=1, column=0, sticky='nsew', padx=5)

    def start_update(self):
        # if None in self.column_configs[1:]:
        #     messagebox.showerror("Error", "Please select attributes for all columns.")
        #     return
        self.update_btn.config(state='disabled')
                        # Clear process text and job processes to avoid stale data
        self.process_text.config(state='normal')
        self.process_text.delete(1.0, tk.END)
        self.process_text.config(state='disabled')
        self.job_processes.clear()
        self.show_log_view()  # Show log view when updating
        threading.Thread(target=self.execute_update, daemon=True).start()

    def execute_update(self):
        self.failed_jobs = []
        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No data to update.")
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        name = self.entries['NAME'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
        except requests.exceptions.HTTPError as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        update_message = self.update_message_entry.get().strip()
        if update_message:

            armt_no = self.armt_no_entry.get().strip()
            current_date = date.today().strftime("%d/%m/%Y")
            doku_entry = f"{armt_no}, {name}, {current_date}, {update_message}"
        for row in rows:
            jobname = str(row[0]).strip()
            if not jobname:
                self.log("Skipping row with empty jobname")
                continue

            self.log(f"Processing {jobname}")
            try:
                resp = automic.getObjects(client_id=cid, object_name=jobname)
                if resp.status != 200:
                    self.log(f"Failed to fetch {jobname}: {resp.status}")
                    self.failed_jobs.append(jobname)
                    continue
                if 'data' not in resp.response or 'jobp' not in resp.response['data']:
                    self.log(f"Error: Object {jobname} response missing required keys")
                    self.failed_jobs.append(jobname)
                    continue
                jobp = resp.response["data"]["jobp"]
                updated = False
                # Update based on column configuration
                for col_idx in range(1, 6):  # Columns 1 to 8 (attribute1 to attribute4)
                    attr_name = self.column_configs[col_idx]
                    if attr_name:
                        if attr_name in ['SMT_ON', 'SMT_OFF', 'EMAIL_ON','EMAIL_OFF']:
                            for child in jobp.get('workflow_definitions', []):
                                if child['object_name'] == 'XX_XXXX_SMT_GET_VALUES_ERP' and attr_name == 'SMT_ON':
                                    child['active'] = 1
                                    self.log(f"Set active=1 for XX_XXXX_SMT_GET_VALUES_ERP in {jobname}")
                                    updated = True
                                elif child['object_name'] == 'XX_XXXX_SMT_GET_VALUES_ERP' and attr_name == 'SMT_OFF':
                                    child.pop('active', None)
                                    self.log(f"Removed active key for XX_XXXX_SMT_GET_VALUES_ERP in {jobname}")
                                    updated = True
                                elif child['object_name'] == 'CI_XXXX_XX99001_EMAIL_NOT_OK' and attr_name == 'EMAIL_ON':
                                    child['active'] = 1
                                    self.log(f"Set active=1 for CI_XXXX_XX99001_EMAIL_NOT_OK in {jobname}")
                                    updated = True
                                elif child['object_name'] == 'CI_XXXX_XX99001_EMAIL_NOT_OK' and attr_name == 'EMAIL_OFF':
                                    child.pop('active', None)
                                    self.log(f"Removed active key for CI_XXXX_XX99001_EMAIL_NOT_OK in {jobname}")
                                    updated = True
                        elif attr_name in ['EXTRA_LINE1', 'EXTRA_LINE2', 'EXTRA_LINE3']:
                            attr_value = str(row[col_idx]).strip()
                            # Treat as regular attribute
                            if attr_value:
                                new_line = f'{attr_value}'

                                if jobp.get('scripts'):
                                    process_lines = jobp['scripts']['process']
                                    process_lines.insert(0, new_line)
                                else:
                                    jobp['scripts'] = {'process': [new_line]}
                                self.log(f"Updated {attr_name} to {attr_value} for {jobname}")
                                updated = True
                        else:
                            attr_value = str(row[col_idx]).strip()
                            # Treat as regular attribute
                            if attr_value:
                                attrs_found = False
                                process_lines = jobp['scripts']['process']
                                for i, line in enumerate(process_lines):
                                    if line is None or line.strip().startswith('!'):
                                        continue
                                    prefixes = ['PUT_ATT', 'SET', 'RSET', 'PSET']
                                    for prefix in prefixes:
                                        pattern = re.compile(rf'(^:\s*(?i){prefix})\s*{re.escape(attr_name)}\s*=')

                                        if re.match(pattern, line.strip()):
                                            attrs_found = True
                                            value_start = line.index('=') + 1
                                            new_line = f"{line[:value_start]}\"{attr_value}\""
                                            process_lines[i] = new_line
                                            self.log(f"Updated {attr_name} to {attr_value} for {jobname} with {prefix}")
                                            updated = True
                                            break
                                if not attrs_found:
                                    prefix = ':SET' if (attr_name.startswith('&') or attr_name.startswith('#')) else ':PUT_ATT'
                                    new_line = f"{prefix} {attr_name}=\"{attr_value}\""
                                    insert_index = len(process_lines) // 2 if process_lines else 0
                                    process_lines.insert(insert_index, new_line)
                                    self.log(f"Inserted {attr_name}={attr_value} for {jobname} with {prefix}")
                                    updated = True
                # Append documentation
                if update_message:
                    
                    for doc in jobp.get('documentation', []):
                        if cid ==1111 and 'Doku' in doc:
                            doku_list = doc['Doku']
                            if isinstance(doku_list, list):
                                doku_list.append(doku_entry)
                                break
                        elif '_STRUKTUR' in doc:
                                struktur = doc['_STRUKTUR']
                                for i, line in enumerate(struktur):
                                    if '</HINTS_CHARACTERISTICS>' in line:
                                        # Split before the closing tag
                                        content, tag = line.split('</HINTS_CHARACTERISTICS>')
                                        # print(content)
                                        # Append , and 'your_defined_string'
                                        new_content = f'{content.strip()}' 
                                        # Rebuild the line
                                        struktur[i] = new_content
                                        new_line =  f'{doku_entry}</HINTS_CHARACTERISTICS>{tag}'
                                        struktur.insert(i+1,new_line)
                                        break                

                if updated or update_message:
                    resp_update = automic.postObjects(client_id=cid, body=resp.response, query="overwrite_existing_objects=true")
                    if resp_update.status is None:
                        self.log(f"Successfully updated {jobname}")
                    else:
                        self.log(f"Failed to update {jobname}: {resp_update.status}")
                        self.failed_jobs.append(jobname)
                else:
                    self.log(f"No updates applied for {jobname}")
                    self.failed_jobs.append(jobname)

            except Exception as e:
                self.log(f"Error updating {jobname}: {str(e)}")
                self.failed_jobs.append(jobname)

        self.log("All updates applied.")
        self.parent.after(0, lambda: self.update_btn.config(state='normal'))
        self.parent.after(0, lambda: self.copy_failed_btn.config(state='normal' if self.failed_jobs else 'disabled'))

    def start_fetch_jobs(self):
        self.see_jobs_btn.config(state='disabled')
        self.show_process_view()  # Show process view when fetching jobs
        threading.Thread(target=self.fetch_jobs, daemon=True).start()

    def fetch_jobs(self):
        self.job_processes.clear()
        self.job_attributes.clear()
        dynamic_attributes = set()

        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No jobs to fetch.")
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
        except requests.exceptions.HTTPError as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials"))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_process_text_for_selection)
            return

        def fetch_single_job(jobname):
            """Helper function to fetch a single job."""
            if not jobname:
                return None, None, None
            try:
                self.log(f"Fetching {jobname}")
                resp = automic.getObjects(client_id=cid, object_name=jobname)
                if resp.status != 200:
                    self.log(f"Failed to fetch {jobname}: {resp.status}")
                    return jobname, None, None
                if 'data' not in resp.response or 'jobp' not in resp.response['data']:
                    self.log(f"Error: Object {jobname} response missing required keys")
                    return jobname, None, None
                jobp = resp.response["data"]["jobp"]
                process_lines = None
                job_attributes = {}
                if jobp.get('scripts'):
                    script = jobp['scripts']
                    if script.get('process'):
                        process_lines = [line for line in script['process'] if line is not None]
                        if process_lines:
                            self.log(f"Stored process for {jobname}")
                            for line in process_lines:
                                line = line.strip()
                                if line.startswith('!'):
                                    continue
                                for prefix in ['PUT_ATT', 'SET', 'RSET', 'PSET']:
                                    pattern = re.compile(
                                        rf"""^:\s*{prefix}\s+         # “: SET ” (case-insensitive)
                                            ([&#]?[^\s=]+)\s*=\s*    # the attribute name
                                            (?:
                                                "([^"]*)"            # double-quoted value
                                            | '([^']*)'            # single-quoted value
                                            | (.+)                 # or everything else
                                            )
                                            $""",
                                        re.IGNORECASE | re.VERBOSE
                                    )
                                    match = pattern.match(line)
                                    if match:
                                        attr_name = match.group(1)
                                        value = next(g for g in match.groups()[1:] if g is not None)
                                        job_attributes[attr_name] = value
                                        dynamic_attributes.add(attr_name)
                                        self.log(f"Extracted {attr_name}={value} for job {jobname}")
                        else:
                            self.log(f"No valid process lines for {jobname}")
                return jobname, process_lines, job_attributes
            except Exception as e:
                self.log(f"Error fetching {jobname}: {str(e)}")
                return jobname, None, None

        # Use ThreadPoolExecutor to fetch jobs concurrently
        jobnames = [str(row[0]).strip() for row in rows if str(row[0]).strip()]
        max_workers = min(len(jobnames), 10)  # Limit to 10 threads to avoid overwhelming the server
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_jobname = {executor.submit(fetch_single_job, jobname): jobname for jobname in jobnames}
            for future in as_completed(future_to_jobname):
                jobname, process_lines, job_attrs = future.result()
                if process_lines is not None:
                    self.job_processes[jobname] = process_lines
                if job_attrs is not None:
                    self.job_attributes[jobname] = job_attrs

        # Update all_options with dynamic attributes
        self.all_options = [''] + sorted(list(dynamic_attributes)) + ['SMT_ON', 'SMT_OFF', 'EMAIL_ON', 'EMAIL_OFF', 'EXTRA_LINE1', 'EXTRA_LINE2', 'EXTRA_LINE3']
        self.parent.after(0, self.update_dropdowns)
        self.log(f"Updated dropdown options: {self.all_options}")
        self.log("All jobs fetched.")
        self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
        self.parent.after(0, self.update_process_text_for_selection)
    def export_to_excel(self):
        """Export job attributes to an Excel file."""
        if not self.job_attributes:
            messagebox.showinfo("Export Info", "No job attribute data to export. Click 'See Jobs' to fetch data.")
            self.log("No job attribute data to export.")
            return

        try:
            # Prepare data for DataFrame
            job_names = sorted(self.job_attributes.keys())
            all_attrs = sorted(set(attr for job_attrs in self.job_attributes.values() for attr in job_attrs))
            data = []
            for job in job_names:
                row = {'Jobp Name': job}
                for attr in all_attrs:
                    row[attr] = self.job_attributes[job].get(attr, '')
                data.append(row)

            # Create DataFrame
            df = pd.DataFrame(data, columns=['Jobp Name'] + all_attrs)
            # Generate timestamp for filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"jobp_attributes_{timestamp}.xlsx"
            # Save to Excel
            df.to_excel(filename, index=False)
            self.log(f"Exported job attributes to {filename}")
            messagebox.showinfo("Export Success", f"Successfully exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to Excel: {str(e)}")
            messagebox.showerror("Export Error", f"Failed to export to Excel: {str(e)}")
    def update_process_text_for_selection(self):
            """Update process_text based on the current Treeview selection."""
            selected = self.attributes_tree.selection()
            self.process_text.config(state='normal')
            self.process_text.delete(1.0, tk.END)
            if not selected:
                self.process_text.config(state='disabled')
                return

            row_id = selected[0]
            jobname = self.attributes_tree.item(row_id)['values'][0].strip()
            if jobname in self.job_processes:
                process_lines = [str(line) for line in self.job_processes[jobname]]  # Ensure all are strings
                self.process_text.insert(tk.END, "\n".join(process_lines))
            else:
                self.process_text.insert(tk.END, "No process data available. Click 'See Jobs' to fetch.")
            self.process_text.config(state='disabled')
    def on_treeview_select(self, event):
        """Handle Treeview selection changes."""
        self.update_process_text_for_selection()
    def copy_failed_jobs(self):
        if self.failed_jobs:
            failed_list = "\n".join(self.failed_jobs)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(failed_list)
            self.log("Copied failed jobp to clipboard.")
        else:
            messagebox.showinfo("Info", "No failed jobp to copy.")

    def log(self, msg):
        self.parent.after(0, lambda: self._log(msg))

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
class AutomicApp:
    TIMEOUT = 30
    MAX_WORKERS = 10  # Adjust based on API rate limits

    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.undo_stack = []
        self.redo_stack = []
        self.connection_lock = Lock()
        self.stop_flag = False
        self.color_map = {}
        self.usage_names_map = {}  # Store multiple usages per row_id
        self.active_popup = None
        self.all_rows = []
        self.current_entry = None  # Track the current Entry widget

        # Setup logging
        self.logger = logging.getLogger(__name__)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )

        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        # Top frame for buttons
        top_frame = ttk.Frame(frm)
        top_frame.grid(row=0, column=0, columnspan=4, sticky="ew", padx=5, pady=5)
        top_frame.grid_columnconfigure(0, weight=1)

        btn_frame = ttk.Frame(top_frame)
        btn_frame.grid(row=0, column=0, sticky="e", padx=5)
        self.start_check_button = ttk.Button(btn_frame, text="📦 Start Check", command=self.fetch_object_usages)
        self.start_check_button.grid(row=0, column=0, padx=5)
        self.stop_check_button = ttk.Button(btn_frame, text="Stop Check", command=self.stop_fetch, state="disabled")
        self.stop_check_button.grid(row=0, column=1, padx=5)

        # Table frame
        table_frame = ttk.Frame(frm)
        table_frame.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=5, pady=5)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.columns = ("Object Name", "Usage", "Type", "Folder", "Last Modified")
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, stretch=True)
            self.tree.heading(col, command=lambda c=col: self.on_column_click(c))

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        # Bindings
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Button-1>", self.on_single_click)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Control-z>", self.undo)
        self.tree.bind("<Control-y>", self.redo)
        self.tree.bind("<Control-a>", self.select_all)
        self.tree.bind("<Delete>", self.delete_selected)
        self.tree.bind("<BackSpace>", self.delete_selected)
        self.tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='Object Name'))
        self.tree.bind("<Control-c>", self.copy_selected)
        self.tree.bind("<Tab>", self.on_tab_press)
        self.tree.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, direction="backward"))

        # Configure tags for multiple usages
        self.tree.tag_configure('multiple_usages', background='#FFFFCC')  # Yellow for multiple usages

        # Log frame
        ttk.Label(frm, text="Searching Log:").grid(row=2, column=0, sticky="nw")
        self.log_box = tk.Text(frm, height=10, state='disabled')
        self.log_box.grid(row=3, column=0, columnspan=4, sticky='ew', padx=5)
        self.log_box.tag_configure('success', foreground='green')
        self.log_box.tag_configure('usage_name', foreground='blue')
        self.log_box.tag_configure('error', foreground='red')
        self.log_box.tag_configure('info', foreground='black')

        # Export frame
        export_frame = ttk.Frame(frm)
        export_frame.grid(row=4, column=0, columnspan=4, sticky="ew", padx=5, pady=5)
        ttk.Button(export_frame, text="📤 Export to Excel", command=self.export_to_excel).pack(side="left", anchor="w")
        self.status = tk.Label(export_frame, text="", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="right", anchor="e", fill="x", expand=True)

        frm.grid_rowconfigure(1, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        # Context menu for paste
        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)
        self.ensure_empty_row()  # Initialize with one empty row
    def ensure_empty_row(self):
        """Ensure there is always one empty row at the end of the Treeview."""
        columns = self.tree['columns']
        children = self.tree.get_children()
        if children:
            last_row_values = self.tree.item(children[-1])['values']
            last_row_empty = all(val == '' or val == 'None' for val in last_row_values)
            if last_row_empty:
                return
        empty_row = [''] * len(columns)
        self.tree.insert('', 'end', values=empty_row)

    def on_tab_press(self, event, current_row_id=None, current_col_num=None, direction="forward"):
        """Handle Tab (forward) and Shift+Tab (backward) key presses to move to the next or previous cell."""
        if current_row_id is None or current_col_num is None:
            selected = self.tree.selection()
            if not selected:
                return 'break'
            current_row_id = selected[0]
            current_col_num = 0 if direction == "forward" else len(self.tree['columns']) - 1
        else:
            if self.current_entry:
                self.save_edit(self.current_entry, current_row_id, self.tree['columns'][current_col_num])
        
        columns = self.tree['columns']
        children = self.tree.get_children()
        current_row_index = children.index(current_row_id)
        
        if direction == "forward":
            if current_col_num < len(columns) - 1:
                next_col_num = current_col_num + 1
                next_row_id = current_row_id
            else:
                if current_row_index < len(children) - 1:
                    next_row_id = children[current_row_index + 1]
                    next_col_num = 0
                else:
                    return 'break'
        else:
            if current_col_num > 0:
                next_col_num = current_col_num - 1
                next_row_id = current_row_id
            else:
                if current_row_index > 0:
                    next_row_id = children[current_row_index - 1]
                    next_col_num = len(columns) - 1
                else:
                    return 'break'
        
        next_col_name = columns[next_col_num]
        self.edit_cell(next_row_id, next_col_name, next_col_num)
        return 'break'

    def edit_cell(self, row_id, column_name, column_num):
        """Open an Entry widget for editing a specific cell."""
        if self.current_entry:
            self.current_entry.destroy()
        x, y, width, height = self.tree.bbox(row_id, f"#{column_num + 1}")
        self.current_entry = ttk.Entry(self.parent)
        self.current_entry.place(x=x, y=y, width=width, height=height)
        current_value = self.tree.item(row_id)['values'][column_num]
        self.current_entry.insert(0, current_value if current_value != 'None' else '')
        self.current_entry.select_range(0, tk.END)
        self.current_entry.focus()
        self.current_entry.bind("<Return>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<FocusOut>", lambda e: self.save_edit(self.current_entry, row_id, column_name))
        self.current_entry.bind("<Escape>", lambda e: self.current_entry.destroy())
        self.current_entry.bind("<Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="forward"))
        self.current_entry.bind("<Shift-Tab>", lambda e: self.on_tab_press(e, row_id, column_num, direction="backward"))

    def copy_selected(self, event=None):
        """Copy selected Treeview rows to the clipboard as tab-separated values."""
        selected = self.tree.selection()
        if not selected:
            self.log("No rows selected to copy.", tag='info')
            return 'break'

        rows_data = []
        for row_id in selected:
            values = self.tree.item(row_id)['values']
            if any(val and val != 'None' for val in values):
                row_values = [str(val) if val != 'None' else '' for val in values]
                rows_data.append('\t'.join(row_values))

        if rows_data:
            clipboard_text = '\n'.join(rows_data)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(clipboard_text)
            self.parent.update()
            self.log(f"Copied {len(rows_data)} row(s) to clipboard.", tag='info')
        else:
            self.log("No non-empty rows selected to copy.", tag='info')

        return 'break'
    def on_column_click(self, col_name):
        col_index = self.columns.index(col_name)
        values = [self.tree.item(item)['values'][col_index] for item in self.tree.get_children()]
        text = "\n".join(str(v) for v in values if v and v != "None")
        if text:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(text)
            self.parent.update()
            self.log(f"Copied column '{col_name}' to clipboard.", tag='info')
    def log(self, msg, tag='info'):
        self.parent.after(0, lambda: self._log(msg, tag))

    def _log(self, msg, tag):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n', tag)
        self.log_box.see('end')
        self.log_box.config(state='disabled')

    def show_context_menu(self, event):
        region = self.tree.identify("region", event.x, event.y)
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if region == "cell" and row_id and col_id:
            col_index = int(col_id[1:]) - 1
            column_name = self.columns[col_index]
            self.tree.selection_set(row_id)
            self.paste_menu.delete(0, "end")
            self.paste_menu.add_command(label=f"📋 Copy {column_name}", command=lambda c=column_name: self.copy_field(c))
            self.paste_menu.tk_popup(event.x_root, event.y_root)

    def on_single_click(self, event):
        if self.active_popup:
            self.active_popup.destroy()
            self.active_popup = None

        row_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not row_id or column != '#2' or row_id not in self.usage_names_map or len(self.usage_names_map[row_id]) <= 1:
            return
        self.show_usage_names_popup(row_id, event.x, event.y)
        return 'break'

    def show_usage_names_popup(self, row_id, x, y):
        listbox = tk.Listbox(self.parent, font=('Montserrat', 10), height=min(len(self.usage_names_map[row_id]), 5))
        listbox_width = 500
        listbox_height = min(len(self.usage_names_map[row_id]), 5) * 20
        self.parent.update()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        x_new = max(0, min(x, parent_width - listbox_width))
        y_new = max(0, min(y, parent_height - listbox_height))
        listbox.place(x=x_new, y=y_new, width=listbox_width)
        listbox.focus_set()
        self.active_popup = listbox

        for usage_name, usage_type, folder, last_modified in self.usage_names_map[row_id]:
            listbox.insert(tk.END, f"{usage_name} (Type: {usage_type}, Folder: {folder})")

        context_menu = tk.Menu(listbox, tearoff=0)
        context_menu.add_command(label="Copy to Clipboard", command=lambda: self.copy_usage_names(row_id))
        listbox.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))
        listbox.bind("<FocusOut>", lambda e: listbox.destroy())
        listbox.bind("<Escape>", lambda e: listbox.destroy())

    def copy_usage_names(self, row_id):
        usage_names = [f"{usage_name} (Type: {usage_type}, Folder: {folder})" for usage_name, usage_type, folder, last_modified in self.usage_names_map[row_id]]
        self.parent.clipboard_clear()
        self.parent.clipboard_append("\n".join(usage_names))
        self.log("Copied multiple usage names to clipboard.", tag='info')

    def fetch_object_usages(self):
        client_id = self.client_var.get().strip()
        userid = self.entries['USERID'].get().strip()
        password = self.entries['PASSWORD'].get().strip()
        env = self.env_var.get().strip()
        object_names = []
        for row_id in self.tree.get_children():
            values = self.tree.item(row_id)["values"]
            obj_name = values[0].strip()
            if obj_name and obj_name != "None":
                object_names.append((row_id, obj_name))

        if not object_names:
            messagebox.showinfo("Input Missing", "Please enter at least one object name in the Treeview.")
            return
        if not userid or not password:
            self.log("Please enter both User ID and Password.", tag='error')
            return

        self.auth = base64.b64encode(f"{userid}:{password}".encode()).decode()
        self.stop_flag = False
        self.start_check_button.config(state="disabled")
        self.stop_check_button.config(state="normal")
        threading.Thread(target=self.execute_batch_fetch, args=(object_names, env, client_id), daemon=True).start()

    def stop_fetch(self):
        self.stop_flag = True
        self.log("Stopping fetch process...", tag='info')
        self.start_check_button.config(state="normal")
        self.stop_check_button.config(state="disabled")

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        retry=retry_if_exception_type((ConnectionError, HTTPError)),
        before_sleep=lambda retry_state: logging.info(f"Retrying {retry_state.fn.__name__} ({retry_state.attempt_number}/3) after {retry_state.next_action.sleep}s")
    )
    def fetch_single_object(self, obj_name, env, client_id):
        if self.stop_flag:
            raise RuntimeError("Fetch stopped by user")
        api_url = f"https://rb-{env}-api.bosch.com"
        with self.connection_lock:
            automic.connection(url=api_url, auth=self.auth, noproxy=True, sslverify=False, timeout=self.TIMEOUT)
        try:
            result = automic.usageObject(client_id=int(client_id), object_name=obj_name)
            if result.status != 200:
                self.logger.warning(f"Object {obj_name} not found, status: {result.status}")
                return obj_name, []
            refs = [(ref["name"], ref["type"], ref["folderpath"], ref["lastmodified"][:10])
                    for ref in result.response.get("references", [])]
            return obj_name, refs
        except Exception as e:
            self.logger.error(f"Error fetching {obj_name}: {e}")
            return obj_name, []

    # def get_last_execution(self, client_id, obj_name):
    #     try:
    #         re = automic.listExecutions(client_id=int(client_id), query=f"name={obj_name}&include_deactivated=true&max_results=1")
    #         o = re.response.get('data', [])
    #         if o:
    #             raw_time = o[0]["start_time"]
    #             dt = datetime.strptime(raw_time, "%Y-%m-%dT%H:%M:%SZ")
    #             return dt.strftime("%Y-%m-%d %H:%M:%S")
    #         return "N/A"
    #     except Exception as e:
    #         self.logger.error(f"Error fetching last execution for {obj_name}: {e}")
    #         return "Error"

    def execute_batch_fetch(self, object_rows, env, client_id):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.usage_names_map.clear()
        total_refs_found = 0
        failed_objects = []

        try:
            with ThreadPoolExecutor(max_workers=self.MAX_WORKERS) as executor:
                future_to_obj = {executor.submit(self.fetch_single_object, obj_name, env, client_id): (row_id, obj_name) for row_id, obj_name in object_rows}
                results = {obj_name: None for _, obj_name in object_rows}

                for future in as_completed(future_to_obj):
                    if self.stop_flag:
                        self.logger.info("Batch fetch interrupted by user.")
                        executor._threads.clear()
                        break
                    row_id, obj_name = future_to_obj[future]
                    try:
                        results[obj_name] = (row_id, future.result())
                        # Update Treeview immediately
                        self.update_row_real_time(results[obj_name], total_refs_found, failed_objects)
                    except Exception as e:
                        self.logger.error(f"Future error for {obj_name}: {e}")
                        results[obj_name] = (row_id, (obj_name, [], None))
                        failed_objects.append(obj_name)
                        # Update Treeview for failed object
                        self.update_row_real_time(results[obj_name], total_refs_found, failed_objects)

                # Final summary
                self.parent.after(0, lambda: self.log(f"Fetched {total_refs_found} references for {len(object_rows)} objects", tag='success'))
                if failed_objects:
                    self.parent.after(0, lambda: self.log(f"Failed to fetch {len(failed_objects)} objects: {', '.join(failed_objects)}", tag='error'))

        except Exception as e:
            self.logger.error(f"Fetch failed: {e}")
            self.parent.after(0, lambda: messagebox.showerror("Error", f"Fetch failed:\n{str(e)}"))
        finally:
            self.parent.after(0, lambda: self.start_check_button.config(state="normal"))
            self.parent.after(0, lambda: self.stop_check_button.config(state="disabled"))
            if self.stop_flag:
                self.parent.after(0, lambda: self.log("Fetch cancelled.", tag='info'))

    def update_row_real_time(self, result, total_refs_found, failed_objects):
        row_id, (obj_name, refs) = result

        def update_row():
            nonlocal total_refs_found
            if not self.tree.exists(row_id):
                return  # Skip if row was deleted during fetch
            if not refs:
                row = (obj_name, "None", "None", "None", "None")
                self.tree.item(row_id, values=row, tags=())
                self.all_rows = [(r_id, self.tree.item(r_id)["values"], self.tree.item(r_id)["tags"])
                                 for r_id in self.tree.get_children()]
                self.log(f"No references found for {obj_name}", tag='error')
            else:
                first_ref = refs[0]
                row = (obj_name, first_ref[0], first_ref[1], first_ref[2], first_ref[3])
                tags = []
                if len(refs) > 1:
                    tags.append('multiple_usages')
                self.tree.item(row_id, values=row, tags=tags)
                self.usage_names_map[row_id] = refs
                self.all_rows = [(r_id, self.tree.item(r_id)["values"], self.tree.item(r_id)["tags"])
                                 for r_id in self.tree.get_children()]
                total_refs_found += len(refs)
                for ref in refs:
                    self.log(f"Found usage for {obj_name}: {ref[0]} (Type: {ref[1]})", tag='usage_name')
            self.tree.update()

        self.parent.after(0, update_row)

    def copy_field(self, field):
        row_id = self.tree.selection()
        if not row_id:
            return
        values = self.tree.item(row_id[0])["values"]
        col_index = self.columns.index(field)
        value = values[col_index]
        if value:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(value)
            self.log(f"Copied {field} to clipboard ✔", tag='info')

    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        columns = self.tree['columns']
        if start_column is None:
            start_column = 'Object Name'
        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return

        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return
            existing_rows = list(self.tree.get_children())
            max_columns = len(columns) - start_col_index
            existing_names = {self.tree.item(row_id)["values"][0].strip() for row_id in existing_rows}

            for i, line in enumerate(lines):
                values = line.split('\t') if '\t' in line else [line]
                values = values[:max_columns]
                while len(values) < max_columns:
                    values.append('')
                obj_name = values[0].strip() if values else ''
                if obj_name and obj_name not in existing_names:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(values):
                        new_values[start_col_index + j] = value.strip()
                    self.tree.insert('', 'end', values=new_values)
                    existing_names.add(obj_name)
                elif obj_name in existing_names:
                    # Update existing row if it matches
                    for row_id in existing_rows:
                        if self.tree.item(row_id)["values"][0].strip() == obj_name:
                            current_values = list(self.tree.item(row_id)['values'])
                            for j, value in enumerate(values):
                                current_values[start_col_index + j] = value.strip()
                            self.tree.item(row_id, values=current_values)
                            break
            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.", tag='info')
            self.ensure_empty_row()  # Ensure an empty row after pasting
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
            self.undo_stack.pop()

    def on_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        column = self.tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.columns[column_num]
        self.edit_cell(row_id, column_name, column_num)

    def save_edit(self, entry, row_id, column_name):
        new_value = entry.get().strip()
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.tree.set(row_id, column_name, new_value)
        entry.destroy()
        self.ensure_empty_row()  # Ensure an empty row after editing
    def save_state(self):
        return [(child, self.tree.item(child)['values'], self.tree.item(child)['tags'], self.usage_names_map.get(child, []))
                for child in self.tree.get_children()]

    def restore_state(self, state):
        self.tree.delete(*self.tree.get_children())
        self.usage_names_map.clear()
        for child, values, tags, usages in state:
            self.tree.insert('', 'end', child, values=values, tags=tags)
            if usages:
                self.usage_names_map[child] = usages
        self.ensure_empty_row()  # Ensure an empty row after restoring state

    def undo(self, event):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def select_all(self, event):
        self.tree.selection_set(self.tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.usage_names_map.pop(item, None)
                self.tree.delete(item)
            self.ensure_empty_row()  # Ensure an empty row after restoring state
        return 'break'

    def export_to_excel(self):
        try:
            rows = []
            for row_id in self.tree.get_children():
                values = self.tree.item(row_id)["values"]
                if row_id in self.usage_names_map and len(self.usage_names_map[row_id]) > 1:
                    for usage_name, usage_type, folder, last_modified in self.usage_names_map[row_id]:
                        rows.append([values[0], usage_name, usage_type, folder, last_modified])
                else:
                    rows.append(values)
            if not rows:
                messagebox.showinfo("No Data", "There is no data to export.")
                return
            df = pd.DataFrame(rows, columns=self.columns)
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save as")
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Export Successful", f"Data exported to:\n{file_path}")
                self.log(f"Exported to {file_path}", tag='info')
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not export to Excel:\n{str(e)}")
logging.basicConfig(level=logging.INFO, filename="childviewer.log", format="%(asctime)s - %(levelname)s - %(message)s")

class ChildViewer:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.parent.grid_rowconfigure(1, weight=3)
        self.parent.grid_rowconfigure(2, weight=1)
        self.parent.grid_columnconfigure(0, weight=1)

        top_frame = ttk.Frame(self.parent)
        top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        top_frame.grid_columnconfigure((0, 1), weight=1)

        ttk.Label(top_frame, text="Multiple Object Names (one per line):").grid(row=0, column=0, columnspan=2, sticky="w", pady=(10, 0), padx=5)
        self.batch_input = tk.Text(top_frame, height=4, width=80, undo=True)
        self.batch_input.bind("<Control-z>", lambda e: self.batch_input.edit_undo())
        self.batch_input.bind("<Control-y>", lambda e: self.batch_input.edit_redo())
        self.batch_input.grid(row=1, column=0, columnspan=1, pady=(0, 10), sticky="nsew", padx=5)

        self.batch_fetch_button = tk.Button(top_frame, text="📦 Batch Fetch", command=self.batch_fetch)
        self.batch_fetch_button.grid(row=1, column=1, sticky="e", padx=10)
        self.cancel_button = ttk.Button(top_frame, text="Cancel Fetch", command=self.cancel_batch_fetch)
        self.cancel_button.grid(row=0, column=1, sticky="e", padx=10)
        self.cancel_button.grid_remove()

        self.spinner = ttk.Progressbar(top_frame, mode='indeterminate')
        self.spinner.grid(row=1, column=1, sticky="e", padx=10)
        self.spinner.grid_remove()

        table_frame = ttk.Frame(self.parent)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=7)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.columns = ("Object Name", "Child", "Type", "Active", "Last Execution")
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        self.tree.bind("<Button-1>", self.on_column_click)

        for col in self.columns:
            if col == "Type":
                self.tree.heading(col, text=f"{col} 🔽", command=self.show_filter_menu)
            else:
                self.tree.heading(col, text=col)
            self.tree.column(col, width=145, stretch=True)

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        details_frame = ttk.Frame(self.parent)
        details_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=6)
        details_frame.grid_columnconfigure(1, weight=1)

        self.detail_labels = {}
        ttk.Label(details_frame, text="Object Name:").grid(row=0, column=0, sticky="e", padx=5)
        self.object_name_label = tk.Label(details_frame, text="", anchor="w", bg="white", relief="sunken")
        self.object_name_label.grid(row=0, column=1, padx=5, pady=1, sticky="ew")
        self.detail_labels["Object Name"] = self.object_name_label

        for i, field in enumerate(self.columns[1:]):
            ttk.Label(details_frame, text=f"{field}:").grid(row=i+1, column=0, sticky="e", padx=5)
            lbl = tk.Label(details_frame, text="", anchor="w", bg="white", relief="sunken")
            lbl.grid(row=i+1, column=1, padx=5, pady=1, sticky="ew")
            self.detail_labels[field] = lbl
            ttk.Button(details_frame, text=f"📋 Copy {field}", command=lambda f=field: self.copy_field(f)).grid(row=i+1, column=2, padx=5, pady=1)

        export_frame = ttk.Frame(self.parent)
        export_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 5))

        ttk.Button(export_frame, text="📤 Export to Excel", command=self.export_to_excel).pack(side="left", anchor="w")
        self.status = tk.Label(export_frame, text="", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="right", anchor="e", fill="x", expand=True)

        self.color_map = {}
        self.setup_treeview_context_menu()
        self.palette_a = ["#b0b1aa", "#a4ddc7", "#dbe0c5", "#b5b095"]
        self.palette_b = ["#ffffea", "#fff2cc", "#fce5cd", "#ead1dc", "#d0e0e3"]
        self.color_index_a = 0
        self.color_index_b = 0
        self.assign_counter = 0
        self.type_filter_var = tk.StringVar(value="All")  # Filter variable for Type column
        self.all_rows = []  # Store all rows for filtering
        self.type_filter_menu = None  # Store filter menu for Type column

    def setup_treeview_context_menu(self):
        self.menu = tk.Menu(self.parent, tearoff=0)
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        region = self.tree.identify("region", event.x, event.y)
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if region == "cell" and row_id and col_id:
            col_index = int(col_id[1:]) - 1
            column_name = self.columns[col_index]
            self.tree.selection_set(row_id)
            self.selected_data = dict(zip(self.columns, self.tree.item(row_id)["values"]))

            self.menu.delete(0, "end")
            self.menu.add_command(label=f"📋 Copy {column_name}", command=lambda c=column_name: self.copy_column_value(c))
            self.menu.tk_popup(event.x_root, event.y_root)
    def show_filter_menu(self):
        # Only show filter menu if there are rows
        if not self.tree.get_children():
            self.status.config(text="No data to filter", fg="red")
            return

        # Ensure Treeview is updated before querying bbox
        self.tree.update()
        # Get unique values for the Type column
        type_index = self.columns.index("Type")
        values = set(self.tree.item(row)["values"][type_index] for row in self.tree.get_children())
        values = sorted(values, key=lambda x: str(x).lower())
        values = ["All"] + list(values)

        # Create or update filter menu
        if self.type_filter_menu is None:
            self.type_filter_menu = tk.Menu(self.parent, tearoff=0)
        else:
            self.type_filter_menu.delete(0, "end")

        for value in values:
            self.type_filter_menu.add_radiobutton(
                label=str(value),
                value=str(value),
                variable=self.type_filter_var,
                command=self.apply_filter
            )

        # Calculate position for the filter menu
        try:
            # Try to get the bounding box of the first visible row
            first_visible_item = self.tree.get_children()[0]
            header_bbox = self.tree.bbox(first_visible_item)
            if not header_bbox:  # If bbox returns empty, use fallback
                header_bbox = (0, 0, 0, 30)  # Estimate header height
            x = self.tree.winfo_rootx() + sum(self.tree.column(col)["width"] for col in self.columns[:type_index])
            y = self.tree.winfo_rooty() + header_bbox[3]
            self.type_filter_menu.tk_popup(x, y)
        except Exception as e:
            logging.error(f"Error showing filter menu: {e}")
            self.status.config(text="Error showing filter menu", fg="red")
    def apply_filter(self):
        # Store current rows if not already stored
        if not self.all_rows:
            self.all_rows = [self.tree.item(row)["values"] for row in self.tree.get_children()]

        # Clear Treeview
        self.tree.delete(*self.tree.get_children())

        # Apply filter on Type column
        type_index = self.columns.index("Type")
        selected_type = self.type_filter_var.get()
        for row in self.all_rows:
            if selected_type == "All" or str(row[type_index]) == selected_type:
                obj_name = row[0]
                self.tree.insert("", "end", values=row, tags=(obj_name,))
                self.tree.tag_configure(obj_name, background=self.color_map.get(obj_name, "#ffffff"))

        self.status.config(text="Type filter applied ✔", fg="green")
        self.tree.update()
    def on_column_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.tree.identify_column(event.x)
            col_index = int(col_id.replace("#", "")) - 1
            col_name = self.columns[col_index]
            if col_name != "Type":  # Only copy for non-Type columns
                values = [self.tree.item(item)["values"][col_index] for item in self.tree.get_children()]
                text = "\n".join(str(v) for v in values)
                self.parent.clipboard_clear()
                self.parent.clipboard_append(text)
                self.parent.update()
                self.status.config(text=f"Copied column '{col_name}' to clipboard ✔", fg="green")

    def copy_column_value(self, column):
        value = self.selected_data.get(column, "")
        if value:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(value)
            self.parent.update()
            self.status.config(text=f"Copied {column} to clipboard ✔")

    def start_batch_fetch_spinner(self):
        self.batch_fetch_button.grid_remove()
        self.spinner.grid()
        self.spinner.start(10)

    def stop_batch_fetch_spinner(self):
        self.spinner.stop()
        self.spinner.grid_remove()
        self.batch_fetch_button.grid()

    def batch_fetch(self):
        client_id = self.client_var.get()
        userid = self.entries['USERID'].get()
        password = self.entries['PASSWORD'].get()
        env = self.env_var.get()
        object_names = [name.strip() for name in self.batch_input.get("1.0", tk.END).strip().splitlines() if name.strip()]

        if not object_names:
            messagebox.showinfo("Input Missing", "Please enter at least one object name.")
            return

        self.cancel_batch = False
        logging.info(f"Starting batch fetch for {len(object_names)} objects")

        def fetch_single_object(obj_name):
            logging.info(f"Fetching object: {obj_name}")
            try:
                if self.cancel_batch:
                    logging.warning(f"Fetch cancelled for {obj_name}")
                    return obj_name, None
                refs = []
                result = automic.getObjects(client_id=int(client_id), object_name=obj_name)
                if result.status != 200:
                    logging.warning(f"Object {obj_name} not found, status: {result.status}")
                    return obj_name, ("Not found", [], None, None)
                for jobpl in result.response["data"]["jobp"]["workflow_definitions"]:
                    if jobpl["object_type"] not in ("<START>", "<END>"):
                        if jobpl["object_name"] in ("CI_XXXX_XX99001_EMAIL_NOT_OK", "XX_XXXX_SMT_GET_VALUES_ERP"):
                            jobpl["object_type"] = "Common"
                        refs.append(jobpl)
                last_exec = self.get_last_execution(client_id, obj_name)
                color = self.get_object_color(obj_name)
                logging.info(f"Fetched {obj_name} with {len(refs)} children")
                return obj_name, (obj_name, refs, last_exec, color)
            except Exception as e:
                logging.error(f"Error fetching {obj_name}: {e}")
                self.status.config(text=f"Error fetching {obj_name}: {e}")
                return obj_name, (obj_name, [], None, None)

        def fetch_objects():
            try:
                auth = base64.b64encode(f"{userid}:{password}".encode()).decode()
                url = f"https://rb-{env}-api.bosch.com"
                automic.connection(url=url, auth=auth, noproxy=True, sslverify=False, cert="/path/to/certfile", timeout=60)

                # Clear Treeview
                self.parent.after(0, lambda: self.tree.delete(*self.tree.get_children()))
                self.all_rows = []
                self.type_filter_var.set("All")
                total_refs_found = 0
                inserted_refs = set()  # Track inserted references to avoid duplicates
                results = {name: None for name in object_names}  # Preserve input order

                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = {executor.submit(fetch_single_object, obj_name): obj_name for obj_name in object_names}
                    for future in as_completed(futures):
                        if self.cancel_batch:
                            logging.info("Batch fetch cancelled")
                            break
                        obj_name = futures[future]
                        try:
                            result = future.result()
                            if result[1] is None:  # Cancelled
                                continue
                            results[obj_name] = result[1]
                        except Exception as e:
                            logging.error(f"Future error for {obj_name}: {e}")
                            results[obj_name] = (obj_name, [], None, None)

                # Insert results in input order
                def insert_rows():
                    nonlocal total_refs_found
                    for obj_name in object_names:  # Follow original order
                        result = results.get(obj_name)
                        if result is None:
                            logging.warning(f"No result for {obj_name}, possibly cancelled")
                            continue
                        obj_name, refs, last_exec, color = result
                        if not refs:
                            row_key = f"{obj_name}_none"
                            if row_key not in inserted_refs:
                                row = (obj_name, "None", "None", "None", last_exec)
                                self.tree.insert("", "end", values=row, tags=(obj_name,))
                                self.all_rows.append(row)
                                inserted_refs.add(row_key)
                                logging.debug(f"Inserted {obj_name} with no children")
                        else:
                            for r in refs:
                                row_key = f"{obj_name}_{r['object_name']}"
                                if row_key not in inserted_refs:
                                    active = r.get("active", 0)  # Default to 0 if "active" is missing
                                    row = (obj_name, r["object_name"], r["object_type"], active, last_exec)
                                    self.tree.insert("", "end", values=row, tags=(obj_name,))
                                    self.all_rows.append(row)
                                    inserted_refs.add(row_key)
                                    total_refs_found += 1
                                    logging.debug(f"Inserted {obj_name} child {r['object_name']}")
                                else:
                                    logging.warning(f"Skipped duplicate child {r['object_name']} for {obj_name}")
                        self.tree.tag_configure(obj_name, background=color)
                    self.tree.update()  # Force Treeview refresh
                    self.status.config(text=f"Fetched {total_refs_found} children for {len(object_names)} objects")

                self.parent.after(0, insert_rows)

                if self.cancel_batch:
                    self.parent.after(0, lambda: self.status.config(text="Fetch cancelled."))

            except Exception as e:
                logging.error(f"Batch fetch failed: {e}")
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Batch fetch failed:\n{str(e)}"))
            finally:
                logging.info("Batch fetch completed")
                self.parent.after(0, self.stop_batch_fetch_spinner)
                self.parent.after(0, self.hide_cancel_button)

        self.start_batch_fetch_spinner()
        self.show_cancel_button()
        threading.Thread(target=fetch_objects, daemon=True).start()

    def show_cancel_button(self):
        self.cancel_button.grid()

    def hide_cancel_button(self):
        self.cancel_button.grid_remove()

    def cancel_batch_fetch(self):
        self.cancel_batch = True

    def get_last_execution(self, client_id, obj_name):
        try:
            re = automic.listExecutions(client_id=int(client_id), query=f"{obj_name}&max_results=1")
            o = re.response.get('data', [])
            if o:
                raw_time = o[0]["start_time"]
                dt = datetime.strptime(raw_time, "%Y-%m-%dT%H:%M:%SZ")
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            else:
                return "N/A"
        except Exception as e:
            logging.error(f"Error fetching last execution for {obj_name}: {e}")
            return "Error"

    def get_object_color(self, obj_name):
        if obj_name not in self.color_map:
            if self.assign_counter % 2 == 0:
                color = self.palette_a[self.color_index_a]
                self.color_index_a = (self.color_index_a + 1) % len(self.palette_a)
            else:
                color = self.palette_b[self.color_index_b]
                self.color_index_b = (self.color_index_b + 1) % len(self.palette_b)
            self.color_map[obj_name] = color
            self.assign_counter += 1
        return self.color_map[obj_name]

    def on_row_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])["values"]
        self.selected_data = dict(zip(self.columns, values))
        self.object_name_label.config(text=self.selected_data.get("Object Name", ""))
        for key in self.columns[1:]:
            self.detail_labels[key].config(text=self.selected_data.get(key, ""))

    def copy_field(self, field):
        value = self.selected_data.get(field, "")
        if value:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(value)
            self.parent.update()
            self.status.config(text=f"Copied {field} to clipboard ✔")

    def export_to_excel(self):
        try:
            rows = [self.tree.item(row)["values"] for row in self.tree.get_children()]
            if not rows:
                messagebox.showinfo("No Data", "There is no data to export.")
                return
            df = pd.DataFrame(rows, columns=self.columns)
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save as")
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Export Successful", f"Data exported to:\n{file_path}")
                self.status.config(text=f"Exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not export to Excel:\n{str(e)}")

class FindBulk:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.parent.grid_rowconfigure(1, weight=3)
        self.parent.grid_rowconfigure(2, weight=1)
        self.parent.grid_columnconfigure(0, weight=1)

        top_frame = ttk.Frame(self.parent)
        top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        top_frame.grid_columnconfigure((0, 1), weight=1)

        ttk.Label(top_frame, text="Multiple Object Names (one per line):").grid(row=0, column=0, columnspan=2, sticky="w", pady=(10, 0), padx=5)
        self.batch_input = tk.Text(top_frame, height=4, width=80, undo=True)
        self.batch_input.bind("<Control-z>", lambda e: self.batch_input.edit_undo())
        self.batch_input.bind("<Control-y>", lambda e: self.batch_input.edit_redo())
        self.batch_input.grid(row=1, column=0, columnspan=1, pady=(0, 10), sticky="nsew", padx=5)
        # Filter Dropdown and Value Input
        filter_frame = ttk.Frame(top_frame)
        filter_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=5, padx=5)
        ttk.Label(filter_frame, text="Search by:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.filter_var = tk.StringVar(value="object_name")
        filter_options = ["process","object_name","documentation"]  # Add more filter options as needed
        self.filter_dropdown = ttk.Combobox(filter_frame, textvariable=self.filter_var, values=filter_options, state="readonly", width=15)
        self.filter_dropdown.grid(row=0, column=1, sticky="w", padx=(0, 5))
        # self.filter_dropdown.bind("<<ComboboxSelected>>", self.toggle_filter_input)

        # # Text Entry for filter value
        # self.filter_value_entry = ttk.Entry(filter_frame, width=20)
        # self.filter_value_entry.grid(row=0, column=2, sticky="w")
        # self.filter_value_entry.config(state="disabled")  # Initially disabled

        # Dropdown for object_type filter values
        self.object_type_var = tk.StringVar(value="JOBS")
        self.object_type_dropdown = ttk.Combobox(filter_frame, textvariable=self.object_type_var, values=["JOBS", "JOBP"], state="readonly", width=20)
        self.object_type_dropdown.grid(row=0, column=2, sticky="w")
        self.object_type_dropdown.grid_remove()  # Initially hidden


        self.batch_fetch_button = tk.Button(top_frame, text="📦 Batch Fetch", command=self.batch_fetch)
        self.batch_fetch_button.grid(row=1, column=1, sticky="e", padx=10)
        self.cancel_button = ttk.Button(top_frame, text="Cancel Fetch", command=self.cancel_batch_fetch)
        self.cancel_button.grid(row=0, column=1, sticky="e", padx=10)
        self.cancel_button.grid_remove()

        self.spinner = ttk.Progressbar(top_frame, mode='indeterminate')
        self.spinner.grid(row=1, column=1, sticky="e", padx=10)
        self.spinner.grid_remove()

        table_frame = ttk.Frame(self.parent)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=7)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.columns = ("Search Name", "Objects", "Type", "Last Execution")
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        self.tree.bind("<Button-1>", self.on_column_click)

        for col in self.columns:
            if col == "Type":
                self.tree.heading(col, text=f"{col} 🔽", command=self.show_filter_menu)
            else:
                self.tree.heading(col, text=col)
            self.tree.column(col, width=145, stretch=True)


        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        details_frame = ttk.Frame(self.parent)
        details_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=6)
        details_frame.grid_columnconfigure(1, weight=1)

        self.detail_labels = {}
        ttk.Label(details_frame, text="Object Name:").grid(row=0, column=0, sticky="e", padx=5)
        self.object_name_label = tk.Label(details_frame, text="", anchor="w", bg="white", relief="sunken")
        self.object_name_label.grid(row=0, column=1, padx=5, pady=1, sticky="ew")
        self.detail_labels["Object Name"] = self.object_name_label

        for i, field in enumerate(self.columns[1:]):
            ttk.Label(details_frame, text=f"{field}:").grid(row=i+1, column=0, sticky="e", padx=5)
            lbl = tk.Label(details_frame, text="", anchor="w", bg="white", relief="sunken")
            lbl.grid(row=i+1, column=1, padx=5, pady=1, sticky="ew")
            self.detail_labels[field] = lbl
            ttk.Button(details_frame, text=f"📋 Copy {field}", command=lambda f=field: self.copy_field(f)).grid(row=i+1, column=2, padx=5, pady=1)

        export_frame = ttk.Frame(self.parent)
        export_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 5))

        ttk.Button(export_frame, text="📤 Export to Excel", command=self.export_to_excel).pack(side="left", anchor="w")
        self.status = tk.Label(export_frame, text="", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="right", anchor="e", fill="x", expand=True)

        self.color_map = {}
        self.setup_treeview_context_menu()
        self.palette_a = ["#b0b1aa", "#a4ddc7", "#dbe0c5", "#b5b095"]
        self.palette_b = ["#ffffea", "#fff2cc", "#fce5cd", "#ead1dc", "#d0e0e3"]
        self.color_index_a = 0
        self.color_index_b = 0
        self.assign_counter = 0
        self.type_filter_var = tk.StringVar(value="All")  # Filter variable for Type column
        self.all_rows = []  # Store all rows for filtering
        self.type_filter_menu = None  # Store filter menu for Type column

    # def toggle_filter_input(self, event=None):
    #         """Toggle between text entry and dropdown based on filter selection."""
    #         filter_type = self.filter_var.get()
    #         if filter_type == "object_types":
    #             self.filter_value_entry.grid_remove()
    #             self.object_type_dropdown.grid()
    #         else:
    #             self.object_type_dropdown.grid_remove()
    #             self.filter_value_entry.grid()
    #             if filter_type == "None":
    #                 self.filter_value_entry.config(state="disabled")
    #                 self.filter_value_entry.delete(0, tk.END)
    #             else:
    #                 self.filter_value_entry.config(state="normal")
    def setup_treeview_context_menu(self):
        self.menu = tk.Menu(self.parent, tearoff=0)
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        region = self.tree.identify("region", event.x, event.y)
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if region == "cell" and row_id and col_id:
            col_index = int(col_id[1:]) - 1
            column_name = self.columns[col_index]
            self.tree.selection_set(row_id)
            self.selected_data = dict(zip(self.columns, self.tree.item(row_id)["values"]))

            self.menu.delete(0, "end")
            self.menu.add_command(label=f"📋 Copy {column_name}", command=lambda c=column_name: self.copy_column_value(c))
            self.menu.tk_popup(event.x_root, event.y_root)

    def show_filter_menu(self):
        # Only show filter menu if there are rows
        if not self.tree.get_children():
            self.status.config(text="No data to filter", fg="red")
            return

        # Ensure Treeview is updated before querying bbox
        self.tree.update()
        # Get unique values for the Type column
        type_index = self.columns.index("Type")
        values = set(self.tree.item(row)["values"][type_index] for row in self.tree.get_children())
        values = sorted(values, key=lambda x: str(x).lower())
        values = ["All"] + list(values)

        # Create or update filter menu
        if self.type_filter_menu is None:
            self.type_filter_menu = tk.Menu(self.parent, tearoff=0)
        else:
            self.type_filter_menu.delete(0, "end")

        for value in values:
            self.type_filter_menu.add_radiobutton(
                label=str(value),
                value=str(value),
                variable=self.type_filter_var,
                command=self.apply_filter
            )

        # Calculate position for the filter menu
        try:
            # Try to get the bounding box of the first visible row
            first_visible_item = self.tree.get_children()[0]
            header_bbox = self.tree.bbox(first_visible_item)
            if not header_bbox:  # If bbox returns empty, use fallback
                header_bbox = (0, 0, 0, 30)  # Estimate header height
            x = self.tree.winfo_rootx() + sum(self.tree.column(col)["width"] for col in self.columns[:type_index])
            y = self.tree.winfo_rooty() + header_bbox[3]
            self.type_filter_menu.tk_popup(x, y)
        except Exception as e:
            logging.error(f"Error showing filter menu: {e}")
            self.status.config(text="Error showing filter menu", fg="red")
    def apply_filter(self):
        # Store current rows if not already stored
        if not self.all_rows:
            self.all_rows = [self.tree.item(row)["values"] for row in self.tree.get_children()]

        # Clear Treeview
        self.tree.delete(*self.tree.get_children())

        # Apply filter on Type column
        type_index = self.columns.index("Type")
        selected_type = self.type_filter_var.get()
        for row in self.all_rows:
            if selected_type == "All" or str(row[type_index]) == selected_type:
                obj_name = row[0]
                self.tree.insert("", "end", values=row, tags=(obj_name,))
                self.tree.tag_configure(obj_name, background=self.color_map.get(obj_name, "#ffffff"))

        self.status.config(text="Type filter applied ✔", fg="green")
        self.tree.update()
    def on_column_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.tree.identify_column(event.x)
            col_index = int(col_id.replace("#", "")) - 1
            col_name = self.columns[col_index]
            if col_name != "Type":  # Only copy for non-Type columns
                values = [self.tree.item(item)["values"][col_index] for item in self.tree.get_children()]
                text = "\n".join(str(v) for v in values)
                self.parent.clipboard_clear()
                self.parent.clipboard_append(text)
                self.parent.update()
                self.status.config(text=f"Copied column '{col_name}' to clipboard ✔", fg="green")


    def copy_column_value(self, column):
        value = self.selected_data.get(column, "")
        if value:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(value)
            self.parent.update()
            self.status.config(text=f"Copied {column} to clipboard ✔")

    def start_batch_fetch_spinner(self):
        self.batch_fetch_button.grid_remove()
        self.spinner.grid()
        self.spinner.start(10)

    def stop_batch_fetch_spinner(self):
        self.spinner.stop()
        self.spinner.grid_remove()
        self.batch_fetch_button.grid()

    def batch_fetch(self):
        client_id = self.client_var.get()
        userid = self.entries['USERID'].get()
        password = self.entries['PASSWORD'].get()
        env = self.env_var.get()
        object_names = [name.strip() for name in self.batch_input.get("1.0", tk.END).strip().splitlines() if name.strip()]
        filter_type = self.filter_var.get()
        # filter_value = self.object_type_var.get() if filter_type == "object_types" else self.filter_value_entry.get().strip()
        
        if not object_names:
            messagebox.showinfo("Input Missing", "Please enter at least one object name.")
            return

        self.cancel_batch = False
        logging.info(f"Starting batch fetch for {len(object_names)} objects")

        def fetch_single_object(obj_name):
            logging.info(f"Fetching object: {obj_name}")
            try:
                if self.cancel_batch:
                    logging.warning(f"Fetch cancelled for {obj_name}")
                    return obj_name, None
                refs = []
                if filter_type == "object_name":
                    query_filter = {"object_name": obj_name}
                else:
                    query_filter = {"query": obj_name}
                body = {
                    "filters": [
                        {
                            **query_filter,
                            "filter_identifier": filter_type
                        }
                    ],
                    "max_results": 9999
                }
                result = automic.findObjects(client_id=int(client_id), body=body)
                if result.status != 200:
                    logging.warning(f"Object {obj_name} not found, status: {result.status}")
                    return obj_name, ("Not found", [], None, None)
                refs = result.response["data"]
                last_exec = self.get_last_execution(client_id, obj_name)
                color = self.get_object_color(obj_name)
                logging.info(f"Fetched {obj_name} with {len(refs)} references")
                return obj_name, (obj_name, refs, last_exec, color)
            except Exception as e:
                logging.error(f"Error fetching {obj_name}: {e}")
                self.status.config(text=f"Error fetching {obj_name}: {e}")
                return obj_name, (obj_name, [], None, None)

        def fetch_objects():
            try:
                auth = base64.b64encode(f"{userid}:{password}".encode()).decode()
                url = f"https://rb-{env}-api.bosch.com"
                automic.connection(url=url, auth=auth, noproxy=True, sslverify=False, cert="/path/to/certfile", timeout=60)

                # Clear Treeview
                self.parent.after(0, lambda: self.tree.delete(*self.tree.get_children()))
                self.all_rows = []
                self.type_filter_var.set("All")
                total_refs_found = 0
                inserted_refs = set()  # Track inserted references to avoid duplicates
                results = {name: None for name in object_names}  # Preserve input order

                with ThreadPoolExecutor(max_workers=5) as executor:
                    futures = {executor.submit(fetch_single_object, obj_name): obj_name for obj_name in object_names}
                    for future in as_completed(futures):
                        if self.cancel_batch:
                            logging.info("Batch fetch cancelled")
                            break
                        obj_name = futures[future]
                        try:
                            result = future.result()
                            if result[1] is None:  # Cancelled
                                continue
                            results[obj_name] = result[1]
                        except Exception as e:
                            logging.error(f"Future error for {obj_name}: {e}")
                            results[obj_name] = (obj_name, [], None, None)

                # Insert results in input order
                def insert_rows():
                    nonlocal total_refs_found
                    for obj_name in object_names:  # Follow original order
                        result = results.get(obj_name)
                        if result is None:
                            logging.warning(f"No result for {obj_name}, possibly cancelled")
                            continue
                        obj_name, refs, last_exec, color = result
                        if not refs:
                            row_key = f"{obj_name}_none"
                            if row_key not in inserted_refs:
                                self.tree.insert("", "end", values=(obj_name, "None", "None", last_exec), tags=(obj_name,))
                                inserted_refs.add(row_key)
                                logging.debug(f"Inserted {obj_name} with no refs")
                        else:
                            for r in refs:
                                row_key = f"{obj_name}_{r['name']}"
                                if row_key not in inserted_refs:
                                    self.tree.insert("", "end", values=(obj_name, r["name"], r["type"], last_exec), tags=(obj_name,))
                                    inserted_refs.add(row_key)
                                    total_refs_found += 1
                                    logging.debug(f"Inserted {obj_name} ref {r['name']}")
                                else:
                                    logging.warning(f"Skipped duplicate ref {r['name']} for {obj_name}")
                        self.tree.tag_configure(obj_name, background=color)
                    self.tree.update()  # Force Treeview refresh
                    self.status.config(text=f"Fetched {total_refs_found} references for {len(object_names)} objects")

                self.parent.after(0, insert_rows)

                if self.cancel_batch:
                    self.parent.after(0, lambda: self.status.config(text="Fetch cancelled."))

            except Exception as e:
                logging.error(f"Batch fetch failed: {e}")
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Batch fetch failed:\n{str(e)}"))
            finally:
                logging.info("Batch fetch completed")
                self.parent.after(0, self.stop_batch_fetch_spinner)
                self.parent.after(0, self.hide_cancel_button)

        self.start_batch_fetch_spinner()
        self.show_cancel_button()
        threading.Thread(target=fetch_objects, daemon=True).start()

    def show_cancel_button(self):
        self.cancel_button.grid()

    def hide_cancel_button(self):
        self.cancel_button.grid_remove()

    def cancel_batch_fetch(self):
        self.cancel_batch = True

    def get_last_execution(self, client_id, obj_name):
        try:
            re = automic.listExecutions(client_id=int(client_id), query=f"{obj_name}&max_results=1")
            o = re.response.get('data', [])
            if o:
                raw_time = o[0]["start_time"]
                dt = datetime.strptime(raw_time, "%Y-%m-%dT%H:%M:%SZ")
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            else:
                return "N/A"
        except Exception as e:
            return "Error"

    def get_object_color(self, obj_name):
        if obj_name not in self.color_map:
            if self.assign_counter % 2 == 0:
                color = self.palette_a[self.color_index_a]
                self.color_index_a = (self.color_index_a + 1) % len(self.palette_a)
            else:
                color = self.palette_b[self.color_index_b]
                self.color_index_b = (self.color_index_b + 1) % len(self.palette_b)
            self.color_map[obj_name] = color
            self.assign_counter += 1
        return self.color_map[obj_name]

    def on_row_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])["values"]
        self.selected_data = dict(zip(self.columns, values))
        self.object_name_label.config(text=self.selected_data.get("Object Name", ""))
        for key in self.columns[1:]:
            self.detail_labels[key].config(text=self.selected_data.get(key, ""))

    def copy_field(self, field):
        value = self.selected_data.get(field, "")
        if value:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(value)
            self.parent.update()
            self.status.config(text=f"Copied {field} to clipboard ✔")

    def export_to_excel(self):
        try:
            rows = [self.tree.item(row)["values"] for row in self.tree.get_children()]
            if not rows:
                messagebox.showinfo("No Data", "There is no data to export.")
                return
            df = pd.DataFrame(rows, columns=self.columns)
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save as")
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Export Successful", f"Data exported to:\n{file_path}")
                self.status.config(text=f"Exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not export to Excel:\n{str(e)}")


import tkinter as tk
from tkinter import ttk
import threading
import logging
import requests
import base64
import pandas as pd
import re
import automic_rest as automic
from datetime import datetime
import pytz

class FirstrunChecker:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.cancel_operation = False
        self.timeout = (5, 30)  # 5s connect, 30s read

        # Configure grid weights for responsiveness
        self.parent.grid_rowconfigure(1, weight=1)
        self.parent.grid_rowconfigure(2, weight=0)
        self.parent.grid_columnconfigure(0, weight=1)
        self.parent.grid_columnconfigure(1, weight=1)

        # Top frame for input and output fields
        top_frame = ttk.Frame(self.parent)
        top_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=8)
        top_frame.grid_columnconfigure(0, weight=1)
        top_frame.grid_columnconfigure(1, weight=1)

        # Folder to check input
        ttk.Label(top_frame, text="Folder to check (one per line)").grid(row=0, column=0, columnspan=2, sticky="w", pady=(10, 0), padx=5)
        self.folder_input = tk.Text(top_frame, height=4, width=40, undo=True)
        self.folder_input.bind("<Control-z>", lambda e: self.folder_input.edit_undo())
        self.folder_input.bind("<Control-y>", lambda e: self.folder_input.edit_redo())
        self.folder_input.grid(row=1, column=0, columnspan=2, pady=(0, 10), sticky="nsew", padx=5)

        # Success and Failed fields (two columns)
        ttk.Label(top_frame, text="Success and to be moved").grid(row=2, column=0, sticky="w", pady=(10, 0), padx=5)
        self.success_input = tk.Text(top_frame, height=5, width=40, undo=True)
        self.success_input.bind("<Control-z>", lambda e: self.success_input.edit_undo())
        self.success_input.bind("<Control-y>", lambda e: self.success_input.edit_redo())
        self.success_input.grid(row=3, column=0, sticky="nsew", padx=(5, 2), pady=(0, 10))

        ttk.Label(top_frame, text="Failed:").grid(row=2, column=1, sticky="w", pady=(10, 0), padx=5)
        self.failed_log = tk.Text(top_frame, height=5, width=40, state='disabled')
        self.failed_log.grid(row=3, column=1, sticky="nsew", padx=(2, 5), pady=(0, 10))

        # Status Log (larger and responsive)
        ttk.Label(top_frame, text="Status Log:").grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0), padx=5)
        self.log_box = tk.Text(top_frame, height=10, width=80, state='disabled')
        self.log_box.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=5, pady=(0, 10))
        top_frame.grid_rowconfigure(5, weight=1)  # Make log_box row expandable

        # Configure tags for color coding
        self.log_box.tag_configure("success", foreground="#90EE90")  # Pastel green
        self.log_box.tag_configure("failed", foreground="#FFB6C1")   # Pastel red

        # Buttons and spinner
        button_frame = ttk.Frame(top_frame)
        button_frame.grid(row=6, column=0, columnspan=2, sticky="e", pady=5)
        self.update_button = ttk.Button(button_frame, text="📦 Firstrun check", command=self.start_check)
        self.update_button.grid(row=0, column=0, sticky="e", padx=(0, 5))
        self.cancel_button = ttk.Button(button_frame, text="Cancel check", command=self.cancel_check)
        self.cancel_button.grid(row=0, column=1, sticky="e", padx=5)
        self.cancel_button.grid_remove()

        self.spinner = ttk.Progressbar(button_frame, mode='indeterminate')
        self.spinner.grid(row=0, column=0, sticky="e", padx=(0, 5))
        self.spinner.grid_remove()

        # Status frame
        status_frame = ttk.Frame(self.parent)
        status_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))
        status_frame.grid_columnconfigure(0, weight=1)
        self.status = tk.Label(status_frame, text="", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="right", anchor="e", fill="x", expand=True)

    def find_all_general_attributes(self, json_data, client_id, folder_name):
        """Extract general_attributes and STARTTIME from JSON data."""
        if not isinstance(json_data.get('data'), list):
            self.log(f"Error: 'data' key not found or not a list for folder {folder_name}.", tag="failed")
            return []
        
        all_attributes = []
        for item in json_data['data']:
            nested_data = item.get('data', {})
            if nested_data.get('jobp') or nested_data.get('jobs'):
                for key, value in nested_data.items():
                    general_attributes = value.get('general_attributes', {})
                    if not general_attributes:
                        continue
                    starttime = None
                    for doc in value.get('documentation', []):
                        doc_key = '_BSH' if client_id == 1111 else '_STRUKTUR'
                        if doc_key in doc:
                            for line in doc[doc_key]:
                                if line.strip().startswith('<Content '):
                                    match = re.search(r'\bSTARTTIME="([^"]*)"', line)
                                    if match:
                                        starttime = match.group(1)
                                        break
                            if starttime:  # Break after finding STARTTIME
                                break
                    all_attributes.append({
                        'general_attributes': general_attributes,
                        'starttime': starttime
                    })
        
        if not all_attributes:
            self.log(f"Error: No 'general_attributes' found in folder {folder_name}.", tag="failed")
        return all_attributes

    def format_starttime(self, starttime):
        """Convert STARTTIME from 'YYYY-MM-DD hh:mm TZ.<zone>(UTC <offset>)' or 'YYYY-MM-DD hh:mm TZ.<zone>' to UC4 format 'YYYY-MM-DDThh:mm:ssZ'."""
        if not starttime:
            return None
        try:
            # Regex to match date-time and timezone parts
            match = re.match(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}) TZ\.(\w+(?:\.\w+)?)(?:\(UTC ([+-]?\d+)(?:: \w+\))?)?', starttime)
            if not match:
                self.log(f"Invalid STARTTIME format: {starttime}", tag="failed")
                return None
            
            # Extract date-time and timezone parts
            datetime_str, tz_identifier, utc_offset = match.groups()
            
            # Parse the date-time part
            parsed_time = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M')
            
            # Map timezone identifiers to pytz timezones
            tz_map = {
                'AEST2': 'Australia/Sydney',      # UTC+10
                'AKST': 'America/Anchorage',      # UTC-9
                'AST': 'Asia/Riyadh',             # UTC+3
                'AWST': 'Australia/Perth',        # UTC+8
                'BRT2': 'America/Sao_Paulo',      # UTC-3
                'BRT3': 'America/Sao_Paulo',      # UTC-3
                'CET': 'Europe/Paris',            # UTC+1
                'CLST': 'America/Santiago',       # UTC-4
                'COT': 'America/Bogota',          # UTC-5
                'CST.1': 'America/Chicago',       # UTC-6
                'CST.2': 'America/Chicago',       # UTC-6
                'CST.3': 'America/Chicago',       # UTC-6
                'CST.4': 'America/Mexico_City',   # UTC-6
                'CST.5': 'Asia/Shanghai',         # UTC+8
                'CXT': 'Indian/Christmas',        # UTC+7
                'EET': 'Europe/Bucharest',        # UTC+2
                'EST.2': 'America/New_York',      # UTC-5
                'FKT': 'Atlantic/Stanley',        # UTC-4
                'GMT': 'Etc/GMT',                 # UTC+0
                'GST': 'Asia/Dubai',              # UTC+4
                'HAST': 'Pacific/Honolulu',       # UTC-10
                'ICT': 'Asia/Bangkok',            # UTC+7
                'IST': 'Asia/Kolkata',            # UTC+5:30
                'JST': 'Asia/Tokyo',              # UTC+9
                'KST': 'Asia/Seoul',              # UTC+9
                'MSK': 'Europe/Moscow',           # UTC+3
                'MST.1': 'America/Denver',        # UTC-7
                'MST.2': 'America/Denver',        # UTC-7
                'MST.3': 'America/Mazatlan',      # UTC-7
                'MYT': 'Asia/Singapore',          # UTC+8
                'NOVT': 'Asia/Novosibirsk',       # UTC+7
                'NZST': 'Pacific/Auckland',       # UTC+12
                'PST.1': 'America/Los_Angeles',   # UTC-8
                'PST.2': 'America/Los_Angeles',   # UTC-8
                'PST.3': 'America/Tijuana',       # UTC-8
                'SAST': 'Africa/Johannesburg',    # UTC+2
                'SGT': 'Asia/Singapore',          # UTC+8
                'TKT': 'Pacific/Fakaofo',         # UTC+14
                'TRT': 'Europe/Istanbul',         # UTC+3
                'UTC': 'Etc/UTC',                 # UTC+0
                'WIB': 'Asia/Jakarta',            # UTC+7
                'WST': 'Australia/Perth',         # UTC+8
                'YEKT': 'Europe/Paris'            # UC4 uses CET (UTC+1)
            }
            
            # Determine timezone
            if utc_offset:
                # Use UTC offset if provided (e.g., UTC +8, UTC -9)
                tz = pytz.FixedOffset(int(utc_offset) * 60)
            elif tz_identifier in tz_map:
                # Use mapped timezone if no UTC offset
                tz = pytz.timezone(tz_map[tz_identifier])
            else:
                self.log(f"Unknown timezone identifier: {tz_identifier}", tag="failed")
                return None
            
            # Localize the datetime with the timezone
            localized_time = tz.localize(parsed_time)
            
            # Convert to UTC and format as UC4 format
            utc_time = localized_time.astimezone(pytz.UTC)
            uc4_time = utc_time.strftime('%Y-%m-%dT%H:%M:%SZ')
            return uc4_time
        except (ValueError, KeyError) as e:
            self.log(f"Error formatting STARTTIME '{starttime}': {str(e)}", tag="failed")
            return None

    def start_check(self):
        """Start the firstrun check process in a separate thread."""
        self.update_button.config(state='disabled')
        self.start_spinner()
        self.show_cancel_button()
        self.cancel_operation = False
        threading.Thread(target=self.check_firstrun, daemon=True).start()

    def check_firstrun(self):
        """Check firstrun status of jobs in the specified folders."""
        self.log("Starting firstrun check...", tag="normal")
        self.status_update("Checking folders...", fg="blue")

        # Clear previous results
        self.success_input.config(state='normal')
        self.success_input.delete(1.0, tk.END)
        self.failed_log.config(state='normal')
        self.failed_log.delete(1.0, tk.END)
        self.failed_log.config(state='disabled')

        # Get folder names from input (one per line)
        folder_names = [f.strip() for f in self.folder_input.get(1.0, tk.END).strip().split('\n') if f.strip()]
        if not folder_names:
            self.log("Error: No folders specified.", tag="failed")
            self.status_update("Error: No folders specified.", fg="red")
            self.stop_spinner()
            self.hide_cancel_button()
            self.update_button.config(state='normal')
            return

        # Validate client ID
        try:
            client_id = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID", tag="failed")
            self.status_update("Error: Invalid Client ID", fg="red")
            self.stop_spinner()
            self.hide_cancel_button()
            self.update_button.config(state='normal')
            return

        # Automic API setup
        api_url = f'https://rb-{self.env_var.get()}-api.bosch.com'
        user = self.entries['USERID'].get().strip()
        credentials = f"{user}:{self.entries['PASSWORD'].get()}"
        auth = base64.b64encode(credentials.encode()).decode()
        headers = {
            'Accept': 'application/json, */*',
            'Authorization': f"Basic {auth}",
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        # Connect to Automic API
        try:
            automic.connection(
                url=api_url,
                auth=auth,
                noproxy=True,
                sslverify=False,
                timeout=self.timeout[1]
            )
        except Exception as e:
            self.log(f"Authentication failed: {str(e)}", tag="failed")
            self.status_update(f"Authentication failed: {str(e)}", fg="red")
            self.stop_spinner()
            self.hide_cancel_button()
            self.update_button.config(state='normal')
            return

        # Process each folder
        success = []
        failed = []
        for folder_name in folder_names:
            if self.cancel_operation:
                self.log("Operation cancelled.", tag="failed")
                self.stop_spinner()
                self.hide_cancel_button()
                self.update_button.config(state='normal')
                return

            self.log(f"Processing folder: {folder_name}", tag="normal")

            # Fetch folder data
            url = f"{api_url}/ae/api/v1/{client_id}/folderobjects/AUTOMATION_JOBS/{user}/{folder_name}"
            try:
                r = requests.get(url, headers=headers, verify=False, timeout=self.timeout)
                r.raise_for_status()
                folder_json = r.json()
            except requests.exceptions.HTTPError as http_err:
                self.log(f"HTTP error for folder {folder_name}: {http_err}", tag="failed")
                self.status_update(f"HTTP error for folder {folder_name}", fg="red")
                continue
            except requests.exceptions.SSLError as ssl_err:
                self.log(f"SSL error for folder {folder_name}: {ssl_err}", tag="failed")
                self.status_update(f"SSL error for folder {folder_name}", fg="red")
                continue
            except requests.exceptions.RequestException as req_err:
                self.log(f"Request error for folder {folder_name}: {req_err}", tag="failed")
                self.status_update(f"Request error for folder {folder_name}", fg="red")
                continue

            # Process folder data
            results = self.find_all_general_attributes(folder_json, client_id, folder_name)
            checked = set()

            for i, entry in enumerate(results, 1):
                if self.cancel_operation:
                    self.log("Operation cancelled.", tag="failed")
                    self.stop_spinner()
                    self.hide_cancel_button()
                    self.update_button.config(state='normal')
                    return

                general_attributes = entry['general_attributes']
                starttime = entry['starttime']
                name = general_attributes['name']
                if name in checked:
                    continue
                checked.add(name)
                self.log(f"{i}: {name}", tag="normal")
                if starttime:
                    formatted_starttime = self.format_starttime(starttime)
                    if formatted_starttime:
                        self.log(f"Firstrun date: {formatted_starttime}", tag="normal")
                    else:
                        self.log(f"{name} has invalid Firstrun date: {starttime}", tag="failed")
                
                # Construct query with time_frame_from if starttime is valid
                query = f"name={name}&include_deactivated=true"
                if formatted_starttime:
                    query += f"&time_frame_from={formatted_starttime}"
                
                re = automic.listExecutions(client_id=client_id, query=query)
                if len(re.response['data']) == 0:
                    self.log(f"No executions found for {name}", tag="failed")
                    continue
                status_text = re.response['data'][0]['status_text']
                end_time = re.response['data'][0]['end_time']
                tag = "success" if "ENDED_OK" in status_text else "failed"
                self.log(f"{name} status: {status_text} (End time: {end_time})", tag=tag)

                success_flag = True
                if general_attributes['type'] == "JOBP":
                    re = automic.getChildrenOfExecution(client_id=client_id, run_id=re.response['data'][0]['run_id'])
                    for child in re.response['data']:
                        if child['name'] in checked:
                            if child['name'] in failed:
                                success_flag = False
                            continue
                        checked.add(child['name'])
                        child_tag = "success" if child['status'] == 1900 else "failed"
                        if child['status'] == 1900:
                            success.append(child['name'])
                        else:
                            success_flag = False
                            failed.append(child['name'])
                        self.log(f"Child {child['name']}: {child['status']} ({child['status_text']})", tag=child_tag)
                
                if success_flag:
                    success.append(name)
                else:
                    failed.append(name)

        # Update UI with results
        if success:
            self.success_input.config(state='normal')
            self.success_input.delete(1.0, tk.END)
            self.success_input.insert(tk.END, "\n".join(sorted(set(success))))
            self.success_input.config(state='normal')
            self.log(f"Success: {len(set(success))}", tag="normal")
        if failed:
            self.failed_log.config(state='normal')
            self.failed_log.delete(1.0, tk.END)
            self.failed_log.insert(tk.END, "\n".join(sorted(set(failed))))
            self.failed_log.config(state='disabled')
            self.log(f"Failed: {len(set(failed))}", tag="normal")

        self.log("Firstrun check completed.", tag="normal")
        self.stop_spinner()
        self.hide_cancel_button()
        self.update_button.config(state='normal')

    def start_spinner(self):
        """Start the progress bar spinner."""
        self.update_button.grid_remove()
        self.spinner.grid()
        self.spinner.start(10)

    def stop_spinner(self):
        """Stop the progress bar spinner."""
        self.spinner.stop()
        self.spinner.grid_remove()
        self.update_button.grid()

    def show_cancel_button(self):
        """Show the cancel button."""
        self.cancel_button.grid()

    def hide_cancel_button(self):
        """Hide the cancel button."""
        self.cancel_button.grid_remove()

    def cancel_check(self):
        """Set the cancel flag to stop the update operation."""
        self.cancel_operation = True
        self.status_update("Cancelling operation...", fg="orange")

    def status_update(self, message, fg="black"):
        """Update the status label with a message and color."""
        self.parent.after(0, lambda: self.status.config(text=message, fg=fg))
        self.log(message, tag="normal" if fg == "black" else "failed")

    def log(self, message, tag="normal"):
        """Log messages to the status log box with color coding."""
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, message + "\n", tag)
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')
class VaraUpdaterApp:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.cancel_operation = False

        # Configure grid weights
        self.parent.grid_rowconfigure(1, weight=1)
        self.parent.grid_rowconfigure(2, weight=0)
        self.parent.grid_columnconfigure(0, weight=1)

        # Top frame for input fields
        top_frame = ttk.Frame(self.parent)
        top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        top_frame.grid_columnconfigure(0, weight=1)
        top_frame.grid_columnconfigure(1, weight=0)

        # Multiple object names input
        ttk.Label(top_frame, text="Multiple Object Names (one per line):").grid(row=0, column=0, columnspan=2, sticky="w", pady=(10, 0), padx=5)
        self.batch_input = tk.Text(top_frame, height=4, width=80, undo=True)
        self.batch_input.bind("<Control-z>", lambda e: self.batch_input.edit_undo())
        self.batch_input.bind("<Control-y>", lambda e: self.batch_input.edit_redo())
        self.batch_input.grid(row=1, column=0, columnspan=2, pady=(0, 10), sticky="nsew", padx=5)

        # VARA object name input with Combobox
        ttk.Label(top_frame, text="VARA Object Name:").grid(row=2, column=0, sticky="w", padx=5)
        self.vara_name_combobox = ttk.Combobox(top_frame, width=40, state="normal")  # Allow text input
        self.vara_name_combobox.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        # Bind left-click to detect dropdown arrow click
        self.vara_name_combobox.bind("<Button-1>", self.on_combobox_click)
                # Loading spinner for dropdown
        self.dropdown_spinner = ttk.Progressbar(top_frame, mode='indeterminate', length=20)
        self.dropdown_spinner.grid(row=2, column=1, sticky="e", padx=(0, 5), pady=2)
        self.dropdown_spinner.grid_remove()


        # Apply custom style for tiny buttons
        style = ttk.Style()
        style.configure('Tiny.TButton', padding=0)  # Minimal padding for icon buttons
        # Start time input with button inside
        ttk.Label(top_frame, text="Start Time (YYYYMMDD;HH:MM):").grid(row=3, column=0, sticky="w", padx=5)
        start_time_frame = ttk.Frame(top_frame)
        start_time_frame.grid(row=3, column=1, sticky="w", padx=5, pady=2)
        self.start_time_entry = ttk.Entry(start_time_frame, width=18)  # Reduced width to fit button
        self.start_time_entry.pack(side="left")
        self.yesterday_button = ttk.Button(start_time_frame, text="📅", command=self.set_yesterday, width=2, style='Tiny.TButton')
        self.yesterday_button.pack(side="left", padx=(1, 0))  # Minimal padding to keep button inside

        # End time input with button inside
        ttk.Label(top_frame, text="End Time (YYYYMMDD;HH:MM):").grid(row=4, column=0, sticky="w", padx=5)
        end_time_frame = ttk.Frame(top_frame)
        end_time_frame.grid(row=4, column=1, sticky="w", padx=5, pady=2)
        self.end_time_entry = ttk.Entry(end_time_frame, width=18)  # Reduced width to fit button
        self.end_time_entry.pack(side="left")
        self.set_44440404_button = ttk.Button(end_time_frame, text="📅", command=self.set_end_time, width=2, style='Tiny.TButton')
        self.set_44440404_button.pack(side="left", padx=(1, 0))  # Minimal padding to keep button inside
        # Buttons and spinner
        self.update_button = ttk.Button(top_frame, text="📦 Update VARA", command=self.start_update)
        self.update_button.grid(row=1, column=2, sticky="e", padx=10)
        self.cancel_button = ttk.Button(top_frame, text="Cancel Update", command=self.cancel_update)
        self.cancel_button.grid(row=0, column=2, sticky="e", padx=10)
        self.cancel_button.grid_remove()

        self.spinner = ttk.Progressbar(top_frame, mode='indeterminate')
        self.spinner.grid(row=1, column=2, sticky="e", padx=10)
        self.spinner.grid_remove()

        # Status frame
        status_frame = ttk.Frame(self.parent)
        status_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 5))
        status_frame.grid_columnconfigure(0, weight=1)

        self.status = tk.Label(status_frame, text="", bd=1, relief="sunken", anchor="w")
        self.status.pack(side="right", anchor="e", fill="x", expand=True)
    def on_combobox_click(self, event):
        """Handle click on Combobox to detect dropdown arrow and populate dropdown."""
        # Get the widget's dimensions and click coordinates
        x, y = event.x, event.y
        width = self.vara_name_combobox.winfo_width()
        # Approximate the dropdown arrow area (rightmost ~20 pixels)
        if x > width - 20:
            self.start_dropdown_spinner()
            threading.Thread(target=self.populate_dropdown, daemon=True).start()
        return "continue"  # Allow normal Combobox behavior    def populate_dropdown(self):
        """Populate the Combobox dropdown based on client_id."""
    def populate_dropdown(self):
        """Populate the Combobox dropdown based on client_id in a separate thread."""
        try:
            client_id = self.client_var.get().strip()
            vara_options = []

            if client_id == "1101":
                vara_options = [
                    "AA_XXXX_VAR_WARTUNG_JOBPLAN", "AE_XXXX_VAR_WARTUNG_JOBPLAN", "ALL_XXXX_VAR_WARTUNG_JOBPLAN",
                    "AS_XXXX_VAR_WARTUNG_JOBPLAN", "BBM_XXXX_VAR_WARTUNG_JOBPLAN", "BD_XXXX_VAR_WARTUNG_JOBPLAN",
                    "BEG_XXXX_VAR_WARTUNG_JOBPLAN", "BHCS_XXXX_VAR_WARTUNG_JOBPLAN", "BT_XXXX_VAR_WARTUNG_JOBPLAN",
                    "BWC_XXXX_VAR_WARTUNG_JOBPLAN", "C_XXXX_VAR_WARTUNG_JOBPLAN", "CBN_XXXX_VAR_WARTUNG_JOBPLAN",
                    "CC_XXXX_VAR_WARTUNG_JOBPLAN", "CGL_XXXX_VAR_WARTUNG_JOBPLAN", "CI_XXXX_VAR_WARTUNG_JOBPLAN",
                    "CN_XXXX_VAR_WARTUNG_JOBPLAN", "CS_XXXX_VAR_WARTUNG_JOBPLAN", "CT_XXXX_VAR_WARTUNG_JOBPLAN",
                    "DC_XXXX_VAR_WARTUNG_JOBPLAN", "EB_XXXX_VAR_WARTUNG_JOBPLAN", "ED_XXXX_VAR_WARTUNG_JOBPLAN",
                    "ET_XXXX_VAR_WARTUNG_JOBPLAN", "ETAS_XXXX_VAR_WARTUNG_JOBPLAN", "GR_XXXX_VAR_WARTUNG_JOBPLAN",
                    "GRC_XXXX_VAR_WARTUNG_JOBPLAN", "GS_XXXX_VAR_WARTUNG_JOBPLAN", "HC_XXXX_VAR_WARTUNG_JOBPLAN",
                    "PA_XXXX_VAR_WARTUNG_JOBPLAN", "PRU_XXXX_VAR_WARTUNG_JOBPLAN", "PS_XXXX_VAR_WARTUNG_JOBPLAN",
                    "PT_XXXX_VAR_WARTUNG_JOBPLAN", "RBAC_XXXX_VAR_WARTUNG_JOBPLAN", "RBAU_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBCN_XXXX_VAR_WARTUNG_JOBPLAN", "RBEI_XXXX_VAR_WARTUNG_JOBPLAN", "RBIN_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBJP_XXXX_VAR_WARTUNG_JOBPLAN", "RBKR_XXXX_VAR_WARTUNG_JOBPLAN", "RBLA_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBMY_XXXX_VAR_WARTUNG_JOBPLAN", "RBNA_XXXX_VAR_WARTUNG_JOBPLAN", "RBS2_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBSA_XXXX_VAR_WARTUNG_JOBPLAN", "RBTW_XXXX_VAR_WARTUNG_JOBPLAN", "RG_XXXX_VAR_WARTUNG_JOBPLAN",
                    "SAR_XXXX_VAR_WARTUNG_JOBPLAN", "SCL_XXXX_VAR_WARTUNG_JOBPLAN", "SE_XXXX_VAR_WARTUNG_JOBPLAN",
                    "SEAN_XXXX_VAR_WARTUNG_JOBPLAN", "SL_XXXX_VAR_WARTUNG_JOBPLAN", "ST_XXXX_VAR_WARTUNG_JOBPLAN",
                    "SY_XXXX_VAR_WARTUNG_JOBPLAN", "TT_XXXX_VAR_WARTUNG_JOBPLAN", "XC_XXXX_VAR_WARTUNG_JOBPLAN",
                    "XX_XXXX_VAR_WARTUNG_JOBPLAN", "ZT_XXXX_VAR_WARTUNG_JOBPLAN"
                ]
            elif client_id == "1001":
                vara_options = [
                    "AA_XXXX_VAR_WARTUNG_JOBPLAN", "AE_XXXX_VAR_WARTUNG_JOBPLAN", "ALL_VAR_WARTUNG_JOBPLAN_VARIABLES",
                    "APM_XXXX_VAR_WARTUNG_JOBPLAN", "AS_XXXX_VAR_WARTUNG_JOBPLAN", "ATMO_XXXX_VAR_WARTUNG_JOBPLAN",
                    "BBM_XXXX_VAR_WARTUNG_JOBPLAN", "BD_XXXX_VAR_WARTUNG_JOBPLAN", "BEG_XXXX_VAR_WARTUNG_JOBPLAN",
                    "BMG_XXXX_VAR_WARTUNG_JOBPLAN", "BRJP_XXXX_VAR_WARTUNG_JOBPLAN", "BST_XXXX_VAR_WARTUNG_JOBPLAN",
                    "BT_XXXX_VAR_WARTUNG_JOBPLAN", "C_XXXX_VAR_WARTUNG_JOBPLAN", "CC_XXXX_VAR_WARTUNG_JOBPLAN",
                    "CGL_XXXX_VAR_WARTUNG_JOBPLAN", "CI_XXXX_VAR_WARTUNG_JOBPLAN", "CM_XXXX_VAR_WARTUNG_JOBPLAN",
                    "COL_XXXX_VAR_WARTUNG_JOBPLAN", "CS_XXXX_VAR_WARTUNG_JOBPLAN", "CT_XXXX_VAR_WARTUNG_JOBPLAN",
                    "CU_XXXX_VAR_WARTUNG_JOBPLAN", "DC_XXXX_VAR_WARTUNG_JOBPLAN", "DS_XXXX_VAR_WARTUNG_JOBPLAN",
                    "EB_XXXX_VAR_WARTUNG_JOBPLAN", "EBI_XXXX_VAR_WARTUNG_JOBPLAN", "ED_XXXX_VAR_WARTUNG_JOBPLAN",
                    "EDDI_XXXX_VAR_WARTUNG_JOBPLAN", "EK_XXXX_VAR_WARTUNG_JOBPLAN", "EPS_XXXX_VAR_WARTUNG_JOBPLAN",
                    "ET_XXXX_VAR_WARTUNG_JOBPLAN", "ETAS_XXXX_VAR_WARTUNG_JOBPLAN", "FIT_CM_XXXX_VAR_WARTUNG_JOBPLAN_REMAS",
                    "GRC_XXXX_VAR_WARTUNG_JOBPLAN", "GS_XXXX_VAR_WARTUNG_JOBPLAN", "INST_XXXX_VAR_WARTUNG_JOBPLAN",
                    "PA_XXXX_VAR_WARTUNG_JOBPLAN", "PA_XXXX_VAR_WARTUNG_JOBPLAN.BACKUP", "PS_XXXX_VAR_WARTUNG_JOBPLAN",
                    "PT_XXXX_VAR_WARTUNG_JOBPLAN", "PT_XXXX_VAR_WARTUNG_JOBPLAN_HOST", "QI_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBAC_XXXX_VAR_WARTUNG_JOBPLAN", "RBAU_XXXX_VAR_WARTUNG_JOBPLAN", "RBCC_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBCD_XXXX_VAR_WARTUNG_JOBPLAN", "RBCN_XXXX_VAR_WARTUNG_JOBPLAN", "RBEI_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBID_XXXX_VAR_WARTUNG_JOBPLAN", "RBIN_XXXX_VAR_WARTUNG_JOBPLAN", "RBJP_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBKR_XXXX_VAR_WARTUNG_JOBPLAN", "RBLA_XXXX_VAR_WARTUNG_JOBPLAN", "RBMY_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBNA_XXXX_VAR_WARTUNG_JOBPLAN", "RBPL_XXXX_VAR_WARTUNG_JOBPLAN", "RBSA_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RBTH_XXXX_VAR_WARTUNG_JOBPLAN", "RBTW_XXXX_VAR_WARTUNG_JOBPLAN", "RBVN_XXXX_VAR_WARTUNG_JOBPLAN",
                    "RG_XXXX_VAR_WARTUNG_JOBPLAN", "SAR_XXXX_VAR_WARTUNG_JOBPLAN", "SCL_XXXX_VAR_WARTUNG_JOBPLAN",
                    "SG_XXXX_VAR_WARTUNG_JOBPLAN", "SL_XXXX_VAR_WARTUNG_JOBPLAN", "ST_XXXX_VAR_WARTUNG_JOBPLAN",
                    "SY_XXXX_VAR_WARTUNG_JOBPLAN", "TT_XXXX_VAR_WARTUNG_JOBPLAN", "XC_XXXX_VAR_WARTUNG_JOBPLAN",
                    "XX_XXXX_VAR_WARTUNG_JOBPLAN", "ZR_XXXX_VAR_WARTUNG_JOBPLAN", "ZT_XXXX_VAR_WARTUNG_JOBPLAN"
                ]
            elif client_id == "1111":
                vara_options = [
                    "ALL_VAR_WARTUNG_JOBPLAN_VARIABLES", "BSH_XXXX_VAR_WARTUNG_JOBPLAN", "CI_XXXX_VAR_WARTUNG_JOBPLAN",
                    "XX_XXXX_VAR_WARTUNG_JOBPLAN"
                ]

            # Sort options alphabetically
            vara_options = sorted(vara_options)

            # Update Combobox values on the main thread
            self.parent.after(0, lambda: self.vara_name_combobox.configure(values=vara_options))
        finally:
            # Stop spinner on the main thread
            self.parent.after(0, self.stop_dropdown_spinner)
    def start_dropdown_spinner(self):
        """Start the dropdown loading spinner."""
        self.dropdown_spinner.grid()
        self.dropdown_spinner.start(10)

    def stop_dropdown_spinner(self):
        """Stop the dropdown loading spinner."""
        self.dropdown_spinner.stop()
        self.dropdown_spinner.grid_remove()
    def set_yesterday(self):
        """Set the start time entry to yesterday's date with time 00:00."""
        yesterday = datetime.now() - timedelta(days=1)
        formatted_date = yesterday.strftime("%Y%m%d;00:00")
        self.start_time_entry.delete(0, tk.END)
        self.start_time_entry.insert(0, formatted_date)
    def set_end_time(self):
        """Set the end time entry to 44440404;04:04."""
        self.end_time_entry.delete(0, tk.END)
        self.end_time_entry.insert(0, "44440404;04:04")
    def start_update(self):
        """Start the VARA update process in a separate thread."""
        self.update_button.config(state='disabled')
        self.start_spinner()
        self.show_cancel_button()
        threading.Thread(target=self.update_vara, daemon=True).start()

    def update_vara(self):
        """Fetch the VARA object, update its static_values, and post it back."""
        try:
            # Retrieve input values
            client_id = self.client_var.get().strip()
            userid = self.entries['USERID'].get().strip()
            password = self.entries['PASSWORD'].get().strip()
            env = self.env_var.get().strip()
            object_names = [name.strip() for name in self.batch_input.get("1.0", tk.END).strip().splitlines() if name.strip()]
            vara_name = self.vara_name_combobox.get().strip()
            start_time = self.start_time_entry.get().strip()
            end_time = self.end_time_entry.get().strip()

            # Validate inputs
            if not client_id or not userid or not password:
                self.status_update("Error: Client ID, User ID, and Password are required.", fg="red")
                messagebox.showerror("Error", "Please provide Client ID, User ID, and Password.")
                return
            if not object_names:
                self.status_update("Error: At least one object name is required.", fg="red")
                messagebox.showinfo("Input Missing", "Please enter at least one object name.")
                return
            if not vara_name:
                self.status_update("Error: VARA object name is required.", fg="red")
                messagebox.showinfo("Input Missing", "Please enter a VARA object name.")
                return
            if not start_time or not end_time:
                self.status_update("Error: Start time and end time are required.", fg="red")
                messagebox.showinfo("Input Missing", "Please enter both start time and end time.")
                return

            # Validate time formats (YYYYMMDD;HH:MM)
            try:
                datetime.strptime(start_time, "%Y%m%d;%H:%M")
                datetime.strptime(end_time, "%Y%m%d;%H:%M")
            except ValueError:
                self.status_update("Error: Invalid time format. Use YYYYMMDD;HH:MM (e.g., 20231212;00:01).", fg="red")
                messagebox.showerror("Error", "Start time and end time must be in YYYYMMDD;HH:MM format.")
                return

            # Authenticate with Automic API
            auth = base64.b64encode(f"{userid}:{password}".encode()).decode()
            url = f"https://rb-{env}-api.bosch.com"
            try:
                automic.connection(url=url, auth=auth, noproxy=True, sslverify=False, cert="/path/to/certfile", timeout=60)
            except Exception as e:
                self.status_update(f"Authentication failed: {str(e)}", fg="red")
                messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}.")
                return

            if self.cancel_operation:
                self.status_update("Operation cancelled.", fg="orange")
                return

            # Fetch the VARA object
            logging.info(f"Fetching VARA object: {vara_name}")
            try:
                result = automic.getObjects(client_id=int(client_id), object_name=vara_name)
                if result.status != 200:
                    self.status_update(f"Failed to fetch VARA {vara_name}: Status {result.status}", fg="red")
                    messagebox.showerror("Error", f"Failed to fetch VARA {vara_name}: Status {result.status}")
                    return
                if 'data' not in result.response or 'vara' not in result.response['data']:
                    self.status_update(f"VARA {vara_name} not found or invalid response.", fg="red")
                    messagebox.showerror("Error", f"VARA {vara_name} not found or invalid response from server.")
                    return
                vara_data = result.response['data']['vara']
            except Exception as e:
                self.status_update(f"Error fetching VARA {vara_name}: {str(e)}", fg="red")
                messagebox.showerror("Error", f"Failed to fetch VARA {vara_name}: {str(e)}")
                return

            if self.cancel_operation:
                self.status_update("Operation cancelled.", fg="orange")
                return

            # Update static_values
            static_values = vara_data.get('static_values', [])
            value1 = f"{start_time};{end_time}"
            existing_keys = {entry['key'] for entry in static_values}

            new_entries = []
            seen_keys = set()  # Track unique keys in new_entries
            for obj_name in object_names:
                if obj_name in existing_keys:
                    logging.warning(f"Skipping {obj_name}: Already exists in static_values.")
                    self.status_update(f"Warning: {obj_name} already exists in VARA, skipping.", fg="orange")
                    continue
                if obj_name in seen_keys:
                    logging.warning(f"Skipping {obj_name}: Duplicate in input object names.")
                    self.status_update(f"Warning: {obj_name} is a duplicate in input, skipping.", fg="orange")
                    continue
                new_entry = {
                    'key': obj_name,
                    'validity_range': 'FREI',
                    'value1': value1
                }
                new_entries.append(new_entry)
                static_values.append(new_entry)
                seen_keys.add(obj_name)
                logging.info(f"Added entry for {obj_name} with value1: {value1}")

            if not new_entries:
                self.status_update("No new entries added: All object names already exist.", fg="orange")
                messagebox.showinfo("No Changes", "All specified object names already exist in the VARA object.")
                return

            if self.cancel_operation:
                self.status_update("Operation cancelled.", fg="orange")
                return

            # Post the updated VARA object
            logging.info(f"Posting updated VARA object: {vara_name}")
            try:
                if int(client_id) == 1111 and result.response.get('path'):
                    result.response['path'] = "/CENTRAL_OBJECTS/VARIABLES"
                response = automic.postObjects(client_id=int(client_id), body=result.response, query="overwrite_existing_objects=true")
                if response.status is None or response.status == 200:
                    self.status_update(f"Successfully updated VARA {vara_name} with {len(new_entries)} new entries.", fg="green")
                    messagebox.showinfo("Success", f"VARA {vara_name} updated with {len(new_entries)} new entries.")
                else:
                    self.status_update(f"Failed to update VARA {vara_name}: Status {response.status}", fg="red")
                    messagebox.showerror("Error", f"Failed to update VARA {vara_name}: Status {response.status}")
            except Exception as e:
                self.status_update(f"Error updating VARA {vara_name}: {str(e)}", fg="red")
                messagebox.showerror("Error", f"Failed to update VARA {vara_name}: {str(e)}")

        except Exception as e:
            self.status_update(f"Unexpected error: {str(e)}", fg="red")
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
        finally:
            self.parent.after(0, self.stop_spinner)
            self.parent.after(0, self.hide_cancel_button)
            self.parent.after(0, lambda: self.update_button.config(state='normal'))

    def start_spinner(self):
        """Start the progress bar spinner."""
        self.update_button.grid_remove()
        self.spinner.grid()
        self.spinner.start(10)

    def stop_spinner(self):
        """Stop the progress bar spinner."""
        self.spinner.stop()
        self.spinner.grid_remove()
        self.update_button.grid()

    def show_cancel_button(self):
        """Show the cancel button."""
        self.cancel_button.grid()

    def hide_cancel_button(self):
        """Hide the cancel button."""
        self.cancel_button.grid_remove()

    def cancel_update(self):
        """Set the cancel flag to stop the update operation."""
        self.cancel_operation = True
        self.status_update("Cancelling operation...", fg="orange")

    def status_update(self, message, fg="black"):
        """Update the status label with a message and color."""
        self.parent.after(0, lambda: self.status.config(text=message, fg=fg))
        logging.info(message)
import re
import tkinter as tk
from tkinter import ttk, messagebox
import base64
import automic_rest as automic
from xml.etree import ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from threading import Thread
from datetime import date
import pandas as pd
from datetime import datetime

class JobpStrukturUpdater:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.undo_stack = []
        self.redo_stack = []
        self.failed_jobs = []
        self.job_structures = {}  # Store _STRUKTUR for each job
        self.job_attributes = {}  # Store {jobname: {attr_name: value}} from _STRUKTUR
        self.all_options = [''] + ['ACTIVATION_MODE', 'CONTACTMAIL_DISTRIBUTIONLIST', 'CONTACT_SOLUTIONGROUP', 'COSTCENTER_CC', 'CREATED_BY', 'CUSTOMER_NAME']
        self.column_configs = ['jobname', None, None, None, None, None]  # Track selected attributes
        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        # Frame for dropdown headers
        header_frame = ttk.Frame(frm)
        header_frame.grid(row=0, column=0, columnspan=6, sticky='nsew')

        # Define initial columns
        self.columns = ('jobname', 'attribute1', 'attribute2', 'attribute3', 'attribute4', 'attribute5')
        self.attributes_tree = ttk.Treeview(frm, columns=self.columns, show='headings')
        self.attributes_tree.heading('jobname', text='Job Name')
        self.attributes_tree.column('jobname', width=200, stretch=True)
        for col in self.columns[1:]:
            self.attributes_tree.heading(col, text=col)
            self.attributes_tree.column(col, width=150, stretch=True)
        self.attributes_tree.grid(row=1, column=0, columnspan=6, sticky='nsew')

        # Dropdowns for attribute selection
        self.dropdowns = []
        ttk.Label(header_frame, text="Select Attributes:").grid(row=0, column=0, sticky='w', pady=5)
        for i, col in enumerate(self.columns, 0):
            var = tk.StringVar()
            if i == 0:  # jobname column
                menu = ttk.OptionMenu(header_frame, var, 'Job Name', 'Job Name')
                menu.config(state='disabled')
            else:
                menu = ttk.OptionMenu(header_frame, var, None, *self.all_options, command=lambda val, idx=i: self.update_column(idx, val))
                self.dropdowns.append((var, menu))
            menu.grid(row=0, column=i, padx=2, sticky='we')

        # Configure column widths and weights
        for i in range(6):
            header_frame.grid_columnconfigure(i, weight=1, minsize=self.attributes_tree.column(self.columns[i], 'width'))
            frm.grid_columnconfigure(i, weight=1)
        frm.grid_rowconfigure(0, weight=0)
        frm.grid_rowconfigure(1, weight=1)

        # Context menu for pasting
        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)
        self.attributes_tree.bind("<Button-3>", self.show_paste_menu)
        self.attributes_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='jobname'))
        self.attributes_tree.bind("<Double-1>", self.on_double_click)
        self.attributes_tree.bind("<Control-z>", self.undo)
        self.attributes_tree.bind("<Control-y>", self.redo)
        self.attributes_tree.bind("<Control-a>", self.select_all)
        self.attributes_tree.bind("<Delete>", self.delete_selected)
        self.attributes_tree.bind("<BackSpace>", self.delete_selected)
        self.attributes_tree.bind("<<TreeviewSelect>>", self.on_treeview_select)

        # Button frame
        button_frame = ttk.Frame(frm)
        button_frame.grid(row=2, column=0, columnspan=6, sticky='ew', pady=5)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        self.update_btn = ttk.Button(button_frame, text='Update Jobs', command=self.start_update, width=10)
        self.update_btn.grid(row=0, column=0, padx=(0, 2), sticky='e')
        self.see_jobs_btn = ttk.Button(button_frame, text='See Jobp', command=self.start_fetch_jobs, width=10)
        self.see_jobs_btn.grid(row=0, column=1, padx=(2, 0), sticky='w')
        self.export_btn = ttk.Button(button_frame, text='Export to Excel', command=self.export_to_excel, width=12)
        self.export_btn.grid(row=0, column=2, padx=(5, 0), sticky='w')

        self.copy_failed_btn = ttk.Button(frm, text='Copy Failed', command=self.copy_failed_jobs, state='disabled', width=10)
        self.copy_failed_btn.grid(row=3, column=0, columnspan=6, sticky='ew', pady=5)

        ttk.Label(frm, text='Update Message:').grid(row=4, column=0, sticky='w')
        self.update_message_entry = ttk.Entry(frm)
        self.update_message_entry.grid(row=4, column=1, columnspan=5, sticky='ew', padx=5)

        ttk.Label(frm, text="ARMT No:").grid(row=5, column=0, sticky="w", padx=5)
        self.armt_no_entry = ttk.Entry(frm)
        self.armt_no_entry.grid(row=5, column=1, columnspan=5, sticky='ew', padx=5)

        # Frame for log and structure views
        self.view_frame = ttk.Frame(frm)
        self.view_frame.grid(row=6, column=0, columnspan=6, sticky='nsew', pady=5)
        self.view_frame.grid_rowconfigure(1, weight=1)
        self.view_frame.grid_columnconfigure(0, weight=1)

        # Update Log view
        self.log_label = ttk.Label(self.view_frame, text='Update Log:')
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box = tk.Text(self.view_frame, height=10, state='disabled')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

        # Structure view
        self.structure_label = ttk.Label(self.view_frame, text='Job Structure (XML):')
        self.structure_text = tk.Text(self.view_frame, height=10, state='disabled')

    def parse_and_modify_xml_attributes(self, xml_lines, target_attribute, new_value):
        attributes = {}
        modified_lines = xml_lines.copy()
        pattern = r'(\w+)="([^"]*)"'
        for i, line in enumerate(xml_lines):
            if line.strip().startswith('<Content '):
                matches = re.findall(pattern, line)
                for key, value in matches:
                    attributes[key] = value
                if target_attribute in attributes:
                    attributes[target_attribute] = new_value
                    modified_line = '<Content'
                    for key, value in attributes.items():
                        modified_line += f' {key}="{value}"'
                    closing_part = line[line.find('>'):]
                    modified_line += closing_part
                    modified_lines[i] = modified_line
                    break
        return attributes, modified_lines

    def update_column(self, col_index, value):
        self.column_configs[col_index] = value if value else None
        col_name = self.columns[col_index]
        self.attributes_tree.heading(col_name, text=value if value else col_name)
        self.log(f"Column {col_index} set to {value}")

    def update_dropdowns(self):
        for i, (var, menu) in enumerate(self.dropdowns, 1):
            menu['menu'].delete(0, 'end')
            for option in self.all_options:
                menu['menu'].add_command(label=option, command=tk._setit(var, option, lambda val, idx=i: self.update_column(idx, val)))

    def load_config(self, config):
        if 'ARMT_NO' in config:
            self.armt_no_entry.insert(0, config['ARMT_NO'])
        if 'COLUMN_CONFIGS' in config:
            for i, attr in enumerate(config['COLUMN_CONFIGS'][1:], 1):
                if attr in self.all_options:
                    self.dropdowns[i-1][0].set(attr)
                    self.update_column(i, attr)

    def save_config(self):
        return {
            'ARMT_NO': self.armt_no_entry.get(),
            'COLUMN_CONFIGS': self.column_configs
        }

    def show_paste_menu(self, event):
        self.clicked_column = self.attributes_tree.identify_column(event.x)[1:]
        self.paste_menu.tk_popup(event.x_root, event.y_root)

    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        columns = self.attributes_tree['columns']
        if start_column is None:
            if hasattr(self, 'clicked_column') and self.clicked_column:
                col_index = int(self.clicked_column) - 1
                start_column = columns[col_index]
            else:
                start_column = 'jobname'
        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return
        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return
            existing_rows = list(self.attributes_tree.get_children())
            max_columns = len(columns) - start_col_index
            for i, line in enumerate(lines):
                clipboard_values = line.split('\t') if '\t' in line else [line]
                clipboard_values = clipboard_values[:max_columns]
                while len(clipboard_values) < max_columns:
                    clipboard_values.append('')
                if i < len(existing_rows):
                    row_id = existing_rows[i]
                    current_values = list(self.attributes_tree.item(row_id)['values'])
                    for j, value in enumerate(clipboard_values):
                        current_values[start_col_index + j] = value.strip()
                    self.attributes_tree.item(row_id, values=current_values)
                else:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(clipboard_values):
                        new_values[start_col_index + j] = value.strip()
                    self.attributes_tree.insert('', 'end', values=new_values)
            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
            self.job_structures.clear()
            self.structure_text.config(state='normal')
            self.structure_text.delete(1.0, tk.END)
            self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
            self.structure_text.config(state='disabled')
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
            self.undo_stack.pop()

    def on_double_click(self, event):
        row_id = self.attributes_tree.identify_row(event.y)
        if not row_id:
            return
        column = self.attributes_tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.attributes_tree['columns'][column_num]
        x, y, width, height = self.attributes_tree.bbox(row_id, column)
        entry = ttk.Entry(self.parent)
        entry.place(x=x, y=y, width=width, height=height)
        current_value = self.attributes_tree.item(row_id)['values'][column_num]
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        entry.bind("<Return>", lambda e: self.save_edit(entry, row_id, column_name))
        entry.bind("<FocusOut>", lambda e: self.save_edit(entry, row_id, column_name))
        entry.bind("<Escape>", lambda e: entry.destroy())

    def save_edit(self, entry, row_id, column_name):
        new_value = entry.get().strip()
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.attributes_tree.set(row_id, column_name, new_value)
        self.structure_text.config(state='normal')
        self.structure_text.delete(1.0, tk.END)
        self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
        self.structure_text.config(state='disabled')
        entry.destroy()

    def save_state(self):
        return [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]

    def restore_state(self, state):
        self.attributes_tree.delete(*self.attributes_tree.get_children())
        for values in state:
            self.attributes_tree.insert('', 'end', values=values)

    def undo(self, event=None):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event=None):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def select_all(self, event):
        self.attributes_tree.selection_set(self.attributes_tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.attributes_tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.attributes_tree.delete(item)
            current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
            jobs_to_remove = set(self.job_structures.keys()) - current_jobs
            for job in jobs_to_remove:
                self.job_structures.pop(job, None)
            self.structure_text.config(state='normal')
            self.structure_text.delete(1.0, tk.END)
            self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
            self.structure_text.config(state='disabled')
        return 'break'

    def show_log_view(self):
        self.structure_label.grid_remove()
        self.structure_text.grid_remove()
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

    def show_structure_view(self):
        self.log_label.grid_remove()
        self.log_box.grid_remove()
        self.structure_label.grid(row=0, column=0, sticky='nw')
        self.structure_text.grid(row=1, column=0, sticky='nsew', padx=5)

    def start_update(self):
        self.update_btn.config(state='disabled')
        self.job_structures.clear()
        self.structure_text.config(state='normal')
        self.structure_text.delete(1.0, tk.END)
        self.structure_text.config(state='disabled')
        self.show_log_view()
        Thread(target=self.execute_update, daemon=True).start()

    def execute_update(self):
        self.failed_jobs = []
        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No data to update.")
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False,cert = "RB RootCA RSA G01-pem.cer")
        except Exception as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        update_message = self.update_message_entry.get().strip()
        armt_no = self.armt_no_entry.get().strip()
        current_date = date.today().strftime("%d-%b-%Y")
        name = self.entries.get('NAME', tk.StringVar(value='Unknown')).get().strip()
        doku_entry = f"{armt_no} {current_date} {name} {update_message}" if update_message else ""

        for row in rows:
            jobname = str(row[0]).strip()
            if not jobname:
                self.log("Skipping row with empty jobname")
                continue

            self.log(f"Processing {jobname}")
            try:
                resp = automic.getObjects(client_id=cid, object_name=jobname)
                if resp.status != 200:
                    self.log(f"Failed to fetch {jobname}: {resp.status}")
                    self.failed_jobs.append(jobname)
                    continue
                if 'data' not in resp.response or 'jobp' not in resp.response['data']:
                    self.log(f"Error: Object {jobname} response missing required keys")
                    self.failed_jobs.append(jobname)
                    continue
                jobp = resp.response["data"]["jobp"]
                updated = False
                struktur_updated = False
                for doc in jobp.get('documentation', []):
                    if '_STRUKTUR' in doc:
                        xml_lines = doc['_STRUKTUR']
                        for col_idx in range(1, 6):
                            attr_name = self.column_configs[col_idx]
                            if attr_name and attr_name in self.all_options[1:]:  # Exclude empty option
                                attr_value = str(row[col_idx]).strip()
                                if attr_value:
                                    _, modified_xml_lines = self.parse_and_modify_xml_attributes(xml_lines, attr_name, attr_value)
                                    if modified_xml_lines != xml_lines:
                                        doc['_STRUKTUR'] = modified_xml_lines
                                        self.log(f"Updated {attr_name} to {attr_value} for {jobname}")
                                        updated = True
                                        struktur_updated = True
                if doku_entry and struktur_updated:
                    for doc in jobp.get('documentation', []):
                        if '_STRUKTUR' in doc:
                            struktur = doc['_STRUKTUR']
                            for i, line in enumerate(struktur):
                                if '</HINTS_CHARACTERISTICS>' in line:
                                    content, tag = line.split('</HINTS_CHARACTERISTICS>')
                                    new_content = f'{content.strip()}\n{doku_entry}'
                                    struktur[i] = f'{new_content}</HINTS_CHARACTERISTICS>{tag}'
                                    updated = True
                                    break
                if updated or update_message:
                    resp_update = automic.postObjects(client_id=cid, body=resp.response, query="overwrite_existing_objects=true")
                    if resp_update.status is None:
                        self.log(f"Successfully updated {jobname}")
                    else:
                        self.log(f"Failed to update {jobname}: {resp_update.status}")
                        self.failed_jobs.append(jobname)
                else:
                    self.log(f"No updates applied for {jobname}")
            except Exception as e:
                self.log(f"Error updating {jobname}: {str(e)}")
                self.failed_jobs.append(jobname)

        self.log("All updates applied.")
        self.parent.after(0, lambda: self.update_btn.config(state='normal'))
        self.parent.after(0, lambda: self.copy_failed_btn.config(state='normal' if self.failed_jobs else 'disabled'))

    def start_fetch_jobs(self):
        self.see_jobs_btn.config(state='disabled')
        self.show_structure_view()
        Thread(target=self.fetch_jobs, daemon=True).start()

    def fetch_jobs(self):
        self.job_structures.clear()
        self.job_attributes.clear()
        dynamic_attributes = set()

        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No jobs to fetch.")
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_structure_text_for_selection)
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_structure_text_for_selection)
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False,cert = "RB RootCA RSA G01-pem.cer")
        except Exception as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials"))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_structure_text_for_selection)
            return

        def fetch_single_job(jobname):
            if not jobname:
                return None, None, None
            try:
                self.log(f"Fetching {jobname}")
                resp = automic.getObjects(client_id=cid, object_name=jobname)
                if resp.status != 200:
                    self.log(f"Failed to fetch {jobname}: {resp.status}")
                    return jobname, None, None
                if 'data' not in resp.response or 'jobp' not in resp.response['data']:
                    self.log(f"Error: Object {jobname} response missing required keys")
                    return jobname, None, None
                jobp = resp.response["data"]["jobp"]
                structure_lines = None
                job_attributes = {}
                for doc in jobp.get('documentation', []):
                    if '_STRUKTUR' in doc:
                        structure_lines = doc['_STRUKTUR']
                        for line in structure_lines:
                            if line.strip().startswith('<Content '):
                                matches = re.findall(r'(\w+)="([^"]*)"', line)
                                for key, value in matches:
                                    job_attributes[key] = value
                                    dynamic_attributes.add(key)
                                    self.log(f"Extracted {key}={value} for job {jobname}")
                        break
                return jobname, structure_lines, job_attributes
            except Exception as e:
                self.log(f"Error fetching {jobname}: {str(e)}")
                return jobname, None, None

        jobnames = [str(row[0]).strip() for row in rows if str(row[0]).strip()]
        max_workers = min(len(jobnames), 10)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_jobname = {executor.submit(fetch_single_job, jobname): jobname for jobname in jobnames}
            for future in as_completed(future_to_jobname):
                jobname, structure_lines, job_attrs = future.result()
                if structure_lines is not None:
                    self.job_structures[jobname] = structure_lines
                if job_attrs is not None:
                    self.job_attributes[jobname] = job_attrs

        self.all_options = [''] + sorted(list(dynamic_attributes)) 
        self.parent.after(0, self.update_dropdowns)
        self.log(f"Updated dropdown options: {self.all_options}")
        self.log("All jobs fetched.")
        self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
        self.parent.after(0, self.update_structure_text_for_selection)

    def export_to_excel(self):
        if not self.job_attributes:
            messagebox.showinfo("Export Info", "No job attribute data to export. Click 'See Jobp' to fetch data.")
            self.log("No job attribute data to export.")
            return
        try:
            job_names = sorted(self.job_attributes.keys())
            all_attrs = sorted(set(attr for job_attrs in self.job_attributes.values() for attr in job_attrs))
            data = []
            for job in job_names:
                row = {'Jobp Name': job}
                for attr in all_attrs:
                    row[attr] = self.job_attributes[job].get(attr, '')
                data.append(row)
            df = pd.DataFrame(data, columns=['Jobp Name'] + all_attrs)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"jobp_struktur_attributes_{timestamp}.xlsx"
            df.to_excel(filename, index=False)
            self.log(f"Exported job attributes to {filename}")
            messagebox.showinfo("Export Success", f"Successfully exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to Excel: {str(e)}")
            messagebox.showerror("Export Error", f"Failed to export to Excel: {str(e)}")

    def update_structure_text_for_selection(self):
        selected = self.attributes_tree.selection()
        self.structure_text.config(state='normal')
        self.structure_text.delete(1.0, tk.END)
        if not selected:
            self.structure_text.config(state='disabled')
            return
        row_id = selected[0]
        jobname = self.attributes_tree.item(row_id)['values'][0].strip()
        if jobname in self.job_structures:
            structure_lines = self.job_structures[jobname]
            attributes = []
            pattern = r'(\w+)="([^"]*)"'
            for line in structure_lines:
                if line.strip().startswith('<Content '):
                    matches = re.findall(pattern, line)
                    attributes.extend([f"{key}={value}" for key, value in matches])
            if attributes:
                self.structure_text.insert(tk.END, "\n".join(attributes))
            else:
                self.structure_text.insert(tk.END, "No attributes found in structure data.")
        else:
            self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
        self.structure_text.config(state='disabled')

    def on_treeview_select(self, event):
        self.update_structure_text_for_selection()

    def copy_failed_jobs(self):
        if self.failed_jobs:
            failed_list = "\n".join(self.failed_jobs)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(failed_list)
            self.log("Copied failed jobp to clipboard.")
        else:
            messagebox.showinfo("Info", "No failed jobp to copy.")

    def log(self, msg):
        self.parent.after(0, lambda: self._log(msg))

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
class UnifiedUpdaterApp:
    def __init__(self, parent, env_var, client_var, entries):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.undo_stack = []
        self.redo_stack = []
        self.failed_jobs = []
        self.job_structures = {}  # Store _STRUKTUR or _BSH for each job
        self.job_attributes = {}  # Store {jobname: {attr_name: value}}
        self.all_options = [''] + ['ACTIVATION_MODE', 'CONTACTMAIL_DISTRIBUTIONLIST', 'CONTACT_SOLUTIONGROUP', 
                                 'COSTCENTER_CC', 'CREATED_BY', 'CUSTOMER_NAME', 'MAIL_ADDRESS', 
                                 'AGGREG_LEVEL', 'IT_PRODUCT', 'RUNTIMELIMIT_RECIPIENT', 'ALERT_TYPE']
        self.column_configs = ['jobname', None, None, None, None, None]  # Track selected attributes
        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)

        # Frame for dropdown headers
        header_frame = ttk.Frame(frm)
        header_frame.grid(row=0, column=0, columnspan=6, sticky='nsew')

        # Define initial columns
        self.columns = ('jobname', 'attribute1', 'attribute2', 'attribute3', 'attribute4', 'attribute5')
        self.attributes_tree = ttk.Treeview(frm, columns=self.columns, show='headings')
        self.attributes_tree.heading('jobname', text='Job Name')
        self.attributes_tree.column('jobname', width=200, stretch=True)
        for col in self.columns[1:]:
            self.attributes_tree.heading(col, text=col)
            self.attributes_tree.column(col, width=150, stretch=True)
        self.attributes_tree.grid(row=1, column=0, columnspan=6, sticky='nsew')

        # Dropdowns for attribute selection
        self.dropdowns = []
        ttk.Label(header_frame, text="Select Attributes:").grid(row=0, column=0, sticky='w', pady=5)
        for i, col in enumerate(self.columns, 0):
            var = tk.StringVar()
            if i == 0:  # jobname column
                menu = ttk.OptionMenu(header_frame, var, 'Job Name', 'Job Name')
                menu.config(state='disabled')
            else:
                menu = ttk.OptionMenu(header_frame, var, None, *self.all_options, command=lambda val, idx=i: self.update_column(idx, val))
                self.dropdowns.append((var, menu))
            menu.grid(row=0, column=i, padx=2, sticky='we')

        # Configure column widths and weights
        for i in range(6):
            header_frame.grid_columnconfigure(i, weight=1, minsize=self.attributes_tree.column(self.columns[i], 'width'))
            frm.grid_columnconfigure(i, weight=1)
        frm.grid_rowconfigure(0, weight=0)
        frm.grid_rowconfigure(1, weight=1)

        # Context menu for pasting
        self.paste_menu = tk.Menu(frm, tearoff=0)
        self.paste_menu.add_command(label="Paste", command=self.paste_from_clipboard)
        self.attributes_tree.bind("<Button-3>", self.show_paste_menu)
        self.attributes_tree.bind("<Control-v>", lambda event: self.paste_from_clipboard(start_column='jobname'))
        self.attributes_tree.bind("<Double-1>", self.on_double_click)
        self.attributes_tree.bind("<Control-z>", self.undo)
        self.attributes_tree.bind("<Control-y>", self.redo)
        self.attributes_tree.bind("<Control-a>", self.select_all)
        self.attributes_tree.bind("<Delete>", self.delete_selected)
        self.attributes_tree.bind("<BackSpace>", self.delete_selected)
        self.attributes_tree.bind("<<TreeviewSelect>>", self.on_treeview_select)

        # Button frame
        button_frame = ttk.Frame(frm)
        button_frame.grid(row=2, column=0, columnspan=6, sticky='ew', pady=5)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        self.update_btn = ttk.Button(button_frame, text='Update Attributes', command=self.start_update, width=10)
        self.update_btn.grid(row=0, column=0, padx=(0, 2), sticky='e')
        self.see_jobs_btn = ttk.Button(button_frame, text='See Attributes', command=self.start_fetch_jobs, width=10)
        self.see_jobs_btn.grid(row=0, column=1, padx=(2, 0), sticky='w')
        self.export_btn = ttk.Button(button_frame, text='Export to Excel', command=self.export_to_excel, width=12)
        self.export_btn.grid(row=0, column=2, padx=(5, 0), sticky='w')

        self.copy_failed_btn = ttk.Button(frm, text='Copy Failed', command=self.copy_failed_jobs, state='disabled', width=10)
        self.copy_failed_btn.grid(row=3, column=0, columnspan=6, sticky='ew', pady=5)

        ttk.Label(frm, text='Update Message:').grid(row=4, column=0, sticky='w')
        self.update_message_entry = ttk.Entry(frm)
        self.update_message_entry.grid(row=4, column=1, columnspan=5, sticky='ew', padx=5)

        ttk.Label(frm, text="ARMT No:").grid(row=5, column=0, sticky="w", padx=5)
        self.armt_no_entry = ttk.Entry(frm)
        self.armt_no_entry.grid(row=5, column=1, columnspan=5, sticky='ew', padx=5)

        # Frame for log and structure views
        self.view_frame = ttk.Frame(frm)
        self.view_frame.grid(row=6, column=0, columnspan=6, sticky='nsew', pady=5)
        self.view_frame.grid_rowconfigure(1, weight=1)
        self.view_frame.grid_columnconfigure(0, weight=1)

        # Update Log view
        self.log_label = ttk.Label(self.view_frame, text='Update Log:')
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box = tk.Text(self.view_frame, height=10, state='disabled')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

        # Structure view
        self.structure_label = ttk.Label(self.view_frame, text='Job Structure (XML):')
        self.structure_text = tk.Text(self.view_frame, height=10, state='disabled')

    def parse_and_modify_xml_attributes(self, xml_lines, target_attribute, new_value):
        attributes = {}
        modified_lines = xml_lines.copy()
        pattern = r'(\w+)="([^"]*)"'
        for i, line in enumerate(xml_lines):
            if line.strip().startswith('<Content '):
                matches = re.findall(pattern, line)
                for key, value in matches:
                    attributes[key] = value
                if target_attribute in attributes:
                    attributes[target_attribute] = new_value
                    modified_line = '<Content'
                    for key, value in attributes.items():
                        modified_line += f' {key}="{value}"'
                    closing_part = line[line.find('>'):]
                    modified_line += closing_part
                    modified_lines[i] = modified_line
                    break
        return attributes, modified_lines

    def update_column(self, col_index, value):
        self.column_configs[col_index] = value if value else None
        col_name = self.columns[col_index]
        self.attributes_tree.heading(col_name, text=value if value else col_name)
        self.log(f"Column {col_index} set to {value}")

    def update_dropdowns(self):
        for i, (var, menu) in enumerate(self.dropdowns, 1):
            menu['menu'].delete(0, 'end')
            for option in self.all_options:
                menu['menu'].add_command(label=option, command=tk._setit(var, option, lambda val, idx=i: self.update_column(idx, val)))

    def load_config(self, config):
        if 'ARMT_NO' in config:
            self.armt_no_entry.insert(0, config['ARMT_NO'])
        if 'COLUMN_CONFIGS' in config:
            for i, attr in enumerate(config['COLUMN_CONFIGS'][1:], 1):
                if attr in self.all_options:
                    self.dropdowns[i-1][0].set(attr)
                    self.update_column(i, attr)

    def save_config(self):
        return {
            'ARMT_NO': self.armt_no_entry.get(),
            'COLUMN_CONFIGS': self.column_configs
        }

    def show_paste_menu(self, event):
        self.clicked_column = self.attributes_tree.identify_column(event.x)[1:]
        self.paste_menu.tk_popup(event.x_root, event.y_root)

    def paste_from_clipboard(self, start_column=None):
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        columns = self.attributes_tree['columns']
        if start_column is None:
            if hasattr(self, 'clicked_column') and self.clicked_column:
                col_index = int(self.clicked_column) - 1
                start_column = columns[col_index]
            else:
                start_column = 'jobname'
        try:
            start_col_index = columns.index(start_column)
        except ValueError:
            messagebox.showerror("Error", f"Invalid column: {start_column}")
            self.undo_stack.pop()
            return
        try:
            clipboard_data = self.parent.clipboard_get()
            lines = clipboard_data.strip().splitlines()
            if not lines:
                messagebox.showerror("Error", "Clipboard is empty.")
                self.undo_stack.pop()
                return
            existing_rows = list(self.attributes_tree.get_children())
            max_columns = len(columns) - start_col_index
            for i, line in enumerate(lines):
                clipboard_values = line.split('\t') if '\t' in line else [line]
                clipboard_values = clipboard_values[:max_columns]
                while len(clipboard_values) < max_columns:
                    clipboard_values.append('')
                if i < len(existing_rows):
                    row_id = existing_rows[i]
                    current_values = list(self.attributes_tree.item(row_id)['values'])
                    for j, value in enumerate(clipboard_values):
                        current_values[start_col_index + j] = value.strip()
                    self.attributes_tree.item(row_id, values=current_values)
                else:
                    new_values = [''] * len(columns)
                    for j, value in enumerate(clipboard_values):
                        new_values[start_col_index + j] = value.strip()
                    self.attributes_tree.insert('', 'end', values=new_values)
            self.log(f"Pasted {len(lines)} rows starting from {start_column} column.")
            self.job_structures.clear()
            self.structure_text.config(state='normal')
            self.structure_text.delete(1.0, tk.END)
            self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
            self.structure_text.config(state='disabled')
        except tk.TclError:
            messagebox.showerror("Error", "Clipboard contains invalid data.")
            self.undo_stack.pop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste data: {str(e)}")
            self.undo_stack.pop()

    def on_double_click(self, event):
        row_id = self.attributes_tree.identify_row(event.y)
        if not row_id:
            return
        column = self.attributes_tree.identify_column(event.x)
        column_num = int(column[1:]) - 1
        column_name = self.attributes_tree['columns'][column_num]
        x, y, width, height = self.attributes_tree.bbox(row_id, column)
        entry = ttk.Entry(self.parent)
        entry.place(x=x, y=y, width=width, height=height)
        current_value = self.attributes_tree.item(row_id)['values'][column_num]
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        entry.bind("<Return>", lambda e: self.save_edit(entry, row_id, column_name))
        entry.bind("<FocusOut>", lambda e: self.save_edit(entry, row_id, column_name))
        entry.bind("<Escape>", lambda e: entry.destroy())

    def save_edit(self, entry, row_id, column_name):
        new_value = entry.get().strip()
        self.undo_stack.append(self.save_state())
        self.redo_stack.clear()
        self.attributes_tree.set(row_id, column_name, new_value)
        self.structure_text.config(state='normal')
        self.structure_text.delete(1.0, tk.END)
        self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
        self.structure_text.config(state='disabled')
        entry.destroy()

    def save_state(self):
        return [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]

    def restore_state(self, state):
        self.attributes_tree.delete(*self.attributes_tree.get_children())
        for values in state:
            self.attributes_tree.insert('', 'end', values=values)

    def undo(self, event=None):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def redo(self, event=None):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append(self.save_state())
            self.restore_state(state)
        return 'break'

    def select_all(self, event):
        self.attributes_tree.selection_set(self.attributes_tree.get_children())
        return 'break'

    def delete_selected(self, event):
        selected = self.attributes_tree.selection()
        if selected:
            self.undo_stack.append(self.save_state())
            self.redo_stack.clear()
            for item in selected:
                self.attributes_tree.delete(item)
            current_jobs = {self.attributes_tree.item(child)['values'][0].strip() for child in self.attributes_tree.get_children()}
            jobs_to_remove = set(self.job_structures.keys()) - current_jobs
            for job in jobs_to_remove:
                self.job_structures.pop(job, None)
            self.structure_text.config(state='normal')
            self.structure_text.delete(1.0, tk.END)
            self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
            self.structure_text.config(state='disabled')
        return 'break'

    def show_log_view(self):
        self.structure_label.grid_remove()
        self.structure_text.grid_remove()
        self.log_label.grid(row=0, column=0, sticky='nw')
        self.log_box.grid(row=1, column=0, sticky='nsew', padx=5)

    def show_structure_view(self):
        self.log_label.grid_remove()
        self.log_box.grid_remove()
        self.structure_label.grid(row=0, column=0, sticky='nw')
        self.structure_text.grid(row=1, column=0, sticky='nsew', padx=5)

    def start_update(self):
        self.update_btn.config(state='disabled')
        self.job_structures.clear()
        self.structure_text.config(state='normal')
        self.structure_text.delete(1.0, tk.END)
        self.structure_text.config(state='disabled')
        self.show_log_view()
        threading.Thread(target=self.execute_update, daemon=True).start()

    def execute_update(self):
        self.failed_jobs = []
        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No data to update.")
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False, cert="RB RootCA RSA G01-pem.cer")
        except Exception as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
            self.parent.after(0, lambda: self.update_btn.config(state='normal'))
            return

        update_message = self.update_message_entry.get().strip()
        armt_no = self.armt_no_entry.get().strip()
        current_date = date.today().strftime("%d-%b-%Y")
        name = self.entries.get('NAME', tk.StringVar(value='Unknown')).get().strip()
        doku_entry = f"{armt_no}, {name}, {current_date}, {update_message}" if update_message else ""

        for row in rows:
            object_name = str(row[0]).strip()
            if not object_name:
                self.log("Skipping row with empty jobname")
                continue

            updates = {}
            for col_idx in range(1, 6):
                attr_name = self.column_configs[col_idx]
                if attr_name and attr_name in self.all_options[1:]:
                    attr_value = str(row[col_idx]).strip()
                    if attr_value:
                        updates[attr_name] = attr_value

            # if not updates:
            #     self.log(f"No attributes to update for {object_name}")
            #     continue

            self.log(f"Processing {object_name}")
            try:
                resp = automic.getObjects(client_id=cid, object_name=object_name)
                if resp.status != 200:
                    self.log(f"Failed to fetch {object_name}: {resp.status}")
                    self.failed_jobs.append(object_name)
                    continue
                if 'data' not in resp.response:
                    self.log(f"Error: Object {object_name} response missing 'data' key")
                    self.failed_jobs.append(object_name)
                    continue

                updating_obj = None
                if 'jobs' in resp.response['data']:
                    updating_obj = resp.response["data"]["jobs"]
                elif 'jobp' in resp.response['data']:
                    updating_obj = resp.response["data"]["jobp"]
                else:
                    self.log(f"Error: Object {object_name} not a job or jobp")
                    self.failed_jobs.append(object_name)
                    continue

                updated = False
                doku_found = False
                if cid == 1111:
                    # _BSH update logic (from AttributeUpdaterApp)
                    for doc in updating_obj.get('documentation', []):
                        if '_BSH' in doc:
                            bsh_list = doc['_BSH']
                            if not isinstance(bsh_list, list) or not bsh_list:
                                self.log(f"Error: Invalid _BSH data for {object_name}")
                                self.failed_jobs.append(object_name)
                                continue
                            content_str = None
                            content_index = None
                            for i, item in enumerate(bsh_list):
                                if item.strip().startswith('<Content'):
                                    content_str = item
                                    content_index = i
                                    break
                            if content_str is None:
                                self.log(f"Error: No <Content> element found in _BSH for {object_name}")
                                self.failed_jobs.append(object_name)
                                continue
                            if not content_str.strip():
                                self.log(f"Error: Empty content_str for {object_name}")
                                self.failed_jobs.append(object_name)
                                continue
                            try:
                                content_element = ET.fromstring(content_str)
                                for attr_name, new_value in updates.items():
                                    content_element.set(attr_name, str(new_value.strip()))
                                    self.log(f"Updated {attr_name} to {new_value} for {object_name}")
                                updated_content_str = ET.tostring(content_element, encoding='unicode')
                                bsh_list[content_index] = updated_content_str
                                updated = True
                            except ET.ParseError as e:
                                self.log(f"Error parsing XML for {object_name}: {str(e)}")
                                self.failed_jobs.append(object_name)
                                continue
                else:
                    # _STRUKTUR update logic (from JobpStrukturUpdater)
                    for doc in updating_obj.get('documentation', []):
                        if '_STRUKTUR' in doc:
                            xml_lines = doc['_STRUKTUR']
                            for attr_name, attr_value in updates.items():
                                _, modified_xml_lines = self.parse_and_modify_xml_attributes(xml_lines, attr_name, attr_value)
                                if modified_xml_lines != xml_lines:
                                    doc['_STRUKTUR'] = modified_xml_lines
                                    self.log(f"Updated {attr_name} to {attr_value} for {object_name}")
                                    updated = True
                # Append documentation
                if update_message:
                    
                    for doc in updating_obj.get('documentation', []):
                        if cid ==1111 and 'Doku' in doc:
                            doku_list = doc['Doku']
                            if isinstance(doku_list, list):
                                doku_list.append(doku_entry)
                                break
                        elif '_STRUKTUR' in doc:
                                struktur = doc['_STRUKTUR']
                                for i, line in enumerate(struktur):
                                    if '</HINTS_CHARACTERISTICS>' in line:
                                        # Split before the closing tag
                                        content, tag = line.split('</HINTS_CHARACTERISTICS>')
                                        # print(content)
                                        # Append , and 'your_defined_string'
                                        new_content = f'{content.strip()}' 
                                        # Rebuild the line
                                        struktur[i] = new_content
                                        new_line =  f'{doku_entry}</HINTS_CHARACTERISTICS>{tag}'
                                        struktur.insert(i+1,new_line)
                                        break      
                if updated or update_message:
                    resp_update = automic.postObjects(client_id=cid, body=resp.response, query="overwrite_existing_objects=true")
                    if resp_update.status is None:
                        self.log(f"Successfully updated {object_name}")
                    else:
                        self.log(f"Failed to update {object_name}: {resp_update.status}")
                        self.failed_jobs.append(object_name)
                else:
                    self.log(f"No updates applied for {object_name}")

            except Exception as e:
                self.log(f"Error updating {object_name}: {str(e)}")
                self.failed_jobs.append(object_name)

        self.log("All updates applied.")
        self.parent.after(0, lambda: self.update_btn.config(state='normal'))
        self.parent.after(0, lambda: self.copy_failed_btn.config(state='normal' if self.failed_jobs else 'disabled'))

    def start_fetch_jobs(self):
        self.see_jobs_btn.config(state='disabled')
        self.show_structure_view()
        threading.Thread(target=self.fetch_jobs, daemon=True).start()

    def fetch_jobs(self):
        self.job_structures.clear()
        self.job_attributes.clear()
        dynamic_attributes = set()

        rows = [self.attributes_tree.item(child)['values'] for child in self.attributes_tree.get_children()]
        rows = [row for row in rows if any(val.strip() for val in row)]
        if not rows:
            self.log("No jobs to fetch.")
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_structure_text_for_selection)
            return

        env = self.env_var.get().strip()
        try:
            cid = int(self.client_var.get().strip())
        except ValueError:
            self.log("Error: Invalid Client ID")
            self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_structure_text_for_selection)
            return

        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        try:
            automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False, cert="RB RootCA RSA G01-pem.cer")
        except Exception as e:
            self.log(f"Authentication failed: {str(e)}")
            self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials"))
            self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
            self.parent.after(0, self.update_structure_text_for_selection)
            return

        def fetch_single_job(jobname):
            if not jobname:
                return None, None, None
            try:
                self.log(f"Fetching {jobname}")
                resp = automic.getObjects(client_id=cid, object_name=jobname)
                if resp.status != 200:
                    self.log(f"Failed to fetch {jobname}: {resp.status}")
                    return jobname, None, None
                if 'data' not in resp.response:
                    self.log(f"Error: Object {jobname} response missing 'data' key")
                    return jobname, None, None
                updating_obj = None
                if 'jobs' in resp.response['data']:
                    updating_obj = resp.response["data"]["jobs"]
                elif 'jobp' in resp.response['data']:
                    updating_obj = resp.response["data"]["jobp"]
                else:
                    self.log(f"Error: Object {jobname} not a job or jobp")
                    return jobname, None, None
                structure_lines = None
                job_attributes = {}
                for doc in updating_obj.get('documentation', []):
                    doc_key = '_BSH' if cid == 1111 else '_STRUKTUR'
                    if doc_key in doc:
                        structure_lines = doc[doc_key]
                        for line in structure_lines:
                            if line.strip().startswith('<Content '):
                                matches = re.findall(r'(\w+)="([^"]*)"', line)
                                for key, value in matches:
                                    job_attributes[key] = value
                                    dynamic_attributes.add(key)
                                    self.log(f"Extracted {key}={value} for job {jobname}")
                        break
                return jobname, structure_lines, job_attributes
            except Exception as e:
                self.log(f"Error fetching {jobname}: {str(e)}")
                return jobname, None, None

        jobnames = [str(row[0]).strip() for row in rows if str(row[0]).strip()]
        max_workers = min(len(jobnames), 10)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_jobname = {executor.submit(fetch_single_job, jobname): jobname for jobname in jobnames}
            for future in as_completed(future_to_jobname):
                jobname, structure_lines, job_attrs = future.result()
                if structure_lines is not None:
                    self.job_structures[jobname] = structure_lines
                if job_attrs is not None:
                    self.job_attributes[jobname] = job_attrs

        self.all_options = [''] + sorted(list(dynamic_attributes))
        self.parent.after(0, self.update_dropdowns)
        self.log(f"Updated dropdown options: {self.all_options}")
        self.log("All jobs fetched.")
        self.parent.after(0, lambda: self.see_jobs_btn.config(state='normal'))
        self.parent.after(0, self.update_structure_text_for_selection)

    def export_to_excel(self):
        if not self.job_attributes:
            messagebox.showinfo("Export Info", "No job attribute data to export. Click 'See Jobp' to fetch data.")
            self.log("No job attribute data to export.")
            return
        try:
            job_names = sorted(self.job_attributes.keys())
            all_attrs = sorted(set(attr for job_attrs in self.job_attributes.values() for attr in job_attrs))
            data = []
            for job in job_names:
                row = {'Jobp Name': job}
                for attr in all_attrs:
                    row[attr] = self.job_attributes[job].get(attr, '')
                data.append(row)
            df = pd.DataFrame(data, columns=['Jobp Name'] + all_attrs)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            filename = f"job_attributes_{timestamp}.xlsx"
            df.to_excel(filename, index=False)
            self.log(f"Exported job attributes to {filename}")
            messagebox.showinfo("Export Success", f"Successfully exported to {filename}")
        except Exception as e:
            self.log(f"Error exporting to Excel: {str(e)}")
            messagebox.showerror("Export Error", f"Failed to export to Excel: {str(e)}")

    def update_structure_text_for_selection(self):
        selected = self.attributes_tree.selection()
        self.structure_text.config(state='normal')
        self.structure_text.delete(1.0, tk.END)
        if not selected:
            self.structure_text.config(state='disabled')
            return
        row_id = selected[0]
        jobname = self.attributes_tree.item(row_id)['values'][0].strip()
        if jobname in self.job_structures:
            structure_lines = self.job_structures[jobname]
            attributes = []
            pattern = r'(\w+)="([^"]*)"'
            for line in structure_lines:
                if line.strip().startswith('<Content '):
                    matches = re.findall(pattern, line)
                    attributes.extend([f"{key}={value}" for key, value in matches])
            if attributes:
                self.structure_text.insert(tk.END, "\n".join(attributes))
            else:
                self.structure_text.insert(tk.END, "No attributes found in structure data.")
        else:
            self.structure_text.insert(tk.END, "No structure data available. Click 'See Jobp' to fetch.")
        self.structure_text.config(state='disabled')

    def on_treeview_select(self, event):
        self.update_structure_text_for_selection()

    def copy_failed_jobs(self):
        if self.failed_jobs:
            failed_list = "\n".join(self.failed_jobs)
            self.parent.clipboard_clear()
            self.parent.clipboard_append(failed_list)
            self.log("Copied failed jobs to clipboard.")
        else:
            messagebox.showinfo("Info", "No failed jobs to copy.")

    def log(self, msg):
        self.parent.after(0, lambda: self._log(msg))

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
import sv_ttk
class AutomicToolsApp:
    CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.automic_tools.json')
    
    ENV_OPTIONS = ['eup4', 'eup6', 'eup7']
    CLIENT_MAP = {
        'eup4': ['1100'],
        'eup6': ['1001', '1111'],
        'eup7': ['1101', '1301', '1401', '7101','7001']
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Automic Tools")
        self.root.geometry("900x650")
        self.root.minsize(800, 600)

        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, 'your_icon.ico')
        else:
            icon_path = 'your_icon.ico'
        self.root.iconbitmap(icon_path)

        style = ttk.Style()
        style.theme_use('clam')


        # --- Color Palette ---
        accent        = '#94A3B8'   # Cool Gray
        accent_hover  = '#6B7280'   # Slightly darker on hover
        background    = '#F7F9FB'   # Light neutral background
        text_color    = '#333333'
        entry_bg      = '#FFFFFF'
        border_color  = '#D1D5DB'   # Subtle border for fields

        # --- Font ---
        font_name = 'Montserrat'
        font = (font_name, 10)

        # --- Root Window ---
        self.root.configure(bg=background)

        # --- Frame & Label ---
        style.configure('TFrame', background=background)
        style.configure('TLabel', background=background, foreground=text_color, font=font)

        # --- Entry ---
        style.configure('TEntry',
            font=font,
            fieldbackground=entry_bg,
            background=entry_bg,
            foreground=text_color,
            bordercolor=border_color,
            relief='flat',
            padding=5
        )

        # --- Combobox ---
        style.configure('TCombobox',
            font=font,
            fieldbackground=entry_bg,
            background=entry_bg,
            foreground=text_color,
            padding=4
        )

        # --- Buttons ---
        style.configure('TButton',
            font=(font_name, 10, 'bold'),
            background=accent,
            foreground='white',
            padding=(10, 6),
            borderwidth=0,
            relief='flat'
        )
        style.map('TButton',
            background=[('active', accent_hover)],
            relief=[('pressed', 'flat')]
        )

        # --- Treeview ---
        style.configure('Treeview.Heading',
            font=(font_name, 10, 'bold'),
            background=accent,
            foreground='white'
        )
        style.configure('Treeview',
            font=font,
            rowheight=28,
            background='white',
            fieldbackground='white',
            foreground=text_color,
            bordercolor=border_color
        )

        # --- Scrollbar ---
        style.configure('Vertical.TScrollbar',
            gripcount=0,
            width=8,
            troughcolor=background,
            background=accent,
            bordercolor=border_color
        )

        # --- Notebook (Tabs) ---
        style.configure('TNotebook', background=background, borderwidth=0)
        style.configure('TNotebook.Tab',
            font=font,
            padding=(10, 6),
            background=entry_bg,
            foreground=text_color
        )
        style.map('TNotebook.Tab',
            background=[('selected', accent)],
            foreground=[('selected', 'white')]
        )

        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill='x', padx=10, pady=8)
        top_frame.grid_columnconfigure((1, 3), weight=1)

        ttk.Label(top_frame, text="Environment:").grid(row=0, column=0, sticky="w", padx=5)
        self.env_var = tk.StringVar(value="eup4")
        env_cb = ttk.Combobox(top_frame, textvariable=self.env_var, values=self.ENV_OPTIONS, state='readonly', width=6)
        env_cb.grid(row=0, column=1, padx=5, sticky="w")
        env_cb.bind('<<ComboboxSelected>>', lambda e: self.update_client_options())

        ttk.Label(top_frame, text="Client ID:").grid(row=0, column=2, sticky="w", padx=5)
        self.client_var = tk.StringVar()
        self.client_cb = ttk.Combobox(top_frame, textvariable=self.client_var, state='readonly')
        self.client_cb.grid(row=0, column=3, padx=5, sticky="ew")

        ttk.Label(top_frame, text="User ID:").grid(row=1, column=0, sticky="w", padx=5)
        self.entries = {'USERID': ttk.Entry(top_frame)}
        self.entries['USERID'].grid(row=1, column=1, padx=5, sticky="ew")

        ttk.Label(top_frame, text="Password:").grid(row=1, column=2, sticky="w", padx=5)
        self.entries['PASSWORD'] = ttk.Entry(top_frame, show="*")
        self.entries['PASSWORD'].grid(row=1, column=3, padx=5, sticky="ew")
        # Add the new "Name:" field
        ttk.Label(top_frame, text="Name:").grid(row=2, column=2, sticky="w", padx=5)
        self.entries['NAME'] = ttk.Entry(top_frame)
        self.entries['NAME'].grid(row=2, column=3, padx=5, sticky="ew")
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)

        self.job_creator_frame = ttk.Frame(self.notebook)
        self.usage_viewer_frame = ttk.Frame(self.notebook)
        self.attribute_updater_frame = ttk.Frame(self.notebook)
        self.child_viewer_frame = ttk.Frame(self.notebook)
        self.find_bulk_frame = ttk.Frame(self.notebook)
        self.jobs_updater_frame = ttk.Frame(self.notebook)
        self.jobp_updater_frame = ttk.Frame(self.notebook)
        self.vara_updater_frame = ttk.Frame(self.notebook)
        self.firstrun_frame = ttk.Frame(self.notebook)
        self.pvl_checker_frame = ttk.Frame(self.notebook)
        self.jobp_struktur_updater_frame = ttk.Frame(self.notebook)
        self.unified_updater_frame = ttk.Frame(self.notebook)


        self.notebook.add(self.job_creator_frame, text='Job Creator')
        self.notebook.add(self.usage_viewer_frame, text='Usage Viewer')
        self.notebook.add(self.child_viewer_frame, text='Child Viewer')
        self.notebook.add(self.find_bulk_frame, text='Bulk Finder')
        self.notebook.add(self.jobs_updater_frame, text='JOBS Updater')
        self.notebook.add(self.jobp_updater_frame, text='JOBP Updater')
        self.notebook.add(self.vara_updater_frame, text='WARTUNG Updater')
        self.notebook.add(self.unified_updater_frame, text='ATTRIBUTES UPDATER')

        # self.notebook.add(self.attribute_updater_frame, text='BSH A/U')
        self.notebook.add(self.firstrun_frame, text='FIRSTRUN & MOVE')
        self.notebook.add(self.pvl_checker_frame, text='PVL CHECKER')
        # self.notebook.add(self.jobp_struktur_updater_frame, text='JOBP STRUKTUR UPDATER')


        self.job_creator = JobCreatorApp(self.job_creator_frame, self.env_var, self.client_var, self.entries, self.CLIENT_MAP)
        self.usage_viewer = AutomicApp(self.usage_viewer_frame, self.env_var, self.client_var, self.entries)
        # self.attribute_updater = AttributeUpdaterApp(self.attribute_updater_frame, self.env_var, self.client_var, self.entries)
        self.child_viewer = ChildViewer(self.child_viewer_frame, self.env_var, self.client_var, self.entries)
        self.bulk_finder = FindBulk(self.find_bulk_frame, self.env_var, self.client_var, self.entries)
        self.jobs_updater = JobsUpdater(self.jobs_updater_frame, self.env_var, self.client_var, self.entries)
        self.jobp_updater = JobpUpdater(self.jobp_updater_frame, self.env_var, self.client_var, self.entries)
        self.vara_updater = VaraUpdaterApp(self.vara_updater_frame, self.env_var, self.client_var, self.entries)
        self.firstrun_checker = FirstrunChecker(self.firstrun_frame, self.env_var, self.client_var, self.entries)
        self.pvl_checker = PvlChecker(self.pvl_checker_frame, self.env_var, self.client_var, self.entries)
        self.jobp_updater = JobpStrukturUpdater(self.jobp_struktur_updater_frame, self.env_var, self.client_var, self.entries)
        self.unified_updater = UnifiedUpdaterApp(self.unified_updater_frame, self.env_var, self.client_var, self.entries)

        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


    def load_config(self):
            try:
                with open(self.CONFIG_PATH, 'r') as f:
                    cfg = json.load(f)
                global_cfg = cfg.get('global', {})
                if global_cfg.get('ENV') in self.ENV_OPTIONS:
                    self.env_var.set(global_cfg['ENV'])
                    self.update_client_options()
                if global_cfg.get('CLIENT_ID') in self.CLIENT_MAP.get(self.env_var.get(), []):
                    self.client_var.set(global_cfg['CLIENT_ID'])
                for key in ['USERID', 'NAME']:
                    if global_cfg.get(key):
                        self.entries[key].insert(0, global_cfg[key])
                # Load password from keyring
                password = keyring.get_password("AutomicTools", "user_password")
                if password:
                    self.entries['PASSWORD'].insert(0, password)
            except FileNotFoundError:
                pass

    def save_all_configs(self):
        global_cfg = {
            'ENV': self.env_var.get(),
            'CLIENT_ID': self.client_var.get(),
            'USERID': self.entries['USERID'].get(),
            'NAME': self.entries['NAME'].get()
        }
        # Save password to keyring
        keyring.set_password("AutomicTools", "user_password", self.entries['PASSWORD'].get())
        cfg = {
            'global': global_cfg,
        }
        with open(self.CONFIG_PATH, 'w') as f:
            json.dump(cfg, f)

    def on_closing(self):
        """Save configurations and close the application."""
        self.save_all_configs()
        self.root.destroy()

    def update_client_options(self):
        opts = self.CLIENT_MAP.get(self.env_var.get(), [])
        self.client_cb['values'] = opts
        if opts and self.client_var.get() not in opts:
            self.client_var.set(opts[0])

if __name__ == '__main__':
    root = tk.Tk()
    app = AutomicToolsApp(root)
    root.mainloop()