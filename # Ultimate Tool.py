import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, font, filedialog
import threading
import base64
import automic_rest as automic
import copy
import re
import os
import json
import sys
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests  # Added for handling HTTP errors

#----------------------------
# Helpers for JobCreatorApp
#----------------------------

def sanitize_string(s):
    s = re.sub(r'^[^0-9A-Za-z_]+', '', s)
    return re.sub(r'[^0-9A-Za-z_]', '_', s)

def parse_flexible_pairs(data):
    headers_by_length = {
        1: ["jobp"],
        3: ["jobname", "program", "variant"],
        4: ["jobname", "program", "variant", "user"],
        5: ["jobname", "program", "variant", "user", "language"],
        6: ["jobname", "program", "variant", "user", "language", "extra"]
    }
    parsed = []
    for line in data.strip().splitlines():
        line = line.strip()
        if not line:  # Skip blank lines
            continue
        parts = line.split()
        n = len(parts)
        if n == 2:
            program, variant = parts
            jobname = f"C_{sanitize_string(program)}_{sanitize_string(variant)}"
            parsed.append({"jobname": jobname, "program": program, "variant": variant, "isBSH": True})
        elif n in headers_by_length:
            parsed.append(dict(zip(headers_by_length[n], parts)))
        else:
            parsed.append({f"col_{i+1}": v for i, v in enumerate(parts)})
    print(parsed)
    return parsed

def extract_default_login(template_jobs):
    for proc in template_jobs.get('scripts', []):
        if 'process' in proc:
            for line in proc['process']:
                m = re.match(r":PUT_ATT\s+LOGIN\s*=\s*'([^']+)'", line)
                if m:
                    return m.group(1)
    return 'LOGIN_R3_060_SY-BATCH-PM'

#----------------------------
# JobCreatorApp
#----------------------------
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, font, filedialog
import threading
import base64
import automic_rest as automic
import copy
import re
import os
import json
import sys
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
import uuid

# ... (Existing helper functions: sanitize_string, parse_flexible_pairs, extract_default_login remain unchanged)

class JobCreatorApp:
    CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.automic_tools.json')

    def __init__(self, parent, env_var, client_var, entries, client_map):
        self.parent = parent
        self.env_var = env_var
        self.client_var = client_var
        self.entries = entries
        self.client_map = client_map
        self.jobs_list = []  # Store created job names
        self.jobps_list = []  # Store created job plan names
        self.load_config()
        self.build_ui()
        self.populate_fields()

    def load_config(self):
        try:
            with open(self.CONFIG_PATH, 'r') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {}

    def save_config(self):
        data = {
            'ENV': self.env_var.get(),
            'CLIENT_ID': self.client_var.get(),
            'USERID': self.entries['USERID'].get(),
            'PASSWORD': base64.b64encode(self.entries['PASSWORD'].get().encode()).decode(),
            'ARMT_NO': self.entries['ARMT_NO'].get(),
            'template_job_armt': self.template_job_armt.get(),
            'template_joplan_armt': self.template_joplan_armt.get(),
            'PAIRS_DATA': self.pairs_text.get('1.0', 'end'),
            'CREATE_MAIN': self.create_main_var.get(),
            'JOBP_MAIN_NAME': self.jobp_main_entry.get()
        }
        with open(self.CONFIG_PATH, 'w') as f:
            json.dump(data, f)

    def populate_fields(self):
        cfg = self.config
        for key in ['ARMT_NO']:
            if cfg.get(key):
                self.entries[key].insert(0, cfg[key])
        for fld in ['template_job_armt', 'template_joplan_armt']:
            if cfg.get(fld): getattr(self, fld).insert(0, cfg[fld])
        if cfg.get('PAIRS_DATA'):
            self.pairs_text.insert('1.0', cfg['PAIRS_DATA'])
        self.create_main_var.set(cfg.get('CREATE_MAIN', False))
        self.toggle_main_fields()
        self.jobp_main_entry.insert(0, cfg.get('JOBP_MAIN_NAME', ''))

    def build_ui(self):
        frm = ttk.Frame(self.parent, padding=15)
        frm.pack(fill='both', expand=True)
        # ARMT No.
        ttk.Label(frm, text='ARMT No.:').grid(row=0, column=0, sticky='w')
        self.entries['ARMT_NO'] = ttk.Entry(frm)
        self.entries['ARMT_NO'].grid(row=0, column=1, sticky='ew', padx=5)
        # Templates
        ttk.Label(frm, text='Jobplan Template:').grid(row=1, column=0, sticky='w')
        self.template_joplan_armt = ttk.Entry(frm)
        self.template_joplan_armt.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(frm, text='Jobs Template:').grid(row=2, column=0, sticky='w')
        self.template_job_armt = ttk.Entry(frm)
        self.template_job_armt.grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        # Main jobplan options
        self.create_main_var = tk.BooleanVar()
        chk = ttk.Checkbutton(frm, text='Create Main Jobplan', variable=self.create_main_var, command=self.toggle_main_fields)
        chk.grid(row=0, column=2, sticky='w')
        self.is_predecessor_var = tk.BooleanVar()
        self.predecessor_chk = ttk.Checkbutton(frm, text='Use Sequential Predecessors', variable=self.is_predecessor_var)
        self.predecessor_chk.grid(row=1, column=2, columnspan=2, sticky='w')
        self.main_label = ttk.Label(frm, text='Main JOBP Name:')
        self.main_entry = ttk.Entry(frm)
        self.jobp_main_entry = self.main_entry
        self.main_label.grid(row=2, column=2, sticky='w')
        self.main_entry.grid(row=2, column=3, sticky='ew', padx=5)
        # Pairs Data
        ttk.Label(frm, text='Program/Variant Pairs:').grid(row=3, column=0, sticky='nw', pady=(10, 2))
        self.pairs_text = scrolledtext.ScrolledText(frm, height=6, undo=True, autoseparators=True, maxundo=-1)
        self.pairs_text.bind("<Control-y>", lambda e: self.pairs_text.event_generate("<<Redo>>"))
        self.pairs_text.bind("<Control-Y>", lambda e: self.pairs_text.event_generate("<<Redo>>"))
        self.pairs_text.grid(row=3, column=1, columnspan=3, sticky='ew', padx=5)
        # Run Button
        self.run_btn = ttk.Button(frm, text='Create Jobs', command=self.start)
        self.run_btn.grid(row=4, column=0, columnspan=4, pady=12)
        # Output Log
        ttk.Label(frm, text='Output:').grid(row=5, column=0, sticky='nw')
        self.log_box = scrolledtext.ScrolledText(frm, height=10, state='disabled')
        self.log_box.grid(row=5, column=1, columnspan=3, sticky='ew', padx=5)
        # New Buttons for Copying Lists
        self.copy_jobs_btn = ttk.Button(frm, text='Copy JOBS List', command=self.copy_jobs_list)
        self.copy_jobs_btn.grid(row=6, column=1, sticky='w', padx=5, pady=5)
        self.copy_jobps_btn = ttk.Button(frm, text='Copy JOBP List', command=self.copy_jobps_list)
        self.copy_jobps_btn.grid(row=6, column=2, sticky='w', padx=5, pady=5)
        frm.columnconfigure((1, 3), weight=1)
        self.toggle_main_fields()

    def toggle_main_fields(self):
        if self.create_main_var.get():
            self.main_label.grid()
            self.main_entry.grid()
            self.predecessor_chk.grid()
        else:
            self.main_label.grid_remove()
            self.main_entry.grid_remove()
            self.predecessor_chk.grid_remove()

    def copy_jobs_list(self):
        """Copy the list of created job names to the clipboard."""
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
        """Copy the list of created job plan names to the clipboard."""
        if not self.jobps_list:
            self.log("No job plans list available to copy.")
            messagebox.showinfo("Info", "No job plans list available to copy.")
            return
        jobps_text = "\n".join(self.jobps_list)
        self.parent.clipboard_clear()
        self.parent.clipboard_append(jobps_text)
        self.parent.update()
        self.log("Copied JOBP list to clipboard.")

    def update_client_options(self):
        opts = self.client_map.get(self.env_var.get(), [])
        self.client_cb['values'] = opts
        if opts and self.client_var.get() not in opts: self.client_var.set(opts[0])

    def log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
        self.parent.update_idletasks()

    def start(self):
        self.run_btn.config(state='disabled')
        threading.Thread(target=self.execute, daemon=True).start()

    def execute(self):
        try:
            self.jobs_list = []  # Reset jobs list
            self.jobps_list = []  # Reset job plans list
            env = self.env_var.get().strip()
            try:
                cid = int(self.client_var.get().strip())
            except ValueError:
                self.parent.after(0, lambda: self.log("Error: Invalid Client ID"))
                self.parent.after(0, lambda: messagebox.showerror("Error", "Invalid Client ID. Please enter a numeric value."))
                return

            user = self.entries['USERID'].get().strip()
            pwd = self.entries['PASSWORD'].get().strip()
            armt = self.entries['ARMT_NO'].get().strip()
            api_url = f'https://rb-{env}-api.bosch.com'
            t_job = self.template_job_armt.get().strip()
            t_joplan = self.template_joplan_armt.get().strip()
            raw = self.pairs_text.get('1.0', 'end')
            create_main = self.create_main_var.get()
            main_name = self.jobp_main_entry.get().strip()

            if not user or not pwd:
                self.parent.after(0, lambda: self.log("Error: User ID and Password are required"))
                self.parent.after(0, lambda: messagebox.showerror("Error", "Please provide both User ID and Password"))
                return

            self.save_config()

            # Authenticate
            auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
            try:
                automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
            except requests.exceptions.HTTPError as e:
                self.parent.after(0, lambda: self.log(f"Authentication failed: {str(e)}"))
                self.parent.after(0, lambda: messagebox.showerror("Authentication Error", f"Failed to authenticate: {str(e)}. Please check your credentials."))
                return

            # Fetch jobplan template
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
                        base_jobp = tmpl_jobp['general_attributes']['name'][:23]
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error fetching jobplan {t_joplan}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to fetch jobplan {t_joplan}: {str(e)}"))
                    return
                except Exception as e:
                    self.parent.after(0, lambda: self.log(f"Unexpected error fetching jobplan {t_joplan}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("Error", f"Unexpected error fetching jobplan {t_joplan}: {str(e)}"))
                    return

            # Fetch job template
            if t_job:
                self.parent.after(0, lambda: self.log(f"Fetching job {t_job}"))
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
                        base_jobs = tmpl_jobs['general_attributes']['name'][:15]
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error fetching job {t_job}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to fetch job {t_job}: {str(e)}"))
                    return
                except Exception as e:
                    self.parent.after(0, lambda: self.log(f"Unexpected error fetching job {t_job}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("Error", f"Unexpected error fetching job {t_job}: {str(e)}"))
                    return

                default_login = extract_default_login(tmpl_jobs)
            pairs = parse_flexible_pairs(raw)

            # Create jobplans and jobs
            self.jobps_list = []  # Ensure list is reset
            self.jobs_list = []   # Ensure list is reset
            if pairs[0].get("jobp"):
                t_joplan = pairs[0]["jobp"]
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
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error fetching jobplan {t_joplan}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to fetch jobplan {t_joplan}: {str(e)}"))
                    return
                except Exception as e:
                    self.parent.after(0, lambda: self.log(f"Unexpected error fetching jobplan {t_joplan}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("Error", f"Unexpected error fetching jobplan {t_joplan}: {str(e)}"))
                    return
                for p in pairs:
                    self.jobps_list.append(p['jobp'])
            else:
                for p in pairs:
                    jn = p['jobname']
                    if tmpl_jobp:
                        name_jobp = f"{base_jobp}_{jn}"
                        self.jobps_list.append(name_jobp)
                        njp = copy.deepcopy(tmpl_jobp)
                        njp['general_attributes']['name'] = name_jobp
                        for wf in njp.get('workflow_definitions', []):
                            if wf.get('object_name') == tmpl_jobs['general_attributes']['name']:
                                wf['object_name'] = f"{base_jobs}_{jn}"
                        try:
                            res_p = automic.postObjects(client_id=cid, body={'total':1,'data':{'jobp':njp},'path':f'AUTOMATION_JOBS/{user}/{armt}','client':cid,'hasmore':False})
                            self.parent.after(0, lambda: self.log(f"JOBP: {name_jobp}" if res_p.status is None else f"FAIL JOBP: {name_jobp} ({res_p.status})"))
                        except requests.exceptions.HTTPError as e:
                            self.parent.after(0, lambda: self.log(f"HTTP error creating jobplan {name_jobp}: {str(e)}"))
                            continue
                        except Exception as e:
                            self.parent.after(0, lambda: self.log(f"Unexpected error creating jobplan {name_jobp}: {str(e)}"))
                            continue

                    name_jobs = f"{base_jobs}_{jn}"
                    self.jobs_list.append(name_jobs)
                    login_val = f"LOGIN_R3_060_{p.get('login', default_login)}"
                    if cid == 1111:
                        script = [
                            f":INC BSH_XXXX_INC_MIGRATION_SIMULATION WAIT_TIME = \"<Random number ...>\" ,NOFOUND=IGNORE",
                            f":PUT_ATT JOB_NAME= \"{jn}\"",
                            f":PUT_ATT LOGIN='{login_val}'",
                            f"R3_ACTIVATE_REPORT REPORT='{p['program']}',VARIANT='{p['variant']}',COPIES=1,EXPIR=8,LINE_COUNT=65,LINE_SIZE=80,LAYOUT=X_FORMAT,DATA_SET=LIST1S,TYPE=TEXT"
                        ]
                    else:
                        script = (
                            ([f':PUT_ATT JOB_NAME= "{jn}"'] if not p.get('isBSH') else [])
                            + [f"R3_ACTIVATE_REPORT REPORT='{p['program']}',VARIANT='{p['variant']}'"]
                        )
                    nj = copy.deepcopy(tmpl_jobs)
                    nj['general_attributes']['name'] = name_jobs
                    for proc in nj.get('scripts', []):
                        if 'process' in proc:
                            proc['process'] = script
                    try:
                        res_j = automic.postObjects(client_id=cid, body={'total':1,'data':{'jobs':nj},'path':f'AUTOMATION_JOBS/{user}/{armt}','client':cid,'hasmore':False})
                        self.parent.after(0, lambda: self.log(f"JOBS: {name_jobs}" if res_j.status is None else f"FAIL JOBS: {name_jobs} ({res_j.status})"))
                    except requests.exceptions.HTTPError as e:
                        self.parent.after(0, lambda: self.log(f"HTTP error creating job {name_jobs}: {str(e)}"))
                        continue
                    except Exception as e:
                        self.parent.after(0, lambda: self.log(f"Unexpected error creating job {name_jobs}: {str(e)}"))
                        continue

            # Create main jobplan
            is_predecessor_var = self.is_predecessor_var.get()
            if create_main and main_name and tmpl_jobp:
                self.jobps_list.append(main_name)
                data = tmpl_jobp
                start_node = next(obj for obj in data['workflow_definitions'] if obj['object_type'] == '<START>')
                end_node = next(obj for obj in data['workflow_definitions'] if obj['object_type'] == '<END>')
                new_defs = [start_node]
                line_no = 2
                for jp in self.jobps_list[:-1]:  # Exclude main jobplan
                    new_node = {
                        'line_number': line_no,
                        'object_type': 'JOBP',
                        'object_name': jp,
                        'precondition_error_action': 'H',
                        'predecessors': 1,
                        'active': 1,
                        'mrt_time': '000000',
                        'childflags': '0000000000000000',
                        'rollback_enabled': 1
                    }
                    if is_predecessor_var:
                        new_node['row'] = 1
                        new_node['column'] = line_no
                    else:
                        new_node['row'] = line_no - 1
                        new_node['column'] = 2
                    new_defs.append(new_node)
                    line_no += 1
                end_node['predecessors'] = line_no - 2
                end_node['line_number'] = line_no
                end_node['row'] = 1
                if is_predecessor_var:
                    end_node['column'] = line_no
                else:
                    end_node['column'] = 3
                new_defs.append(end_node)
                data['workflow_definitions'] = new_defs

                def gen_conditions(defs):
                    conds = []
                    for node in defs:
                        ln = node['line_number']
                        if 'predecessors' in node:
                            preds = node.get('predecessors', [])
                            if not is_predecessor_var:
                                if node['object_type'] == '<END>': preds = list(range(2, preds + 2))
                                else: preds = [1]
                            else:
                                preds = [ln - 1]
                            for idx, p in enumerate(preds, 1):
                                conds.append({'workflow_line_number': ln, 'line_number': idx, 'predecessor_line_number': p})
                    return conds

                data['line_conditions'] = gen_conditions(new_defs)
                data['general_attributes']['name'] = main_name
                body = {'total': 1, 'data': {'jobp': data}, 'path': f'AUTOMATION_JOBS/{user}/{armt}', 'client': cid, 'hasmore': False}
                try:
                    resp_main = automic.postObjects(client_id=cid, body=body)
                    self.parent.after(0, lambda: self.log(f"MAIN JOBP: {main_name}" if resp_main.status is None else f"FAIL MAIN JOBP: {main_name} ({resp_main.status})"))
                except requests.exceptions.HTTPError as e:
                    self.parent.after(0, lambda: self.log(f"HTTP error creating main jobplan {main_name}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("HTTP Error", f"Failed to create main jobplan {main_name}: {str(e)}"))
                except Exception as e:
                    self.parent.after(0, lambda: self.log(f"Unexpected error creating main jobplan {main_name}: {str(e)}"))
                    self.parent.after(0, lambda: messagebox.showerror("Error", f"Unexpected error creating main jobplan {main_name}: {str(e)}"))

            self.parent.after(0, lambda: self.log("All done."))

        except Exception as e:
            self.parent.after(0, lambda: self.log(f"Unexpected error: {str(e)}"))
            self.parent.after(0, lambda: messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}"))
        finally:
            self.parent.after(0, lambda: self.run_btn.config(state='normal'))

#----------------------------
# AutomicApp
#----------------------------
class AutomicApp:
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

        self.columns = ("Object Name", "Usage", "Type", "Folder", "Last Modified", "Last Execution")
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        self.tree.bind("<Button-1>", self.on_column_click)

        for col in self.columns:
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

    def on_column_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.tree.identify_column(event.x)
            col_index = int(col_id.replace("#", "")) - 1
            col_name = self.columns[col_index]
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

        def fetch_single_object(obj_name):
            try:
                if self.cancel_batch:
                    return None
                result = automic.usageObject(client_id=int(client_id), object_name=obj_name)
                response_data = result.response
                refs = response_data.get("references", [])
                last_exec = self.get_last_execution(client_id, obj_name)
                color = self.get_object_color(obj_name)
                return obj_name, refs, last_exec, color
            except Exception as e:
                print(f"Error fetching {obj_name}: {e}")
                return obj_name, [], None, None

        def fetch_objects():
            try:
                auth = base64.b64encode(f"{userid}:{password}".encode()).decode()
                url = f"https://rb-{env}-api.bosch.com"
                automic.connection(url=url, auth=auth, noproxy=True, sslverify=False, cert="/path/to/certfile", timeout=60)

                self.parent.after(0, lambda: self.tree.delete(*self.tree.get_children()))
                total_refs_found = 0

                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = {executor.submit(fetch_single_object, obj_name): obj_name for obj_name in object_names}
                    for future in as_completed(futures):
                        if self.cancel_batch:
                            print("Batch fetch cancelled.")
                            break
                        result = future.result()
                        if result is None:
                            continue
                        obj_name, refs, last_exec, color = result

                        def insert_row():
                            nonlocal total_refs_found
                            if not refs:
                                self.tree.insert("", "end", values=(obj_name, "None", "None", "None", "None", last_exec))
                            else:
                                for r in refs:
                                    self.tree.insert("", "end", values=(obj_name, r["name"], r["type"], r["folderpath"], r["lastmodified"][:10], last_exec), tags=(obj_name,))
                                self.tree.tag_configure(obj_name, background=color)
                            total_refs_found += len(refs)
                            self.status.config(text=f"Fetched {total_refs_found} references so far...")

                        self.parent.after(0, insert_row)

                if not self.cancel_batch:
                    self.parent.after(0, lambda: self.status.config(text=f"Done fetching {len(object_names)} objects. {total_refs_found} references total."))
                else:
                    self.parent.after(0, lambda: self.status.config(text="Fetch cancelled."))

            except Exception as e:
                self.parent.after(0, lambda: messagebox.showerror("Error", f"Batch fetch failed:\n{str(e)}"))
            finally:
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
            print(f"Error fetching last execution for {obj_name}: {e}")
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

#----------------------------
# Main Application
#----------------------------
class AutomicToolsApp:
    CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.automic_tools.json')
    ENV_OPTIONS = ['eup4', 'eup6', 'eup7']
    CLIENT_MAP = {
        'eup4': ['1100'],
        'eup6': ['1001', '1111'],
        'eup7': ['1101', '1301', '1401', '7101']
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
        accent = '#0096D6'
        bg = '#F0F2F5'
        font_name = 'Montserrat'
        font = (font_name, 10)
        self.root.configure(bg=bg)
        style.configure('TFrame', background=bg)
        style.configure('TLabel', background=bg, font=font)
        style.configure('TEntry', font=font)
        style.configure('TCombobox', font=font)
        style.configure('TButton', font=(font_name, 10, 'bold'), background=accent, foreground='white', padding=6, relief='flat')
        style.map('TButton', background=[('active', '#007BB5')])
        style.configure('Treeview.Heading', font=(font_name, 10, 'bold'), background=accent, foreground='white')
        style.configure('Treeview', font=font, rowheight=25, background='white', fieldbackground='white')
        style.configure('Vertical.TScrollbar', gripcount=0, width=8)
        style.configure('TNotebook', background=bg)
        style.configure('TNotebook.Tab', font=font)

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

        self.load_config()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)

        self.job_creator_frame = ttk.Frame(self.notebook)
        self.usage_viewer_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.job_creator_frame, text='Job Creator')
        self.notebook.add(self.usage_viewer_frame, text='Usage Viewer')

        self.job_creator = JobCreatorApp(self.job_creator_frame, self.env_var, self.client_var, self.entries, self.CLIENT_MAP)
        self.usage_viewer = AutomicApp(self.usage_viewer_frame, self.env_var, self.client_var, self.entries)

    def load_config(self):
        try:
            with open(self.CONFIG_PATH, 'r') as f:
                cfg = json.load(f)
            if cfg.get('ENV') in self.ENV_OPTIONS:
                self.env_var.set(cfg['ENV'])
                self.update_client_options()
            if cfg.get('CLIENT_ID') in self.CLIENT_MAP.get(self.env_var.get(), []):
                self.client_var.set(cfg['CLIENT_ID'])
            for key in ['USERID', 'PASSWORD']:
                if cfg.get(key):
                    val = cfg[key]
                    if key == 'PASSWORD':
                        try: val = base64.b64decode(val).decode()
                        except: val = ''
                    self.entries[key].insert(0, val)
        except FileNotFoundError:
            pass

    def update_client_options(self):
        opts = self.CLIENT_MAP.get(self.env_var.get(), [])
        self.client_cb['values'] = opts
        if opts and self.client_var.get() not in opts: self.client_var.set(opts[0])

if __name__ == '__main__':
    root = tk.Tk()
    app = AutomicToolsApp(root)
    root.mainloop()