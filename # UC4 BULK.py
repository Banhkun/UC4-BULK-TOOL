import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import base64
import automic_rest as automic
import copy
import re

#----------------------------
# Helpers
#----------------------------

def sanitize_string(s):
    s = re.sub(r'^[^0-9A-Za-z_]+', '', s)
    return re.sub(r'[^0-9A-Za-z_]', '_', s)


def parse_flexible_pairs(data):
    headers_by_length = {
        3: ["jobname", "program", "variant"],
        4: ["jobname", "program", "variant", "user"],
        5: ["jobname", "program", "variant", "user", "language"],
        6: ["jobname", "program", "variant", "user", "language", "extra"]
    }
    parsed = []
    for line in data.strip().splitlines():
        parts = line.strip().split()
        n = len(parts)
        if n == 2:
            program, variant = parts
            jobname = f"{sanitize_string(program)}_{sanitize_string(variant)}"
            parsed.append({"jobname": jobname, "program": program, "variant": variant})
        elif n in headers_by_length:
            parsed.append(dict(zip(headers_by_length[n], parts)))
        else:
            parsed.append({f"col_{i+1}": v for i, v in enumerate(parts)})
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
# UI Application
#----------------------------
class JobCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automic Job Creator")
        self.build_ui()

    def build_ui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill='both', expand=True)
        # Config entries
        labels = ["ENV", "USERID", "PASSWORD", "CLIENT_ID", "ARMT_NO"]
        self.entries = {}
        for i, lb in enumerate(labels):
            ttk.Label(frm, text=lb+":").grid(row=i, column=0, sticky='w')
            ent = ttk.Entry(frm, show='*' if lb=='PASSWORD' else None)
            ent.grid(row=i, column=1, sticky='ew', pady=2)
            self.entries[lb] = ent
        # Pairs
        ttk.Label(frm, text="PAIRS_DATA:").grid(row=5, column=0, sticky='nw', pady=(10,0))
        self.pairs_text = tk.Text(frm, height=6)
        self.pairs_text.grid(row=5, column=1, sticky='ew', pady=(10,2))
        # Templates
        ttk.Label(frm, text="template_job_armt:").grid(row=6, column=0, sticky='w')
        self.template_job_armt = ttk.Entry(frm)
        self.template_job_armt.grid(row=6, column=1, sticky='ew', pady=2)
        ttk.Label(frm, text="template_joplan_armt:").grid(row=7, column=0, sticky='w')
        self.template_joplan_armt = ttk.Entry(frm)
        self.template_joplan_armt.grid(row=7, column=1, sticky='ew', pady=2)
        # Execute button
        self.run_btn = ttk.Button(frm, text="Create Jobs", command=self.start)
        self.run_btn.grid(row=8, column=0, columnspan=2, pady=10)
        # Output log
        ttk.Label(frm, text="Output:").grid(row=9, column=0, sticky='nw')
        self.log_box = scrolledtext.ScrolledText(frm, height=10, state='disabled')
        self.log_box.grid(row=9, column=1, sticky='ew')
        frm.columnconfigure(1, weight=1)

    def log(self, msg):
        self.log_box.configure(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    def start(self):
        self.run_btn.config(state='disabled')
        threading.Thread(target=self.execute, daemon=True).start()

    def execute(self):
        # Read inputs
        env = self.entries['ENV'].get().strip()
        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        cid = int(self.entries['CLIENT_ID'].get().strip())
        armt = self.entries['ARMT_NO'].get().strip()
        t_job = self.template_job_armt.get().strip()
        t_joplan = self.template_joplan_armt.get().strip()
        raw = self.pairs_text.get('1.0', 'end')
        # Authenticate
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)
        # Fetch
        self.log(f"Fetching {t_joplan} and {t_job}")
        rp = automic.getObjects(client_id=cid, object_name=t_joplan)
        tmpl_jobp = rp.response['data']['jobp']
        rj = automic.getObjects(client_id=cid, object_name=t_job)
        tmpl_jobs = rj.response['data']['jobs']
        base_jobp = tmpl_jobp['general_attributes']['name'][:31]
        base_jobs = tmpl_jobp['general_attributes']['name'][:21]
        # Pairs
        pairs = parse_flexible_pairs(raw)
        # Default login
        default_login = extract_default_login(tmpl_jobs)
        # Loop
        for p in pairs:
            jn = p['jobname']
            name_jobp = f"{base_jobp}_{jn}"
            name_jobs = f"{base_jobs}_{jn}"
            # JOBP
            njp = copy.deepcopy(tmpl_jobp)
            njp['general_attributes']['name'] = name_jobp
            for wf in njp.get('workflow_definitions', []):
                if wf.get('object_name') == tmpl_jobs['general_attributes']['name']:
                    wf['object_name'] = name_jobs
            res_p = automic.postObjects(client_id=cid, body={'total':1,'data':{'jobp':njp},'path':f'AUTOMATION_JOBS/{user}/{armt}','client':cid,'hasmore':False})
            if res_p.status==None: self.log(f"JOBP: {name_jobp}")
            else: self.log(f"FAIL JOBP: {name_jobp} ({res_p.status})")
            # R3 script
            login_val = f"LOGIN_R3_060_{p.get('login', default_login)}"
            script = [
                f":INC BSH_XXXX_INC_MIGRATION_SIMULATION WAIT_TIME = \"<Random number ...>\" ,NOFOUND=IGNORE",
                f":PUT_ATT JOB_NAME= \"{jn}\"",
                f":PUT_ATT LOGIN='{login_val}'",
                f"R3_ACTIVATE_REPORT REPORT='{p['program']}',VARIANT='{p['variant']}',COPIES=1,EXPIR=8,LINE_COUNT=65,LINE_SIZE=80,LAYOUT=X_FORMAT,DATA_SET=LIST1S,TYPE=TEXT"
            ]
            nj = copy.deepcopy(tmpl_jobs)
            nj['general_attributes']['name'] = name_jobs
            for proc in nj.get('scripts', []):
                if 'process' in proc:
                    proc['process'] = script
            res_j = automic.postObjects(client_id=cid, body={'total':1,'data':{'jobs':nj},'path':f'AUTOMATION_JOBS/{user}/{armt}','client':cid,'hasmore':False})
            if res_j.status==None: self.log(f"JOBS: {name_jobs}")
            else: self.log(f"FAIL JOBS: {name_jobs} ({res_j.status})")
        self.log("All done.")
        self.run_btn.config(state='normal')

if __name__ == '__main__':
    root = tk.Tk()
    JobCreatorApp(root)
    root.mainloop()
