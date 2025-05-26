import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import base64
import automic_rest as automic
import copy
import re
import os
import json

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
    CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.automic_job_creator.json')
    ENV_OPTIONS = ['eup6', 'eup7']
    CLIENT_MAP = {
        'eup6': ['1001', '1111'],
        'eup7': ['1001', '1301', '1401', '7101']
    }

    def __init__(self, root):
        self.root = root
        self.setup_style()
        self.root.title("Automic Job Creator")
        self.entries = {}
        self.load_config()
        self.build_ui()
        self.populate_fields()

    def setup_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        accent = '#0096D6'
        bg = '#F0F2F5'
        font = ('Montserrat', 10)
        self.root.configure(bg=bg)
        style.configure('TFrame', background=bg)
        style.configure('TLabel', background=bg, font=font)
        style.configure('TEntry', font=font)
        style.configure('TCombobox', font=font)
        style.configure('TButton', font=font, background=accent, foreground='white', padding=6)
        style.map('TButton', background=[('active', '#007BB5')])
        style.configure('Treeview', font=font, rowheight=24)

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
            'PAIRS_DATA': self.pairs_text.get('1.0', 'end')
        }
        with open(self.CONFIG_PATH, 'w') as f:
            json.dump(data, f)

    def populate_fields(self):
        cfg = self.config
        if cfg.get('ENV') in self.ENV_OPTIONS:
            self.env_var.set(cfg['ENV']); self.update_client_options()
        if cfg.get('CLIENT_ID') in self.CLIENT_MAP.get(self.env_var.get(), []):
            self.client_var.set(cfg['CLIENT_ID'])
        for key in ['USERID', 'PASSWORD', 'ARMT_NO']:
            if cfg.get(key):
                val = cfg[key]
                if key=='PASSWORD':
                    try: val = base64.b64decode(val).decode()
                    except: val = ''
                self.entries[key].insert(0, val)
        for fld in ['template_job_armt','template_joplan_armt']:
            if cfg.get(fld): getattr(self, fld).insert(0, cfg[fld])
        if cfg.get('PAIRS_DATA'):
            self.pairs_text.insert('1.0', cfg['PAIRS_DATA'])

    def build_ui(self):
        frm = ttk.Frame(self.root, padding=15)
        frm.pack(fill='both', expand=True)
        # ENV & CLIENT
        ttk.Label(frm,text='Environment:').grid(row=0,column=0,sticky='w')
        self.env_var=tk.StringVar(); env_cb=ttk.Combobox(frm,textvariable=self.env_var,values=self.ENV_OPTIONS,state='readonly'); env_cb.grid(row=0,column=1,sticky='ew',padx=5, pady=2)
        env_cb.bind('<<ComboboxSelected>>',lambda e:self.update_client_options())
        ttk.Label(frm,text='Client ID:').grid(row=0,column=2,sticky='w')
        self.client_var=tk.StringVar(); self.client_cb=ttk.Combobox(frm,textvariable=self.client_var,state='readonly'); self.client_cb.grid(row=0,column=3,sticky='ew',padx=5,pady=2)
        # Credentials
        ttk.Label(frm,text='User ID:').grid(row=1,column=0,sticky='w'); self.entries['USERID']=ttk.Entry(frm); self.entries['USERID'].grid(row=1,column=1,sticky='ew',padx=5)
        ttk.Label(frm,text='Password:').grid(row=1,column=2,sticky='w'); self.entries['PASSWORD']=ttk.Entry(frm,show='*'); self.entries['PASSWORD'].grid(row=1,column=3,sticky='ew',padx=5)
        ttk.Label(frm,text='ARMT No.:').grid(row=2,column=0,sticky='w'); self.entries['ARMT_NO']=ttk.Entry(frm); self.entries['ARMT_NO'].grid(row=2,column=1,sticky='ew',padx=5)
        # Templates
        ttk.Label(frm,text='Jobplan Template:').grid(row=3,column=0,sticky='w'); self.template_joplan_armt=ttk.Entry(frm); self.template_joplan_armt.grid(row=3,column=1,columnspan=3,sticky='ew',padx=5,pady=2)
        ttk.Label(frm,text='Jobs Template:').grid(row=4,column=0,sticky='w'); self.template_job_armt=ttk.Entry(frm); self.template_job_armt.grid(row=4,column=1,columnspan=3,sticky='ew',padx=5,pady=2)
        # Pairs Data
        ttk.Label(frm,text='Program/Variant Pairs:').grid(row=5,column=0,sticky='nw', pady=(10,2))
        self.pairs_text=scrolledtext.ScrolledText(frm,height=6); self.pairs_text.grid(row=5,column=1,columnspan=3,sticky='ew',padx=5)
        # Run Button
        self.run_btn=ttk.Button(frm,text='Create Jobs',command=self.start); self.run_btn.grid(row=6,column=0,columnspan=4,pady=12)
        # Output Log
        ttk.Label(frm,text='Output:').grid(row=7,column=0,sticky='nw')
        self.log_box=scrolledtext.ScrolledText(frm,height=10,state='disabled'); self.log_box.grid(row=7,column=1,columnspan=3,sticky='ew',padx=5)
        frm.columnconfigure((1,3),weight=1)

    def update_client_options(self):
        opts=self.CLIENT_MAP.get(self.env_var.get(),[])
        self.client_cb['values']=opts
        if opts and self.client_var.get() not in opts: self.client_var.set(opts[0])

    def log(self,msg):
        self.log_box.config(state='normal'); self.log_box.insert('end',msg+'\n'); self.log_box.see('end'); self.log_box.config(state='disabled')

    def start(self):
        self.run_btn.config(state='disabled'); threading.Thread(target=self.execute,daemon=True).start()

    def execute(self):
        # Read inputs
        env = self.env_var.get().strip()
        cid = int(self.client_var.get().strip())
        user = self.entries['USERID'].get().strip()
        pwd = self.entries['PASSWORD'].get().strip()
        armt = self.entries['ARMT_NO'].get().strip()
        api_url = f'https://rb-{env}-api.bosch.com'
        t_job = self.template_job_armt.get().strip()
        t_joplan = self.template_joplan_armt.get().strip()
        raw = self.pairs_text.get('1.0', 'end')

        self.save_config()

        # Authenticate
        auth = base64.b64encode(f"{user}:{pwd}".encode()).decode()
        automic.connection(url=api_url, auth=auth, noproxy=True, sslverify=False)

        # Fetch templates
        if t_joplan:
            self.log(f"Fetching jobplan {t_joplan}")
            rp = automic.getObjects(client_id=cid, object_name=t_joplan)
            tmpl_jobp = rp.response['data']['jobp']
            base_jobp = tmpl_jobp['general_attributes']['name'][:31]
        else:
            tmpl_jobp = None
            base_jobp = ''

        self.log(f"Fetching job {t_job}")
        rj = automic.getObjects(client_id=cid, object_name=t_job)
        tmpl_jobs = rj.response['data']['jobs']
        base_jobs = tmpl_jobs['general_attributes']['name'][:21]

        # Parse pairs and default login
        pairs = parse_flexible_pairs(raw)
        default_login = extract_default_login(tmpl_jobs)

        # Loop create
        jobps = []
        for p in pairs:
            jn = p['jobname']
            if tmpl_jobp:
                name_jobp = f"{base_jobp}_{jn}"
                jobps.append(name_jobp)
                njp = copy.deepcopy(tmpl_jobp)
                njp['general_attributes']['name'] = name_jobp
                for wf in njp.get('workflow_definitions', []):
                    if wf.get('object_name') == tmpl_jobs['general_attributes']['name']:
                        wf['object_name'] = f"{base_jobs}_{jn}"
                res_p = automic.postObjects(client_id=cid, body={'total':1,'data':{'jobp':njp},'path':f'AUTOMATION_JOBS/{user}/{armt}','client':cid,'hasmore':False})
                self.log(f"JOBP: {name_jobp}" if res_p.status==None else f"FAIL JOBP: {name_jobp} ({res_p.status})")

            name_jobs = f"{base_jobs}_{jn}"
            # Create JOBS_R3
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
            self.log(f"JOBS: {name_jobs}" if res_j.status==None else f"FAIL JOBS: {name_jobs} ({res_j.status})")

        self.log("All done.")
        self.run_btn.config(state='normal')

if __name__ == '__main__':
    root = tk.Tk()
    JobCreatorApp(root)
    root.mainloop()
