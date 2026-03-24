
import csv
import json
import os
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_TITLE = "PT. Sagamartha Ultima Indonesia - HR Enterprise Desktop"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(BASE_DIR, "enterprise_records.csv")
USERS_FILE = os.path.join(BASE_DIR, "enterprise_users.json")
PENDING_USERS_FILE = os.path.join(BASE_DIR, "enterprise_pending_users.json")
ATTENDANCE_FILE = os.path.join(BASE_DIR, "enterprise_attendance.csv")
SETTINGS_FILE = os.path.join(BASE_DIR, "enterprise_settings.json")
LOGO_FILE = os.path.join(BASE_DIR, "LOGO SAGAMARTHA.png")
SHEET_NAME = "EnterpriseHRData"

FIELDNAMES = [
    "record_id",
    "employee_name",
    "department",
    "role_scope",
    "supervisor_name",
    "work_date",
    "project_title",
    "target_description",
    "target_category",
    "daily_target",
    "daily_performance",
    "achievement_percent",
    "overtime_date",
    "start_time",
    "end_time",
    "duration_hours",
    "reason",
    "manager_approval",
    "comments",
    "created_by",
    "created_at",
    "updated_at",
]

ATTENDANCE_FIELDS = [
    "attendance_id",
    "employee_name",
    "department",
    "attendance_date",
    "check_in_time",
    "check_out_time",
    "status",
    "notes",
    "created_at",
]

DEFAULT_SETTINGS = {
    "mode": "local",
    "google_credentials_file": "",
    "google_spreadsheet_id": "",
    "auto_sync_on_save": True,
    "window_width": 1600,
    "window_height": 930
}

DEFAULT_USERS = [
    {
        "username": "adi",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Adi Purnomo, S. Akun.",

        "supervisor_name": "Annisa Dira Hariyanto, S.T., M.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "ufaira",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Ufaira Aulia Nur Ramadhani, S.Pi.",

        "supervisor_name": "Andini Putri Salsabillah, S.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "karina",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Karina Azizah, S.Pi.",

        "supervisor_name": "Dimas Tri Rendra Graha, S.T., M. Ars.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "zahra",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Zahra Mustafafi",

        "supervisor_name": "Annisa Dira Hariyanto, S.T., M.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "syahwa",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Syahwa Novelia Safitri",

        "supervisor_name": "Andini Putri Salsabillah, S.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "angga",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "admin",
        "employee_name": "Angga Anugerah Ardana, S.T.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "johan",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "admin",
        "employee_name": "Johan Wahyu Panuntun, S.T.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "annisa",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "supervisor",
        "employee_name": "Annisa Dira Hariyanto, S.T., M.P.W.K.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "nuryantiningsih",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "admin",
        "employee_name": "Nuryantiningsih Pusporini, S.T., M. Ars.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "fiko",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Fiko Virgin Septarina, S.P.W.K.",

        "supervisor_name": "Dimas Tri Rendra Graha, S.T., M. Ars.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "rizal",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Rizal Brilliant Nugraha, S.P.W.K.",

        "supervisor_name": "Annisa Dira Hariyanto, S.T., M.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "muhammad",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Muhammad Kholifatkhur Rohman",

        "supervisor_name": "Andini Putri Salsabillah, S.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "hardi",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Hardi Adityasna, S.P.W.K.",

        "supervisor_name": "Dimas Tri Rendra Graha, S.T., M. Ars.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "angelly",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Angelly Taruli Emmanuella, S.P.W.K.",

        "supervisor_name": "Annisa Dira Hariyanto, S.T., M.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "firman",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "admin",
        "employee_name": "Firman Afrianto, S.T., M.T.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "nayyara",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Nayyara Nazmi Fayyaza",

        "supervisor_name": "Andini Putri Salsabillah, S.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "afdhal",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Afdhal Ibnu Asya'fa",

        "supervisor_name": "Dimas Tri Rendra Graha, S.T., M. Ars.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "muhammad.afri",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "karyawan",
        "employee_name": "Muhammad Afri Naufal",

        "supervisor_name": "Annisa Dira Hariyanto, S.T., M.P.W.K.",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "andini",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "supervisor",
        "employee_name": "Andini Putri Salsabillah, S.P.W.K.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    },
    {
        "username": "dimas",
        "password": "6adb0cd9dbaf187ad4435d1a3e975c3803f718a17e0feff8335a14573a82b5d0",
        "role": "supervisor",
        "employee_name": "Dimas Tri Rendra Graha, S.T., M. Ars.",

        "supervisor_name": "",
        "is_active": true,
        "must_change_password": true
    }
]

TARGET_CATEGORIES = [
    "Analisis Spasial",
    "Produksi Peta",
    "Dokumen Perencanaan",
    "Analisis Data",
    "Koordinasi Stakeholder",
    "Modeling & Simulasi",
    "Perencanaan Teknis",
    "Pelaporan & Presentasi",
    "Administrasi Proyek",
    "Riset & Inovasi"
]


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str():
    return datetime.now().strftime("%Y-%m-%d")


def hash_password(password, salt="sagamartha_hr"):
    import hashlib
    return hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120000
    ).hex()


def verify_password(stored_value, plain_password):
    if not stored_value:
        return False
    if len(stored_value) == 64 and all(c in "0123456789abcdef" for c in stored_value.lower()):
        return stored_value == hash_password(plain_password)
    return stored_value == plain_password


def normalize_users_security(users):
    changed = False
    for user in users:
        pw = user.get("password", "")
        if not (len(pw) == 64 and all(c in "0123456789abcdef" for c in pw.lower())):
            user["password"] = hash_password(pw)
            changed = True
        if "must_change_password" not in user:
            user["must_change_password"] = False
            changed = True
    return users, changed


def ensure_csv(path, fields):
    if not os.path.exists(path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()


def load_csv(path, fields):
    ensure_csv(path, fields)
    with open(path, "r", newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def save_csv(path, fields, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def append_csv(path, fields, row):
    ensure_csv(path, fields)
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writerow(row)


def ensure_json(path, default_value):
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(default_value, f, indent=2)


def load_json(path, default_value):
    ensure_json(path, default_value)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d")


def parse_time(s):
    return datetime.strptime(s, "%H:%M")


def calculate_duration(start_str, end_str):
    start = parse_time(start_str)
    end = parse_time(end_str)
    delta = end - start
    if delta.total_seconds() < 0:
        delta += timedelta(days=1)
    return round(delta.total_seconds() / 3600, 2)


def monday_of_week(date_str):
    dt = parse_date(date_str)
    return dt - timedelta(days=dt.weekday())


def calculate_achievement_percent(target_value, performance_value):
    if target_value <= 0:
        return 0.0
    return round((performance_value / target_value) * 100.0, 2)


def next_id(rows, key_name):
    if not rows:
        return "1"
    return str(max(int(r.get(key_name, "0") or 0) for r in rows) + 1)


def load_settings():
    return load_json(SETTINGS_FILE, DEFAULT_SETTINGS)


def save_settings(data):
    save_json(SETTINGS_FILE, data)


class GoogleSheetsSync:
    def __init__(self, credentials_file, spreadsheet_id):
        self.credentials_file = credentials_file
        self.spreadsheet_id = spreadsheet_id
        self.service = None

    def connect(self):
        try:
            from google.oauth2.service_account import Credentials
            from googleapiclient.discovery import build
        except ImportError as e:
            raise RuntimeError(
                "Library Google API belum terpasang. Jalankan:\n"
                "pip install google-api-python-client google-auth"
            ) from e

        if not self.credentials_file or not os.path.exists(self.credentials_file):
            raise RuntimeError("File credentials JSON Google belum dipilih atau tidak ditemukan.")
        if not self.spreadsheet_id:
            raise RuntimeError("Spreadsheet ID Google Sheets belum diisi.")

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(self.credentials_file, scopes=scopes)
        self.service = build("sheets", "v4", credentials=creds)
        return self.service

    def ensure_sheet_exists(self):
        service = self.service or self.connect()
        meta = service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if SHEET_NAME not in titles:
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": SHEET_NAME}}}]}
            ).execute()

    def upload_all(self, records):
        service = self.service or self.connect()
        self.ensure_sheet_exists()
        rows = [FIELDNAMES] + [[r.get(f, "") for f in FIELDNAMES] for r in records]
        service.spreadsheets().values().clear(
            spreadsheetId=self.spreadsheet_id,
            range=f"{SHEET_NAME}!A:Z"
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            body={"values": rows}
        ).execute()

    def download_all(self):
        service = self.service or self.connect()
        self.ensure_sheet_exists()
        values = service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"{SHEET_NAME}!A:Z"
        ).execute().get("values", [])
        if not values:
            return []
        rows = []
        for row in values[1:]:
            item = {}
            for i, col in enumerate(FIELDNAMES):
                item[col] = row[i] if i < len(row) else ""
            rows.append(item)
        return rows

    def create_spreadsheet(self, title="Sagamartha HR Enterprise Data"):
        service = self.service or self.connect()
        body = {
            "properties": {"title": title},
            "sheets": [{"properties": {"title": SHEET_NAME}}]
        }
        spreadsheet = service.spreadsheets().create(body=body).execute()
        self.spreadsheet_id = spreadsheet["spreadsheetId"]
        service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            body={"values": [FIELDNAMES]}
        ).execute()
        return spreadsheet

    def share_spreadsheet(self, email, role="writer"):
        try:
            from googleapiclient.discovery import build
        except ImportError as e:
            raise RuntimeError("google-api-python-client belum terpasang.") from e
        if not email:
            raise RuntimeError("Email tujuan share belum diisi.")
        self.service = self.service or self.connect()
        drive = build("drive", "v3", credentials=self.service._http.credentials)
        permission = {
            "type": "user",
            "role": role,
            "emailAddress": email
        }
        drive.permissions().create(
            fileId=self.spreadsheet_id,
            body=permission,
            sendNotificationEmail=False
        ).execute()



def is_duplicate_record(records, candidate, exclude_record_id=None):
    keys = [
        "employee_name",
        "work_date",
        "project_title",
        "target_description",
        "target_category",
        "daily_target",
        "daily_performance",
        "overtime_date",
        "start_time",
        "end_time",
        "reason",
    ]
    for row in records:
        if exclude_record_id and row.get("record_id") == exclude_record_id:
            continue
        same = True
        for key in keys:
            if str(row.get(key, "")).strip() != str(candidate.get(key, "")).strip():
                same = False
                break
        if same:
            return True
    return False


def attendance_exists(attendance_rows, employee_name, attendance_date, exclude_id=None):
    for row in attendance_rows:
        if exclude_id and row.get("attendance_id") == exclude_id:
            continue
        if row.get("employee_name") == employee_name and row.get("attendance_date") == attendance_date:
            return True
    return False


class RegisterDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Registrasi Karyawan")
        self.geometry("520x360")
        self.resizable(False, False)
        self.build_ui()
        self.grab_set()

    def build_ui(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.employee_name_var = tk.StringVar()
        self.department_var = tk.StringVar() # deprecated
        self.supervisor_name_var = tk.StringVar()

        fields = [
            ("Username", self.username_var),
            ("Password", self.password_var),
            ("Nama Karyawan", self.employee_name_var),
            # ("Departemen", self.department_var),
            ("Nama Supervisor", self.supervisor_name_var),
        ]
        for i, (label, var) in enumerate(fields):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=6, padx=(0, 8))
            show = "*" if label == "Password" else ""
            ttk.Entry(frm, textvariable=var, width=34, show=show).grid(row=i, column=1, sticky="ew", pady=6)

        note = (
            "Akun karyawan dapat dibuat sendiri, tetapi statusnya Pending Approval.\n"
            "Admin harus menyetujui dulu sebelum akun bisa dipakai login."
        )
        ttk.Label(frm, text=note, justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 12))
        ttk.Button(frm, text="Kirim Registrasi", command=self.submit).grid(row=6, column=0, columnspan=2, sticky="ew")
        frm.columnconfigure(1, weight=1)

    def submit(self):
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        employee_name = self.employee_name_var.get().strip()
        department = self.department_var.get().strip()
        supervisor_name = self.supervisor_name_var.get().strip()

        if not all([username, password, employee_name, department, supervisor_name]):
            messagebox.showerror("Validasi", "Semua field wajib diisi.")
            return

        users = load_json(USERS_FILE, DEFAULT_USERS)
        pending = load_json(PENDING_USERS_FILE, [])
        if any(u["username"] == username for u in users) or any(u["username"] == username for u in pending):
            messagebox.showerror("Registrasi gagal", "Username sudah dipakai.")
            return

        pending.append({
            "username": username,
            "password": hash_password(password),
            "role": "karyawan",
            "employee_name": employee_name,
            "department": department,
            "supervisor_name": supervisor_name,
            "is_active": False,
            "must_change_password": False,
            "requested_at": now_str()
        })
        save_json(PENDING_USERS_FILE, pending)
        messagebox.showinfo("Berhasil", "Registrasi dikirim. Tunggu persetujuan admin.")
        self.destroy()


class LoginWindow(ttk.Frame):
    def __init__(self, master, on_login):
        super().__init__(master, padding=24)
        self.master = master
        self.on_login = on_login
        self.logo_ref = None
        self.build_ui()

    def build_ui(self):
        self.pack(fill="both", expand=True)
        box = ttk.Frame(self)
        box.place(relx=0.5, rely=0.5, anchor="center")

        if os.path.exists(LOGO_FILE):
            try:
                logo = tk.PhotoImage(file=LOGO_FILE)
                logo = logo.subsample(max(1, int(logo.width() / 340)), max(1, int(logo.height() / 140)))
                self.logo_ref = logo
                ttk.Label(box, image=self.logo_ref).pack(pady=(0, 16))
            except Exception:
                ttk.Label(box, text="SAGAMARTHA", font=("Segoe UI", 24, "bold")).pack(pady=(0, 16))
        else:
            ttk.Label(box, text="SAGAMARTHA", font=("Segoe UI", 24, "bold")).pack(pady=(0, 16))

        ttk.Label(box, text="HR Enterprise Desktop", font=("Segoe UI", 16, "bold")).pack(pady=(0, 10))

        form = ttk.Frame(box)
        form.pack()

        self.username_var = tk.StringVar(value="admin")
        self.password_var = tk.StringVar(value="admin123")

        ttk.Label(form, text="Username").grid(row=0, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(form, textvariable=self.username_var, width=28).grid(row=0, column=1, pady=6)
        ttk.Label(form, text="Password").grid(row=1, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(form, textvariable=self.password_var, width=28, show="*").grid(row=1, column=1, pady=6)

        btns = ttk.Frame(box)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Login", command=self.attempt_login).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(btns, text="Registrasi Karyawan", command=self.register).pack(side="left", fill="x", expand=True, padx=(4, 0))

        ttk.Label(
            box,
            text="Developed by Firman Afrianto",
            justify="center"
        ).pack(pady=(12, 0))

    def attempt_login(self):
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        users = load_json(USERS_FILE, DEFAULT_USERS)
        users, changed = normalize_users_security(users)
        if changed:
            save_json(USERS_FILE, users)
        for user in users:
            if user["username"] == username and verify_password(user["password"], password):
                if not user.get("is_active", True):
                    messagebox.showwarning("Akun belum aktif", "Akun belum diaktifkan admin.")
                    return
                self.on_login(user)
                return
        messagebox.showerror("Login gagal", "Username atau password salah.")

    def register(self):
        RegisterDialog(self.master)


class SettingsDialog(tk.Toplevel):
    def __init__(self, master, settings):
        super().__init__(master)
        self.title("Pengaturan Sinkronisasi")
        self.geometry("740x300")
        self.resizable(False, False)
        self.result = None

        self.mode_var = tk.StringVar(value=settings.get("mode", "local"))
        self.cred_var = tk.StringVar(value=settings.get("google_credentials_file", ""))
        self.sheet_var = tk.StringVar(value=settings.get("google_spreadsheet_id", ""))
        self.auto_sync_var = tk.BooleanVar(value=settings.get("auto_sync_on_save", True))
        self.window_width_var = tk.StringVar(value=str(settings.get("window_width", 1600)))
        self.window_height_var = tk.StringVar(value=str(settings.get("window_height", 930)))

        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Mode Penyimpanan").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Combobox(body, textvariable=self.mode_var, state="readonly", values=["local", "google_sheets"], width=28).grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(body, text="Credentials JSON").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(body, textvariable=self.cred_var, width=60).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(body, text="Browse", command=self.browse_credentials).grid(row=1, column=2, padx=(8, 0), pady=6)

        ttk.Label(body, text="Spreadsheet ID").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(body, textvariable=self.sheet_var, width=60).grid(row=2, column=1, sticky="ew", pady=6)

        ttk.Checkbutton(body, text="Auto sync saat simpan", variable=self.auto_sync_var).grid(row=3, column=1, sticky="w", pady=(8, 6))

        size_frame = ttk.Frame(body)
        size_frame.grid(row=4, column=1, sticky="w", pady=(0, 12))
        ttk.Label(body, text="Ukuran Awal Window").grid(row=4, column=0, sticky="w", pady=(0, 12))
        ttk.Entry(size_frame, textvariable=self.window_width_var, width=8).pack(side="left")
        ttk.Label(size_frame, text="x").pack(side="left", padx=4)
        ttk.Entry(size_frame, textvariable=self.window_height_var, width=8).pack(side="left")

        ttk.Button(body, text="Simpan", command=self.on_save).grid(row=5, column=1, sticky="e")
        body.columnconfigure(1, weight=1)
        self.grab_set()

    def browse_credentials(self):
        path = filedialog.askopenfilename(title="Pilih credentials JSON", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if path:
            self.cred_var.set(path)

    def on_save(self):
        try:
            width = int(self.window_width_var.get().strip())
            height = int(self.window_height_var.get().strip())
        except Exception:
            messagebox.showerror("Input salah", "Ukuran window harus berupa angka.")
            return
        self.result = {
            "mode": self.mode_var.get().strip(),
            "google_credentials_file": self.cred_var.get().strip(),
            "google_spreadsheet_id": self.sheet_var.get().strip(),
            "auto_sync_on_save": bool(self.auto_sync_var.get()),
            "window_width": width,
            "window_height": height,
        }
        self.destroy()


class HRApp(ttk.Frame):
    def __init__(self, master, user):
        super().__init__(master, padding=10)
        self.master = master
        self.current_user = user
        self.settings = load_settings()
        self.records = load_csv(DATA_FILE, FIELDNAMES)
        self.attendance = load_csv(ATTENDANCE_FILE, ATTENDANCE_FIELDS)
        self.users = load_json(USERS_FILE, DEFAULT_USERS)
        self.pending_users = load_json(PENDING_USERS_FILE, [])
        self.logo_ref = None

        self.search_name_var = tk.StringVar()
        self.search_date_var = tk.StringVar()
        self.filter_month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        self.user_role_filter_var = tk.StringVar(value="Semua")

        self.employee_name_var = tk.StringVar()
        self.department_var = tk.StringVar() # deprecated
        self.supervisor_name_var = tk.StringVar()
        self.work_date_var = tk.StringVar(value=today_str())
        self.project_title_var = tk.StringVar()
        self.target_description_var = tk.StringVar()
        self.target_category_var = tk.StringVar(value=TARGET_CATEGORIES[0])
        self.daily_target_var = tk.StringVar()
        self.daily_performance_var = tk.StringVar()
        self.achievement_percent_var = tk.StringVar(value="0.00")
        self.overtime_date_var = tk.StringVar(value=today_str())
        self.start_time_var = tk.StringVar()
        self.end_time_var = tk.StringVar()
        self.duration_var = tk.StringVar(value="0.00")
        self.reason_var = tk.StringVar()
        self.approval_var = tk.StringVar(value="Pending")
        self.comments_var = tk.StringVar()

        self.selected_record_id = None
        self.tree = None
        self.recap_tree = None
        self.att_tree = None
        self.pending_tree = None
        self.users_tree = None
        self.status_label = None
        self.mode_badge = None
        self.dashboard_cards = {}
        self.fig = None
        self.canvas = None

        self.pack(fill="both", expand=True)
        self.build_ui()
        self.apply_user_defaults()
        self.refresh_all()
        if self.current_user.get("must_change_password", False):
            self.master.after(300, self.change_password_force)

    def build_ui(self):
        topbar = ttk.Frame(self)
        topbar.pack(fill="x", pady=(0, 8))

        left = ttk.Frame(topbar)
        left.pack(side="left", fill="x", expand=True)

        if os.path.exists(LOGO_FILE):
            try:
                logo = tk.PhotoImage(file=LOGO_FILE)
                logo = logo.subsample(max(1, int(logo.width() / 180)), max(1, int(logo.height() / 70)))
                self.logo_ref = logo
                ttk.Label(left, image=self.logo_ref).pack(side="left", padx=(0, 10))
            except Exception:
                pass

        meta = ttk.Frame(left)
        meta.pack(side="left")
        ttk.Label(meta, text="PT. Sagamartha Ultima Indonesia", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(meta, text=f"Login: {self.current_user['employee_name']} ({self.current_user['role']})").pack(anchor="w")

        right = ttk.Frame(topbar)
        right.pack(side="right")
        ttk.Button(right, text="Pengaturan Sync", command=self.open_settings).pack(side="left", padx=4)
        ttk.Button(right, text="Buat Google Sheet", command=self.auto_create_google_sheet).pack(side="left", padx=4)
        ttk.Button(right, text="Upload Cloud", command=self.sync_upload).pack(side="left", padx=4)
        ttk.Button(right, text="Download Cloud", command=self.sync_download).pack(side="left", padx=4)
        ttk.Button(right, text="Ubah Password", command=self.change_password).pack(side="left", padx=4)
        ttk.Button(right, text="Logout", command=self.logout).pack(side="left", padx=4)
        self.mode_badge = ttk.Label(right, text="")
        self.mode_badge.pack(side="left", padx=(8, 0))
        self.update_mode_badge()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.tab_dashboard = ttk.Frame(self.notebook, padding=10)
        self.tab_entry = ttk.Frame(self.notebook, padding=10)
        self.tab_records = ttk.Frame(self.notebook, padding=10)
        self.tab_recap = ttk.Frame(self.notebook, padding=10)
        self.tab_attendance = ttk.Frame(self.notebook, padding=10)

        role = self.current_user["role"]
        self.notebook.add(self.tab_dashboard, text="Dashboard")
        self.notebook.add(self.tab_entry, text="Input")
        self.notebook.add(self.tab_records, text="Data Saya" if role != "admin" else "Semua Data")
        self.notebook.add(self.tab_attendance, text="Absensi")

        if role == "admin":
            self.notebook.add(self.tab_recap, text="Rekap Bulanan")
            self.tab_users = ttk.Frame(self.notebook, padding=10)
            self.notebook.add(self.tab_users, text="User Management")
        elif role == "supervisor":
            self.notebook.add(self.tab_recap, text="Rekap Tim")
            self.tab_users = None
        else:
            self.notebook.add(self.tab_recap, text="Rekap Saya")
            self.tab_users = None

        self.build_dashboard_tab()
        self.build_entry_tab()
        self.build_records_tab()
        self.build_recap_tab()
        self.build_attendance_tab()
        if self.tab_users:
            self.build_users_tab()

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", pady=(8, 0))
        self.status_label = ttk.Label(bottom, text="Siap.")
        self.status_label.pack(side="left")

    def build_dashboard_tab(self):
        cards = ttk.Frame(self.tab_dashboard)
        cards.pack(fill="x", pady=(0, 12))
        role = self.current_user["role"]
        if role == "admin":
            labels = ["Total Jam Lembur Bulan Ini", "Rata-rata Capaian", "Pending Approval", "Jumlah Karyawan Terlihat", "Kehadiran Hari Ini"]
        elif role == "supervisor":
            labels = ["Total Jam Lembur Tim", "Rata-rata Capaian Tim", "Pending Approval Tim", "Jumlah Anggota Tim", "Kehadiran Hari Ini"]
        else:
            labels = ["Total Jam Lembur Saya", "Rata-rata Capaian Saya", "Pending Approval Saya", "Jumlah Data Saya", "Kehadiran Hari Ini"]
        for i, label in enumerate(labels):
            box = ttk.LabelFrame(cards, text=label, padding=10)
            box.grid(row=0, column=i, sticky="nsew", padx=6)
            val = ttk.Label(box, text="0", font=("Segoe UI", 18, "bold"))
            val.pack()
            self.dashboard_cards[label] = val
            cards.columnconfigure(i, weight=1)

        chart_wrap = ttk.LabelFrame(self.tab_dashboard, text="Grafik Performa", padding=10)
        chart_wrap.pack(fill="both", expand=True)

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.figure import Figure

        self.fig = Figure(figsize=(8, 4), dpi=100)
        ax = self.fig.add_subplot(111)
        ax.set_title("Achievement % per Tanggal")
        ax.set_xlabel("Tanggal")
        ax.set_ylabel("Achievement %")

        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_wrap)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def build_entry_tab(self):
        form = ttk.Frame(self.tab_entry)
        form.pack(fill="both", expand=True)

        role = self.current_user["role"]
        left_title = "Kinerja Harian" if role == "admin" else ("Kinerja Harian Tim" if role == "supervisor" else "Kinerja Harian Saya")
        right_title = "Lembur" if role == "admin" else ("Lembur Tim" if role == "supervisor" else "Lembur Saya")
        left = ttk.LabelFrame(form, text=left_title, padding=12)
        right = ttk.LabelFrame(form, text=right_title, padding=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right.grid(row=0, column=1, sticky="nsew")

        left_fields = [
            ("Nama Karyawan", self.employee_name_var),
            # ("Departemen", self.department_var),
            ("Supervisor", self.supervisor_name_var),
            ("Tanggal Kinerja", self.work_date_var),
        ]
        for i, (label, var) in enumerate(left_fields):
            ttk.Label(left, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 8))
            state = "normal"
            if self.current_user["role"] == "karyawan" and label in ("Nama Karyawan", "Departemen", "Supervisor"):
                state = "readonly"
            ttk.Entry(left, textvariable=var, width=38, state=state).grid(row=i, column=1, sticky="ew", pady=5)

        ttk.Label(left, text="Judul Proyek").grid(row=4, column=0, sticky="w", pady=5, padx=(0, 8))
        ttk.Entry(left, textvariable=self.project_title_var, width=38).grid(row=4, column=1, sticky="ew", pady=5)

        ttk.Label(left, text="Penjelasan Target").grid(row=5, column=0, sticky="w", pady=5, padx=(0, 8))
        ttk.Entry(left, textvariable=self.target_description_var, width=38).grid(row=5, column=1, sticky="ew", pady=5)

        ttk.Label(left, text="Kategori Target").grid(row=6, column=0, sticky="w", pady=5, padx=(0, 8))
        ttk.Combobox(left, textvariable=self.target_category_var, values=TARGET_CATEGORIES, state="readonly", width=35).grid(row=6, column=1, sticky="ew", pady=5)

        ttk.Label(left, text="Target Harian").grid(row=7, column=0, sticky="w", pady=5, padx=(0, 8))
        ttk.Entry(left, textvariable=self.daily_target_var, width=38).grid(row=7, column=1, sticky="ew", pady=5)

        ttk.Label(left, text="Pencapaian Harian").grid(row=8, column=0, sticky="w", pady=5, padx=(0, 8))
        ttk.Entry(left, textvariable=self.daily_performance_var, width=38).grid(row=8, column=1, sticky="ew", pady=5)

        ttk.Label(left, text="% Pencapaian").grid(row=9, column=0, sticky="w", pady=5, padx=(0, 8))
        ttk.Entry(left, textvariable=self.achievement_percent_var, width=38, state="readonly").grid(row=9, column=1, sticky="ew", pady=5)

        right_fields = [
            ("Tanggal Lembur", self.overtime_date_var),
            ("Jam Mulai", self.start_time_var),
            ("Jam Selesai", self.end_time_var),
            ("Durasi Lembur (jam)", self.duration_var),
            ("Alasan Lembur", self.reason_var),
            ("Komentar", self.comments_var),
        ]
        for i, (label, var) in enumerate(right_fields):
            ttk.Label(right, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 8))
            state = "readonly" if "Durasi" in label else "normal"
            ttk.Entry(right, textvariable=var, width=38, state=state).grid(row=i, column=1, sticky="ew", pady=5)

        ttk.Label(right, text="Approval").grid(row=6, column=0, sticky="w", pady=5, padx=(0, 8))
        approval_state = "readonly" if self.current_user["role"] in ("admin", "supervisor") else "disabled"
        ttk.Combobox(right, textvariable=self.approval_var, values=["Pending", "Approved", "Rejected"], state=approval_state, width=35).grid(row=6, column=1, sticky="ew", pady=5)

        btns = ttk.Frame(self.tab_entry)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Hitung % Kinerja", command=self.compute_achievement).pack(side="left", padx=4)
        ttk.Button(btns, text="Hitung Durasi", command=self.compute_duration).pack(side="left", padx=4)
        ttk.Button(btns, text="Simpan", command=self.save_new_record).pack(side="left", padx=4)
        ttk.Button(btns, text="Update", command=self.update_record).pack(side="left", padx=4)
        ttk.Button(btns, text="Hapus", command=self.delete_record).pack(side="left", padx=4)
        ttk.Button(btns, text="Reset Form", command=self.clear_form).pack(side="left", padx=4)

        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        left.columnconfigure(1, weight=1)
        right.columnconfigure(1, weight=1)

    def build_records_tab(self):
        top = ttk.Frame(self.tab_records)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Nama").pack(side="left")
        ttk.Entry(top, textvariable=self.search_name_var, width=22).pack(side="left", padx=6)
        ttk.Label(top, text="Tanggal").pack(side="left")
        ttk.Entry(top, textvariable=self.search_date_var, width=14).pack(side="left", padx=6)
        ttk.Button(top, text="Filter", command=self.refresh_records_table).pack(side="left", padx=4)
        ttk.Button(top, text="Reset", command=self.reset_filters).pack(side="left", padx=4)

        wrap = ttk.Frame(self.tab_records)
        wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(wrap, columns=FIELDNAMES, show="headings", height=18)
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        headings = {c: c for c in FIELDNAMES}
        headings.update({
            "record_id": "ID", "employee_name": "Nama",  "role_scope": "Role",
            "supervisor_name": "Supervisor", "work_date": "Tgl Kinerja", "project_title": "Judul Proyek",
            "target_description": "Penjelasan Target", "target_category": "Kategori",
            "daily_target": "Target", "daily_performance": "Pencapaian", "achievement_percent": "% Capaian",
            "overtime_date": "Tgl Lembur", "start_time": "Mulai", "end_time": "Selesai", "duration_hours": "Jam",
            "reason": "Alasan", "manager_approval": "Approval", "comments": "Komentar", "created_by": "Input Oleh",
            "created_at": "Created", "updated_at": "Updated",
        })
        widths = {c: 100 for c in FIELDNAMES}
        widths.update({
            "record_id": 50, "employee_name": 130, "department": 110, "role_scope": 80, "supervisor_name": 110,
            "work_date": 90, "project_title": 170, "target_description": 220, "target_category": 100, "daily_target": 75, "daily_performance": 90, "achievement_percent": 85,
            "overtime_date": 90, "start_time": 65, "end_time": 65, "duration_hours": 65, "reason": 150, "manager_approval": 80,
            "comments": 150, "created_by": 90, "created_at": 120, "updated_at": 120,
        })
        for c in FIELDNAMES:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

    def build_recap_tab(self):
        top = ttk.Frame(self.tab_recap)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Bulan (YYYY-MM)").pack(side="left")
        ttk.Entry(top, textvariable=self.filter_month_var, width=12).pack(side="left", padx=6)
        ttk.Button(top, text="Refresh Rekap", command=self.refresh_recap_table).pack(side="left", padx=4)
        ttk.Button(top, text="Export Excel", command=self.export_excel).pack(side="left", padx=4)
        ttk.Button(top, text="Export PDF", command=self.export_pdf).pack(side="left", padx=4)

        wrap = ttk.Frame(self.tab_recap)
        wrap.pack(fill="both", expand=True)
        cols = ["employee_name", "department", "records_count", "total_overtime_hours", "avg_achievement_percent", "approved_count"]
        self.recap_tree = ttk.Treeview(wrap, columns=cols, show="headings", height=16)
        self.recap_tree.pack(side="left", fill="both", expand=True)
        labels = {
            "employee_name": "Nama",  "records_count": "Jumlah Data",
            "total_overtime_hours": "Total Jam Lembur", "avg_achievement_percent": "Rata-rata Capaian %", "approved_count": "Approved"
        }
        widths = {"employee_name": 180, "department": 140, "records_count": 100, "total_overtime_hours": 140, "avg_achievement_percent": 160, "approved_count": 90}
        for c in cols:
            self.recap_tree.heading(c, text=labels[c])
            self.recap_tree.column(c, width=widths[c], anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.recap_tree.yview)
        sb.pack(side="right", fill="y")
        self.recap_tree.configure(yscrollcommand=sb.set)

    def build_attendance_tab(self):
        top = ttk.Frame(self.tab_attendance)
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="Check In Hari Ini", command=self.check_in).pack(side="left", padx=4)
        ttk.Button(top, text="Check Out Hari Ini", command=self.check_out).pack(side="left", padx=4)

        wrap = ttk.Frame(self.tab_attendance)
        wrap.pack(fill="both", expand=True)
        self.att_tree = ttk.Treeview(wrap, columns=ATTENDANCE_FIELDS, show="headings", height=16)
        self.att_tree.pack(side="left", fill="both", expand=True)
        labels = {
            "attendance_id": "ID", "employee_name": "Nama",  "attendance_date": "Tanggal",
            "check_in_time": "Check In", "check_out_time": "Check Out", "status": "Status", "notes": "Catatan", "created_at": "Created"
        }
        widths = {
            "attendance_id": 50, "employee_name": 150, "department": 120, "attendance_date": 95,
            "check_in_time": 90, "check_out_time": 90, "status": 100, "notes": 180, "created_at": 130
        }
        for c in ATTENDANCE_FIELDS:
            self.att_tree.heading(c, text=labels[c])
            self.att_tree.column(c, width=widths[c], anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.att_tree.yview)
        sb.pack(side="right", fill="y")
        self.att_tree.configure(yscrollcommand=sb.set)

    def build_users_tab(self):
        outer = ttk.Frame(self.tab_users)
        outer.pack(fill="both", expand=True)

        left = ttk.LabelFrame(outer, text="Pending Registrations", padding=10)
        right = ttk.LabelFrame(outer, text="Users", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right.grid(row=0, column=1, sticky="nsew")

        self.pending_tree = ttk.Treeview(left, columns=["username", "employee_name", "department", "supervisor_name", "requested_at"], show="headings", height=14)
        self.pending_tree.pack(fill="both", expand=True)
        for c, w, t in [
            ("username", 110, "Username"),
            ("employee_name", 140, "Nama"),
            ("department", 110, "Departemen"),
            ("supervisor_name", 120, "Supervisor"),
            ("requested_at", 130, "Requested"),
        ]:
            self.pending_tree.heading(c, text=t)
            self.pending_tree.column(c, width=w, anchor="w")
        ttk.Button(left, text="Approve Selected", command=self.approve_pending_user).pack(fill="x", pady=(8, 0))
        ttk.Button(left, text="Reject Selected", command=self.reject_pending_user).pack(fill="x", pady=(6, 0))

        top = ttk.Frame(right)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Filter Role").pack(side="left")
        ttk.Combobox(top, textvariable=self.user_role_filter_var, values=["Semua", "admin", "supervisor", "karyawan"], state="readonly", width=12).pack(side="left", padx=6)
        ttk.Button(top, text="Refresh", command=self.refresh_users_tables).pack(side="left")
        self.users_tree = ttk.Treeview(right, columns=["username", "employee_name", "role", "department", "supervisor_name", "is_active"], show="headings", height=14)
        self.users_tree.pack(fill="both", expand=True)
        for c, w, t in [
            ("username", 110, "Username"),
            ("employee_name", 140, "Nama"),
            ("role", 90, "Role"),
            ("department", 120, "Departemen"),
            ("supervisor_name", 120, "Supervisor"),
            ("is_active", 70, "Active"),
        ]:
            self.users_tree.heading(c, text=t)
            self.users_tree.column(c, width=w, anchor="w")
        ttk.Button(right, text="Toggle Active", command=self.toggle_user_active).pack(fill="x", pady=(8, 0))
        ttk.Button(right, text="Reset Password User", command=self.admin_reset_password).pack(fill="x", pady=(6, 0))

        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)

    def apply_user_defaults(self):
        if self.current_user["role"] == "karyawan":
            self.employee_name_var.set(self.current_user["employee_name"])
            self.department_var.set(self.current_user["department"])
            self.supervisor_name_var.set(self.current_user.get("supervisor_name", ""))
        elif self.current_user["role"] == "supervisor":
            self.supervisor_name_var.set(self.current_user["employee_name"])

    def set_status(self, text):
        self.status_label.config(text=text)

    def update_mode_badge(self):
        self.mode_badge.config(text=f"Mode: {'Google Sheets' if self.settings.get('mode') == 'google_sheets' else 'Local CSV'}")

    def visible_records(self):
        rows = self.records
        role = self.current_user["role"]
        if role == "karyawan":
            rows = [r for r in rows if r.get("employee_name") == self.current_user["employee_name"]]
        elif role == "supervisor":
            supervisor = self.current_user["employee_name"]
            rows = [r for r in rows if r.get("supervisor_name") == supervisor or r.get("employee_name") == supervisor]
        return rows

    def visible_attendance(self):
        rows = self.attendance
        role = self.current_user["role"]
        if role == "karyawan":
            rows = [r for r in rows if r.get("employee_name") == self.current_user["employee_name"]]
        elif role == "supervisor":
            team_names = set([u["employee_name"] for u in self.users if u.get("supervisor_name") == self.current_user["employee_name"]])
            team_names.add(self.current_user["employee_name"])
            rows = [r for r in rows if r.get("employee_name") in team_names]
        return rows

    def compute_achievement(self):
        try:
            target = float(self.daily_target_var.get().strip())
            performance = float(self.daily_performance_var.get().strip())
            pct = calculate_achievement_percent(target, performance)
            self.achievement_percent_var.set(f"{pct:.2f}")
            self.set_status(f"Persentase pencapaian {pct:.2f}%.")
        except Exception:
            messagebox.showerror("Input salah", "Target harian dan pencapaian harian harus berupa angka.")

    def compute_duration(self):
        try:
            duration = calculate_duration(self.start_time_var.get().strip(), self.end_time_var.get().strip())
            self.duration_var.set(f"{duration:.2f}")
            self.set_status(f"Durasi lembur {duration:.2f} jam.")
        except Exception:
            messagebox.showerror("Input salah", "Jam mulai dan selesai harus berformat HH:MM.")

    def validate_form(self):
        required = {
            "Nama Karyawan": self.employee_name_var.get().strip(),
            
            "Supervisor": self.supervisor_name_var.get().strip(),
            "Tanggal Kinerja": self.work_date_var.get().strip(),
            "Judul Proyek": self.project_title_var.get().strip(),
            "Penjelasan Target": self.target_description_var.get().strip(),
            "Target Harian": self.daily_target_var.get().strip(),
            "Pencapaian Harian": self.daily_performance_var.get().strip(),
            "Tanggal Lembur": self.overtime_date_var.get().strip(),
            "Jam Mulai": self.start_time_var.get().strip(),
            "Jam Selesai": self.end_time_var.get().strip(),
            "Alasan Lembur": self.reason_var.get().strip(),
        }
        missing = [k for k, v in required.items() if not v]
        if missing:
            raise ValueError("Kolom wajib belum terisi: " + ", ".join(missing))

        parse_date(self.work_date_var.get().strip())
        parse_date(self.overtime_date_var.get().strip())
        parse_time(self.start_time_var.get().strip())
        parse_time(self.end_time_var.get().strip())

        target = float(self.daily_target_var.get().strip())
        performance = float(self.daily_performance_var.get().strip())
        achievement = calculate_achievement_percent(target, performance)
        duration = calculate_duration(self.start_time_var.get().strip(), self.end_time_var.get().strip())

        self.achievement_percent_var.set(f"{achievement:.2f}")
        self.duration_var.set(f"{duration:.2f}")
        return target, performance, achievement, duration

    def build_record_from_form(self, record_id, existing=None):
        target, performance, achievement, duration = self.validate_form()
        approval = self.approval_var.get().strip() or "Pending"
        if self.current_user["role"] == "karyawan":
            approval = existing.get("manager_approval", "Pending") if existing else "Pending"

        created_at = existing["created_at"] if existing else now_str()
        created_by = existing["created_by"] if existing else self.current_user["employee_name"]

        return {
            "record_id": record_id,
            "employee_name": self.employee_name_var.get().strip(),
            "department": self.department_var.get().strip(),
            "role_scope": self.current_user["role"],
            "supervisor_name": self.supervisor_name_var.get().strip(),
            "work_date": self.work_date_var.get().strip(),
            "project_title": self.project_title_var.get().strip(),
            "target_description": self.target_description_var.get().strip(),
            "target_category": self.target_category_var.get().strip(),
            "daily_target": f"{target:.2f}",
            "daily_performance": f"{performance:.2f}",
            "achievement_percent": f"{achievement:.2f}",
            "overtime_date": self.overtime_date_var.get().strip(),
            "start_time": self.start_time_var.get().strip(),
            "end_time": self.end_time_var.get().strip(),
            "duration_hours": f"{duration:.2f}",
            "reason": self.reason_var.get().strip(),
            "manager_approval": approval,
            "comments": self.comments_var.get().strip(),
            "created_by": created_by,
            "created_at": created_at,
            "updated_at": now_str(),
        }

    def warning_text(self, record, exclude_record_id=None):
        warns = []
        duration = float(record["duration_hours"])
        if duration > 4:
            warns.append("Durasi lembur harian melebihi 4 jam.")
        weekly_total = 0.0
        week_start = monday_of_week(record["overtime_date"]).date()
        for r in self.records:
            if exclude_record_id and r["record_id"] == exclude_record_id:
                continue
            if r.get("employee_name") != record["employee_name"]:
                continue
            try:
                if monday_of_week(r["overtime_date"]).date() == week_start:
                    weekly_total += float(r.get("duration_hours", 0) or 0)
            except Exception:
                pass
        weekly_total += duration
        if weekly_total > 18:
            warns.append(f"Total lembur mingguan menjadi {weekly_total:.2f} jam, melebihi 18 jam.")
        if float(record["achievement_percent"]) < 100:
            warns.append(f"Capaian harian baru {record['achievement_percent']}% dari target.")
        return warns

    def can_edit_record(self, record):
        if self.current_user["role"] == "admin":
            return True
        if self.current_user["role"] == "supervisor":
            return record.get("supervisor_name") == self.current_user["employee_name"] or record.get("employee_name") == self.current_user["employee_name"]
        if self.current_user["role"] == "karyawan":
            return record.get("employee_name") == self.current_user["employee_name"] and record.get("manager_approval") != "Approved"
        return False

    def save_new_record(self):
        try:
            record = self.build_record_from_form(next_id(self.records, "record_id"))
        except Exception as e:
            messagebox.showerror("Validasi gagal", str(e))
            return
        if is_duplicate_record(self.records, record):
            messagebox.showwarning("Duplikasi terdeteksi", "Data yang sama persis sudah ada. Multi target per hari diperbolehkan, tetapi kombinasi proyek, target, jam lembur, dan alasan yang identik akan ditolak.")
            return
        warns = self.warning_text(record)
        if warns and not messagebox.askyesno("Peringatan", "\n".join(warns) + "\n\nTetap simpan?"):
            return
        append_csv(DATA_FILE, FIELDNAMES, record)
        self.records.append(record)
        self.maybe_auto_sync()
        self.refresh_all()
        self.clear_form()
        self.set_status("Data berhasil disimpan.")

    def selected_record(self):
        if not self.selected_record_id:
            return None
        for r in self.records:
            if r["record_id"] == self.selected_record_id:
                return r
        return None

    def update_record(self):
        existing = self.selected_record()
        if not existing:
            messagebox.showwarning("Belum memilih data", "Pilih data dulu pada tab Data.")
            return
        if not self.can_edit_record(existing):
            messagebox.showwarning("Akses ditolak", "Anda tidak punya hak update untuk data ini.")
            return
        try:
            updated = self.build_record_from_form(existing["record_id"], existing)
        except Exception as e:
            messagebox.showerror("Validasi gagal", str(e))
            return
        if is_duplicate_record(self.records, updated, exclude_record_id=existing["record_id"]):
            messagebox.showwarning("Duplikasi terdeteksi", "Update dibatalkan karena menghasilkan data duplikat.")
            return
        warns = self.warning_text(updated, exclude_record_id=existing["record_id"])
        if warns and not messagebox.askyesno("Peringatan", "\n".join(warns) + "\n\nTetap update?"):
            return
        for i, row in enumerate(self.records):
            if row["record_id"] == existing["record_id"]:
                self.records[i] = updated
                break
        save_csv(DATA_FILE, FIELDNAMES, self.records)
        self.maybe_auto_sync()
        self.refresh_all()
        self.set_status("Data berhasil diupdate.")

    def delete_record(self):
        existing = self.selected_record()
        if not existing:
            messagebox.showwarning("Belum memilih data", "Pilih data dulu pada tab Data.")
            return
        if not self.can_edit_record(existing):
            messagebox.showwarning("Akses ditolak", "Anda tidak punya hak hapus untuk data ini.")
            return
        if not messagebox.askyesno("Konfirmasi", f"Hapus data ID {existing['record_id']}?"):
            return
        self.records = [r for r in self.records if r["record_id"] != existing["record_id"]]
        save_csv(DATA_FILE, FIELDNAMES, self.records)
        self.maybe_auto_sync()
        self.refresh_all()
        self.clear_form()
        self.set_status("Data berhasil dihapus.")

    def clear_form(self):
        if self.current_user["role"] == "karyawan":
            self.employee_name_var.set(self.current_user["employee_name"])
            self.department_var.set(self.current_user["department"])
            self.supervisor_name_var.set(self.current_user.get("supervisor_name", ""))
        elif self.current_user["role"] == "supervisor":
            self.employee_name_var.set("")
            self.department_var.set(self.current_user["department"])
            self.supervisor_name_var.set(self.current_user["employee_name"])
        else:
            self.employee_name_var.set("")
            self.department_var.set("")
            self.supervisor_name_var.set("")
        self.work_date_var.set(today_str())
        self.project_title_var.set("")
        self.target_description_var.set("")
        self.target_category_var.set(TARGET_CATEGORIES[0])
        self.daily_target_var.set("")
        self.daily_performance_var.set("")
        self.achievement_percent_var.set("0.00")
        self.overtime_date_var.set(today_str())
        self.start_time_var.set("")
        self.end_time_var.set("")
        self.duration_var.set("0.00")
        self.reason_var.set("")
        self.approval_var.set("Pending")
        self.comments_var.set("")
        self.selected_record_id = None

    def reset_filters(self):
        self.search_name_var.set("")
        self.search_date_var.set("")
        self.refresh_records_table()

    def refresh_chart(self, rows):
        ax = self.fig.axes[0]
        ax.clear()
        ax.set_title("Achievement % per Tanggal")
        ax.set_xlabel("Tanggal")
        ax.set_ylabel("Achievement %")
        rows = sorted(rows, key=lambda x: x.get("work_date", ""))
        xs = [r.get("work_date", "")[-2:] for r in rows]
        ys = [float(r.get("achievement_percent", 0) or 0) for r in rows]
        if xs and ys:
            ax.plot(xs, ys, marker="o")
            ax.axhline(100, linestyle="--")
        self.canvas.draw()

    def refresh_dashboard(self):
        rows = self.visible_records()
        month = datetime.now().strftime("%Y-%m")
        rows_month = [r for r in rows if r.get("work_date", "").startswith(month) or r.get("overtime_date", "").startswith(month)]
        total_hours = sum(float(r.get("duration_hours", 0) or 0) for r in rows_month)
        avg_achievement = sum(float(r.get("achievement_percent", 0) or 0) for r in rows_month) / len(rows_month) if rows_month else 0
        pending = sum(1 for r in rows if r.get("manager_approval") == "Pending")
        employee_count = len(set(r.get("employee_name", "") for r in rows if r.get("employee_name")))

        att_today = [r for r in self.visible_attendance() if r.get("attendance_date") == today_str()]
        present_today = sum(1 for r in att_today if r.get("status") in ("Checked In", "Checked Out"))

        first_label = list(self.dashboard_cards.keys())[0]
        second_label = list(self.dashboard_cards.keys())[1]
        third_label = list(self.dashboard_cards.keys())[2]
        fourth_label = list(self.dashboard_cards.keys())[3]
        fifth_label = list(self.dashboard_cards.keys())[4]
        self.dashboard_cards[first_label].config(text=f"{total_hours:.2f}")
        self.dashboard_cards[second_label].config(text=f"{avg_achievement:.2f}%")
        self.dashboard_cards[third_label].config(text=str(pending))
        self.dashboard_cards[fourth_label].config(text=str(employee_count))
        self.dashboard_cards[fifth_label].config(text=str(present_today))
        self.refresh_chart(rows_month)

    def filtered_visible_records(self):
        rows = self.visible_records()
        q_name = self.search_name_var.get().strip().lower()
        q_date = self.search_date_var.get().strip()
        out = []
        for r in rows:
            ok = True
            if q_name and q_name not in r.get("employee_name", "").lower():
                ok = False
            if q_date and q_date not in (r.get("work_date", ""), r.get("overtime_date", "")):
                ok = False
            if ok:
                out.append(r)
        return out

    def refresh_records_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = self.filtered_visible_records()
        for r in rows:
            self.tree.insert("", "end", values=[r.get(c, "") for c in FIELDNAMES])
        self.set_status(f"Menampilkan {len(rows)} data.")

    def recap_rows(self):
        month = self.filter_month_var.get().strip()
        rows = self.visible_records()
        if month:
            rows = [r for r in rows if r.get("work_date", "").startswith(month) or r.get("overtime_date", "").startswith(month)]
        grouped = {}
        for r in rows:
            key = (r.get("employee_name", ""), r.get("department", ""))
            if key not in grouped:
                grouped[key] = {
                    "employee_name": key[0],
                    "department": key[1],
                    "records_count": 0,
                    "total_overtime_hours": 0.0,
                    "achievement_sum": 0.0,
                    "approved_count": 0,
                }
            grouped[key]["records_count"] += 1
            grouped[key]["total_overtime_hours"] += float(r.get("duration_hours", 0) or 0)
            grouped[key]["achievement_sum"] += float(r.get("achievement_percent", 0) or 0)
            if r.get("manager_approval") == "Approved":
                grouped[key]["approved_count"] += 1
        out = []
        for g in grouped.values():
            count = g["records_count"] or 1
            out.append({
                "employee_name": g["employee_name"],
                "department": g["department"],
                "records_count": g["records_count"],
                "total_overtime_hours": round(g["total_overtime_hours"], 2),
                "avg_achievement_percent": round(g["achievement_sum"] / count, 2),
                "approved_count": g["approved_count"],
            })
        return sorted(out, key=lambda x: (-x["avg_achievement_percent"], -x["total_overtime_hours"], x["employee_name"]))

    def refresh_recap_table(self):
        for item in self.recap_tree.get_children():
            self.recap_tree.delete(item)
        for row in self.recap_rows():
            self.recap_tree.insert("", "end", values=[
                row["employee_name"], row["department"], row["records_count"],
                row["total_overtime_hours"], row["avg_achievement_percent"], row["approved_count"]
            ])

    def refresh_attendance_table(self):
        for item in self.att_tree.get_children():
            self.att_tree.delete(item)
        for row in self.visible_attendance():
            self.att_tree.insert("", "end", values=[row.get(c, "") for c in ATTENDANCE_FIELDS])

    def refresh_users_tables(self):
        if not self.tab_users:
            return
        self.pending_users = load_json(PENDING_USERS_FILE, [])
        self.users = load_json(USERS_FILE, DEFAULT_USERS)

        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)
        for row in self.pending_users:
            self.pending_tree.insert("", "end", values=[
                row.get("username", ""), row.get("employee_name", ""), row.get("department", ""),
                row.get("supervisor_name", ""), row.get("requested_at", "")
            ])

        for item in self.users_tree.get_children():
            self.users_tree.delete(item)
        role_filter = self.user_role_filter_var.get()
        for row in self.users:
            if role_filter != "Semua" and row.get("role") != role_filter:
                continue
            self.users_tree.insert("", "end", values=[
                row.get("username", ""), row.get("employee_name", ""), row.get("role", ""),
                row.get("department", ""), row.get("supervisor_name", ""), str(row.get("is_active", True))
            ])

    def refresh_all(self):
        self.records = load_csv(DATA_FILE, FIELDNAMES)
        self.attendance = load_csv(ATTENDANCE_FILE, ATTENDANCE_FIELDS)
        self.refresh_dashboard()
        self.refresh_records_table()
        self.refresh_recap_table()
        self.refresh_attendance_table()
        self.refresh_users_tables()

    def on_row_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        row = dict(zip(FIELDNAMES, values))
        self.selected_record_id = row["record_id"]
        self.employee_name_var.set(row["employee_name"])
        self.department_var.set(row["department"])
        self.supervisor_name_var.set(row["supervisor_name"])
        self.work_date_var.set(row["work_date"])
        self.project_title_var.set(row.get("project_title", ""))
        self.target_description_var.set(row.get("target_description", ""))
        self.target_category_var.set(row["target_category"] or TARGET_CATEGORIES[0])
        self.daily_target_var.set(row["daily_target"])
        self.daily_performance_var.set(row["daily_performance"])
        self.achievement_percent_var.set(row["achievement_percent"])
        self.overtime_date_var.set(row["overtime_date"])
        self.start_time_var.set(row["start_time"])
        self.end_time_var.set(row["end_time"])
        self.duration_var.set(row["duration_hours"])
        self.reason_var.set(row["reason"])
        self.approval_var.set(row["manager_approval"] or "Pending")
        self.comments_var.set(row["comments"])
        self.notebook.select(self.tab_entry)

    def check_in(self):
        employee = self.current_user["employee_name"]
        today = today_str()
        existing = None
        for r in self.attendance:
            if r["employee_name"] == employee and r["attendance_date"] == today:
                existing = r
                break
        if existing and existing.get("check_in_time"):
            messagebox.showwarning("Sudah check in", "Anda sudah check in hari ini.")
            return
        if not existing:
            if attendance_exists(self.attendance, employee, today):
                messagebox.showwarning("Duplikasi absensi", "Absensi hari ini sudah ada.")
                return
            existing = {
                "attendance_id": next_id(self.attendance, "attendance_id"),
                "employee_name": employee,
                "department": self.current_user["department"],
                "attendance_date": today,
                "check_in_time": datetime.now().strftime("%H:%M"),
                "check_out_time": "",
                "status": "Checked In",
                "notes": "",
                "created_at": now_str(),
            }
            self.attendance.append(existing)
        else:
            existing["check_in_time"] = datetime.now().strftime("%H:%M")
            existing["status"] = "Checked In"
        save_csv(ATTENDANCE_FILE, ATTENDANCE_FIELDS, self.attendance)
        self.refresh_attendance_table()
        self.refresh_dashboard()
        self.set_status("Check in berhasil.")

    def check_out(self):
        employee = self.current_user["employee_name"]
        today = today_str()
        existing = None
        for r in self.attendance:
            if r["employee_name"] == employee and r["attendance_date"] == today:
                existing = r
                break
        if not existing or not existing.get("check_in_time"):
            messagebox.showwarning("Belum check in", "Silakan check in dulu.")
            return
        if existing.get("check_out_time"):
            messagebox.showwarning("Sudah check out", "Anda sudah check out hari ini.")
            return
        existing["check_out_time"] = datetime.now().strftime("%H:%M")
        existing["status"] = "Checked Out"
        save_csv(ATTENDANCE_FILE, ATTENDANCE_FIELDS, self.attendance)
        self.refresh_attendance_table()
        self.refresh_dashboard()
        self.set_status("Check out berhasil.")

    def approve_pending_user(self):
        sel = self.pending_tree.selection()
        if not sel:
            messagebox.showwarning("Belum memilih", "Pilih registrasi pending dulu.")
            return
        username = self.pending_tree.item(sel[0], "values")[0]
        idx = next((i for i, u in enumerate(self.pending_users) if u["username"] == username), None)
        if idx is None:
            return
        user = self.pending_users.pop(idx)
        user["is_active"] = True
        user.pop("requested_at", None)
        self.users.append(user)
        save_json(USERS_FILE, self.users)
        save_json(PENDING_USERS_FILE, self.pending_users)
        self.refresh_users_tables()
        self.set_status(f"User {username} berhasil diapprove.")

    def reject_pending_user(self):
        sel = self.pending_tree.selection()
        if not sel:
            messagebox.showwarning("Belum memilih", "Pilih registrasi pending dulu.")
            return
        username = self.pending_tree.item(sel[0], "values")[0]
        self.pending_users = [u for u in self.pending_users if u["username"] != username]
        save_json(PENDING_USERS_FILE, self.pending_users)
        self.refresh_users_tables()
        self.set_status(f"User {username} ditolak.")

    def toggle_user_active(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showwarning("Belum memilih", "Pilih user dulu.")
            return
        username = self.users_tree.item(sel[0], "values")[0]
        for user in self.users:
            if user["username"] == username:
                user["is_active"] = not bool(user.get("is_active", True))
                break
        save_json(USERS_FILE, self.users)
        self.refresh_users_tables()
        self.set_status(f"Status user {username} diubah.")

    def export_excel(self):
        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror("Library belum ada", "Pandas belum terpasang.")
            return

        path = filedialog.asksaveasfilename(title="Simpan Excel", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        detail = self.filtered_visible_records()
        recap = self.recap_rows()
        ranking = sorted(recap, key=lambda x: (-x["avg_achievement_percent"], -x["total_overtime_hours"]))
        attendance = self.visible_attendance()
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            pd.DataFrame(detail).to_excel(writer, index=False, sheet_name="DetailData")
            pd.DataFrame(recap).to_excel(writer, index=False, sheet_name="RekapBulanan")
            pd.DataFrame(ranking).to_excel(writer, index=False, sheet_name="Ranking")
            pd.DataFrame(attendance).to_excel(writer, index=False, sheet_name="Attendance")
        messagebox.showinfo("Sukses", f"Excel tersimpan:\n{path}")

    def export_pdf(self):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet
        except ImportError:
            messagebox.showerror("Library belum ada", "ReportLab belum terpasang.")
            return
        path = filedialog.asksaveasfilename(title="Simpan PDF", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        recap = self.recap_rows()
        doc = SimpleDocTemplate(path, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        data = [["Nama", "Departemen", "Jumlah Data", "Total Jam Lembur", "Rata-rata Capaian %", "Approved"]]
        for row in recap:
            data.append([
                row["employee_name"], row["department"], row["records_count"],
                row["total_overtime_hours"], row["avg_achievement_percent"], row["approved_count"]
            ])
        table = Table(data)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B5D7A")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ]))
        elements = [
            Paragraph("PT. Sagamartha Ultima Indonesia", styles["Title"]),
            Paragraph("Rekap Bulanan Lembur dan Kinerja", styles["Heading2"]),
            Spacer(1, 12),
            table,
        ]
        doc.build(elements)
        messagebox.showinfo("Sukses", f"PDF tersimpan:\n{path}")



    def change_password_force(self):
        self.change_password(force=True)

    def change_password(self, force=False):
        win = tk.Toplevel(self)
        win.title("Ubah Password")
        win.geometry("420x240")
        win.resizable(False, False)

        old_var = tk.StringVar()
        new_var = tk.StringVar()
        confirm_var = tk.StringVar()

        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Password Lama").grid(row=0, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(frm, textvariable=old_var, show="*", width=28).grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(frm, text="Password Baru").grid(row=1, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(frm, textvariable=new_var, show="*", width=28).grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(frm, text="Konfirmasi Password").grid(row=2, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(frm, textvariable=confirm_var, show="*", width=28).grid(row=2, column=1, sticky="ew", pady=6)

        note = "Password minimal 8 karakter."
        if force:
            note = "Anda wajib mengganti password sebelum melanjutkan. Password minimal 8 karakter."
        ttk.Label(frm, text=note, justify="left").grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 12))

        def save_pw():
            old_pw = old_var.get().strip()
            new_pw = new_var.get().strip()
            conf_pw = confirm_var.get().strip()

            if len(new_pw) < 8:
                messagebox.showerror("Validasi", "Password baru minimal 8 karakter.")
                return
            if new_pw != conf_pw:
                messagebox.showerror("Validasi", "Konfirmasi password tidak sama.")
                return

            users = load_json(USERS_FILE, DEFAULT_USERS)
            users, changed = normalize_users_security(users)
            if changed:
                save_json(USERS_FILE, users)

            for user in users:
                if user["username"] == self.current_user["username"]:
                    if not verify_password(user["password"], old_pw):
                        messagebox.showerror("Validasi", "Password lama salah.")
                        return
                    user["password"] = hash_password(new_pw)
                    user["must_change_password"] = False
                    save_json(USERS_FILE, users)
                    self.current_user = user
                    messagebox.showinfo("Berhasil", "Password berhasil diubah.")
                    win.destroy()
                    return

        ttk.Button(frm, text="Simpan Password", command=save_pw).grid(row=4, column=0, columnspan=2, sticky="ew")
        frm.columnconfigure(1, weight=1)
        if force:
            win.protocol("WM_DELETE_WINDOW", lambda: None)
        win.grab_set()

    def admin_reset_password(self):
        if self.current_user["role"] != "admin":
            return
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showwarning("Belum memilih", "Pilih user dulu.")
            return
        username = self.users_tree.item(sel[0], "values")[0]
        if username == self.current_user["username"]:
            messagebox.showwarning("Ditolak", "Gunakan menu Ubah Password untuk akun sendiri.")
            return
        users = load_json(USERS_FILE, DEFAULT_USERS)
        users, changed = normalize_users_security(users)
        for user in users:
            if user["username"] == username:
                user["password"] = hash_password("12345678")
                user["must_change_password"] = True
                break
        save_json(USERS_FILE, users)
        self.users = users
        self.refresh_users_tables()
        messagebox.showinfo("Berhasil", f"Password user {username} direset ke 12345678 dan wajib diganti saat login pertama.")

    def auto_create_google_sheet(self):
        if self.current_user["role"] != "admin":
            messagebox.showwarning("Akses ditolak", "Hanya admin yang dapat membuat Google Sheet baru.")
            return
        if not self.settings.get("google_credentials_file"):
            messagebox.showwarning("Credentials belum diisi", "Isi dulu credentials JSON di Pengaturan Sync.")
            return

        win = tk.Toplevel(self)
        win.title("Auto Generate Google Sheet")
        win.geometry("520x220")
        win.resizable(False, False)

        title_var = tk.StringVar(value="Sagamartha HR Enterprise Data")
        share_var = tk.StringVar()

        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Judul Spreadsheet").grid(row=0, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(frm, textvariable=title_var, width=40).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Label(frm, text="Share ke email Google (opsional)").grid(row=1, column=0, sticky="w", pady=6, padx=(0, 8))
        ttk.Entry(frm, textvariable=share_var, width=40).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Label(frm, text="Spreadsheet dibuat otomatis via Sheets API.\nSpreadsheet ID langsung disimpan ke pengaturan aplikasi.", justify="left").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 12))

        def do_create():
            try:
                client = GoogleSheetsSync(
                    self.settings.get("google_credentials_file", ""),
                    self.settings.get("google_spreadsheet_id", "")
                )
                client.connect()
                spreadsheet = client.create_spreadsheet(title_var.get().strip() or "Sagamartha HR Enterprise Data")
                self.settings["google_spreadsheet_id"] = client.spreadsheet_id
                self.settings["mode"] = "google_sheets"
                save_settings(self.settings)
                self.update_mode_badge()
                share_email = share_var.get().strip()
                if share_email:
                    client.share_spreadsheet(share_email, role="writer")
                url = spreadsheet.get("spreadsheetUrl", "")
                messagebox.showinfo("Berhasil", f"Google Sheet berhasil dibuat.\n\nSpreadsheet ID:\n{client.spreadsheet_id}\n\nURL:\n{url}")
                self.set_status("Google Sheet baru berhasil dibuat dan dihubungkan.")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Gagal membuat sheet", str(e))

        ttk.Button(frm, text="Buat Sekarang", command=do_create).grid(row=3, column=0, columnspan=2, sticky="ew")
        frm.columnconfigure(1, weight=1)
        win.grab_set()

    def open_settings(self):
        dialog = SettingsDialog(self.master, self.settings)
        self.master.wait_window(dialog)
        if dialog.result:
            self.settings = dialog.result
            save_settings(self.settings)
            self.update_mode_badge()
            self.set_status("Pengaturan sync disimpan.")

    def get_sync_client(self):
        if self.settings.get("mode") != "google_sheets":
            raise RuntimeError("Mode saat ini bukan google_sheets.")
        return GoogleSheetsSync(
            self.settings.get("google_credentials_file", ""),
            self.settings.get("google_spreadsheet_id", "")
        )

    def maybe_auto_sync(self):
        if self.settings.get("mode") == "google_sheets" and self.settings.get("auto_sync_on_save", True):
            try:
                client = self.get_sync_client()
                client.connect()
                client.upload_all(self.records)
            except Exception as e:
                self.set_status(f"Tersimpan lokal, sync cloud gagal: {e}")

    def sync_upload(self):
        try:
            client = self.get_sync_client()
            client.connect()
            client.upload_all(self.records)
            self.set_status("Upload cloud berhasil.")
            messagebox.showinfo("Sukses", "Upload ke Google Sheets berhasil.")
        except Exception as e:
            messagebox.showerror("Sync gagal", str(e))

    def sync_download(self):
        try:
            client = self.get_sync_client()
            client.connect()
            self.records = client.download_all()
            save_csv(DATA_FILE, FIELDNAMES, self.records)
            self.refresh_all()
            self.set_status("Download cloud berhasil.")
            messagebox.showinfo("Sukses", "Download dari Google Sheets berhasil.")
        except Exception as e:
            messagebox.showerror("Download gagal", str(e))

    def logout(self):
        self.destroy()
        self.master.show_login()


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.current_frame = None
        ensure_json(USERS_FILE, DEFAULT_USERS)
        ensure_json(PENDING_USERS_FILE, [])
        ensure_json(SETTINGS_FILE, DEFAULT_SETTINGS)
        settings = load_settings()
        self.geometry(f"{settings.get('window_width', 1600)}x{settings.get('window_height', 930)}")
        self.minsize(1100, 760)
        self.resizable(True, True)
        users = load_json(USERS_FILE, DEFAULT_USERS)
        users, changed = normalize_users_security(users)
        if changed:
            save_json(USERS_FILE, users)
        ensure_csv(DATA_FILE, FIELDNAMES)
        ensure_csv(ATTENDANCE_FILE, ATTENDANCE_FIELDS)
        self.show_login()

    def show_login(self):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = LoginWindow(self, self.show_app)

    def show_app(self, user):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = HRApp(self, user)


def main():
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
