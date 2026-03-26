from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from datetime import date, datetime
import json
import os
import hashlib

from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

app = FastAPI(title="Sagamartha HR Enterprise v4 Backend")
app.mount("/static", StaticFiles(directory="."), name="static")

@app.get("/")
def read_root():
    return FileResponse("frontend_sagamartha_v4.html")

# Konfigurasi CORS agar frontend dari domain lain (Vercel/Netlify) bisa memanggil API backend ini
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # Ganti "*" dengan spesifik URL Vercel/Netlify Anda di produksi
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory storage mimicking DB for prototype
DATA_DIR = os.getenv("DATA_DIR", ".")
if DATA_DIR != "." and not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR, exist_ok=True)

USERS_FILE = os.path.join(DATA_DIR, "enterprise_users_pwkv4.json")
RECORDS_FILE = os.path.join(DATA_DIR, "enterprise_records.json")
ATTENDANCE_FILE = os.path.join(DATA_DIR, "enterprise_attendance.json")
OVERTIME_FILE = os.path.join(DATA_DIR, "enterprise_overtime.json")
WORKLOAD_FILE = os.path.join(DATA_DIR, "project_workload_pwkv4.json")
AGENDAS_FILE = os.path.join(DATA_DIR, "enterprise_agendas.json")
CHECKINGS_FILE = os.path.join(DATA_DIR, "enterprise_checkings.json")

class UserBase(BaseModel):
    username: str
    employee_name: str
    role: str
    supervisor_name: Optional[str] = ""

class PerformanceRecord(BaseModel):
    id: str
    employee_name: str
    work_date: str
    project_title: str
    target_description: str
    target_category: str
    daily_target: float
    daily_performance: float
    achievement_percent: float
    supervisor_name: Optional[str] = ""
    status: str = "pending"
    kendala: Optional[str] = ""

class AgendaRecord(BaseModel):
    id: str
    project_title: str
    description: str
    start_date: str
    end_date: str
    supervisor_name: str
    involved_employees: List[str] # Usernames

class CheckingRecord(BaseModel):
    id: str
    project_title: str
    check_date: str
    supervisor_name: str
    description: str

class AttendanceRecord(BaseModel):
    id: str
    employee_name: str
    attendance_date: str
    check_in_time: str
    check_out_time: Optional[str] = ""
    check_in_lat: Optional[float] = None
    check_in_lng: Optional[float] = None
    check_out_lat: Optional[float] = None
    check_out_lng: Optional[float] = None
    status: str = "hadir"
    notes: Optional[str] = ""

class OvertimeRecord(BaseModel):
    id: str
    employee_name: str
    overtime_date: str
    start_time: str
    end_time: str
    duration_hours: float
    reason: str
    manager_approval: str = "pending"
    comments: Optional[str] = ""

class WorkloadData(BaseModel):
    projects: List[str]
    matrix: dict # {username: {project_nama: value}}

class LoginRequest(BaseModel):
    username: str
    password: str

class PasswordChangeRequest(BaseModel):
    old_password: str
    new_password: str

def load_json(path, default):
    if not os.path.exists(path):
        # Migrasi: Jika di /data belum ada, tapi di root project ada, copy ke /data
        basename = os.path.basename(path)
        if os.path.exists(basename):
            print(f"DEBUG: Migrating {basename} from root to {path}")
            try:
                with open(basename, "r", encoding="utf-8") as fs:
                    data = json.load(fs)
                with open(path, "w", encoding="utf-8") as fd:
                    json.dump(data, fd, indent=2)
                return data
            except Exception as e:
                print(f"DEBUG: Migration error for {basename}: {e}")
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

@app.post("/api/login")
def login(creds: LoginRequest):
    users = load_json(USERS_FILE, [])
    # hash password dengan sha256
    pwd_hash = hashlib.sha256(creds.password.encode("utf-8")).hexdigest()
    
    for u in users:
        if u["username"] == creds.username and u["password"] == pwd_hash:
            return {
                "message": "Login successful",
                "user": {
                    "username": u["username"],
                    "employee_name": u["employee_name"],
                    "role": u["role"],
                    "supervisor_name": u.get("supervisor_name", "")
                }
            }
            
    raise HTTPException(status_code=401, detail="Masukan username atau password salah")

@app.get("/", response_class=HTMLResponse)
def read_root():
    # Serve the frontend file for local verification
    try:
        with open("frontend_sagamartha_v4.html", "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        return f"<h1>Error loading frontend: {e}</h1>"

@app.get("/api/users", response_model=List[UserBase])
def get_users():
    users = load_json(USERS_FILE, [])
    # Filter sensitive data
    return [{"username": u["username"], "employee_name": u["employee_name"], "role": u["role"], "supervisor_name": u.get("supervisor_name", "")} for u in users]

@app.delete("/api/users/{username}")
def delete_user(username: str):
    users = load_json(USERS_FILE, [])
    initial_count = len(users)
    users = [u for u in users if u["username"] != username]
    if len(users) == initial_count:
        raise HTTPException(status_code=404, detail="User not found.")
    
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2)
    return {"message": f"User {username} deleted successfully."}

@app.put("/api/users/{username}/password")
def change_password(username: str, req: PasswordChangeRequest, admin_override: bool = False):
    users = load_json(USERS_FILE, [])
    for u in users:
        if u["username"] == username:
            old_hash = hashlib.sha256(req.old_password.encode("utf-8")).hexdigest()
            # If admin_override is true, bypass old_password check (In real app, verify admin token)
            if not admin_override and u["password"] != old_hash:
                raise HTTPException(status_code=400, detail="Password lama salah.")
            
            new_hash = hashlib.sha256(req.new_password.encode("utf-8")).hexdigest()
            u["password"] = new_hash
            
            with open(USERS_FILE, "w", encoding="utf-8") as f:
                json.dump(users, f, indent=2)
            return {"message": "Password berhasil diubah."}
            
    raise HTTPException(status_code=404, detail="User not found.")

@app.get("/api/performance", response_model=List[PerformanceRecord])
def get_performance(user: Optional[str] = None, dt: Optional[str] = None):
    records = load_json(RECORDS_FILE, [])
    if user:
        records = [r for r in records if r["employee_name"] == user]
    if dt:
        records = [r for r in records if r["work_date"] == dt]
    return records

@app.post("/api/performance", status_code=status.HTTP_201_CREATED)
def create_performance(record: PerformanceRecord):
    records = load_json(RECORDS_FILE, [])
    
    # Anti-duplication check v4
    for r in records:
        if (r["employee_name"] == record.employee_name and 
            r["work_date"] == record.work_date and 
            r["project_title"] == record.project_title and 
            r["target_description"] == record.target_description):
            raise HTTPException(status_code=400, detail="Data kinerja identik sudah ada (Duplicate).")

    records.append(record.dict())
    with open(RECORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": "Record created successfully."}

@app.put("/api/performance/{record_id}")
def update_performance(record_id: str, record: PerformanceRecord):
    records = load_json(RECORDS_FILE, [])
    for i, r in enumerate(records):
        if r["id"] == record_id:
            # Update fields
            record_dict = record.dict()
            # Recalculate achievement if needed
            if record_dict["daily_target"] > 0:
                record_dict["achievement_percent"] = min(100, round((record_dict["daily_performance"] / record_dict["daily_target"]) * 100))
            
            records[i] = record_dict
            with open(RECORDS_FILE, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=2)
            return {"message": f"Record {record_id} updated successfully."}
    raise HTTPException(status_code=404, detail="Record not found.")

@app.delete("/api/performance/{record_id}")
def delete_performance(record_id: str):
    records = load_json(RECORDS_FILE, [])
    initial_count = len(records)
    records = [r for r in records if r["id"] != record_id]
    if len(records) == initial_count:
        raise HTTPException(status_code=404, detail="Record not found.")
    
    with open(RECORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": f"Record {record_id} deleted successfully."}

@app.put("/api/performance/{record_id}/approval")
def approve_performance(record_id: str, approval_status: str):
    if approval_status not in ["approved", "rejected", "pending"]:
        raise HTTPException(status_code=400, detail="Invalid status")
    
    records = load_json(RECORDS_FILE, [])
    for r in records:
        if r["id"] == record_id:
            r["status"] = approval_status
            with open(RECORDS_FILE, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=2)
            return {"message": f"Record {record_id} marked as {approval_status}."}
            
    raise HTTPException(status_code=404, detail="Record not found.")

@app.get("/api/attendance", response_model=List[AttendanceRecord])
def get_attendance(user: Optional[str] = None, dt: Optional[str] = None):
    records = load_json(ATTENDANCE_FILE, [])
    if user:
        records = [r for r in records if r["employee_name"] == user]
    if dt:
        records = [r for r in records if r["attendance_date"] == dt]
    return records

@app.post("/api/attendance", status_code=status.HTTP_201_CREATED)
def create_attendance(record: AttendanceRecord):
    records = load_json(ATTENDANCE_FILE, [])
    # Validasi satu absensi per hari per user
    for r in records:
        if r["employee_name"] == record.employee_name and r["attendance_date"] == record.attendance_date:
            raise HTTPException(status_code=400, detail="Data absensi hari ini sudah ada.")
    records.append(record.dict())
    print(f"DEBUG: Saving attendance to {ATTENDANCE_FILE}. Total records: {len(records)}")
    with open(ATTENDANCE_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": "Attendance recorded successfully."}

class CheckoutRequest(BaseModel):
    check_out_time: str
    check_out_lat: Optional[float] = None
    check_out_lng: Optional[float] = None

@app.put("/api/attendance/{record_id}/checkout")
def checkout_attendance(record_id: str, req: CheckoutRequest):
    records = load_json(ATTENDANCE_FILE, [])
    for r in records:
        if r["id"] == record_id:
            if r.get("check_out_time"):
                raise HTTPException(status_code=400, detail="Sudah check-out sebelumnya.")
            r["check_out_time"] = req.check_out_time
            r["check_out_lat"] = req.check_out_lat
            r["check_out_lng"] = req.check_out_lng
            with open(ATTENDANCE_FILE, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=2)
            return {"message": "Check-out berhasil."}
    raise HTTPException(status_code=404, detail="Record absensi tidak ditemukan.")

@app.delete("/api/attendance/{record_id}")
def delete_attendance(record_id: str):
    records = load_json(ATTENDANCE_FILE, [])
    initial_count = len(records)
    records = [r for r in records if r["id"] != record_id]
    if len(records) == initial_count:
        raise HTTPException(status_code=404, detail="Record absensi tidak ditemukan.")
    
    with open(ATTENDANCE_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": f"Record absensi {record_id} berhasil dihapus/reset."}

@app.get("/api/overtime", response_model=List[OvertimeRecord])
def get_overtime(user: Optional[str] = None):
    records = load_json(OVERTIME_FILE, [])
    if user:
        records = [r for r in records if r["employee_name"] == user]
    return records

@app.post("/api/overtime", status_code=status.HTTP_201_CREATED)
def create_overtime(record: OvertimeRecord):
    records = load_json(OVERTIME_FILE, [])
    records.append(record.dict())
    print(f"DEBUG: Saving overtime to {OVERTIME_FILE}. Total records: {len(records)}")
    with open(OVERTIME_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": "Overtime request created successfully."}

@app.put("/api/overtime/{record_id}/approval")
def approve_overtime(record_id: str, manager_approval: str, comments: Optional[str] = ""):
    if manager_approval not in ["approved", "rejected", "pending"]:
        raise HTTPException(status_code=400, detail="Invalid approval status")
    records = load_json(OVERTIME_FILE, [])
    for r in records:
        if r["id"] == record_id:
            r["manager_approval"] = manager_approval
            if comments: r["comments"] = comments
            with open(OVERTIME_FILE, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=2)
            return {"message": f"Overtime {record_id} marked as {manager_approval}."}
    raise HTTPException(status_code=404, detail="Overtime not found.")

@app.get("/api/workload")
def get_workload(month: str = None):
    # Returns workload for a specific month (YYYY-MM). 
    all_data = load_json(WORKLOAD_FILE, {})
    
    # MIGRATION: If old format detected (projects/matrix keys at root), move to current month
    if "projects" in all_data and "matrix" in all_data:
        from datetime import datetime
        curr_month = datetime.now().strftime("%Y-%m")
        old_data = {"projects": all_data["projects"], "matrix": all_data["matrix"]}
        all_data = { curr_month: old_data }
        with open(WORKLOAD_FILE, "w", encoding="utf-8") as f:
            json.dump(all_data, f, indent=2)
            
    if not month:
        from datetime import datetime
        month = datetime.now().strftime("%Y-%m")
        
    if month in all_data:
        return all_data[month]
    
    # Copy-on-access logic for new months
    sorted_months = sorted([m for m in all_data.keys() if "-" in m], reverse=True)
    for m in sorted_months:
        if m < month:
            # Found a previous month, return it as a template
            return all_data[m]
            
    return {"projects": [], "matrix": {}}

@app.post("/api/workload")
def save_workload(data: WorkloadData, month: str = None):
    # Saves workload for a specific month
    if not month:
        from datetime import datetime
        month = datetime.now().strftime("%Y-%m")
        
    all_data = load_json(WORKLOAD_FILE, {})
    all_data[month] = data.dict()
    with open(WORKLOAD_FILE, "w", encoding="utf-8") as f:
        json.dump(all_data, f, indent=2)
    return {"message": f"Workload data for {month} saved successfully."}

@app.get("/api/agendas", response_model=List[AgendaRecord])
def get_agendas():
    return load_json(AGENDAS_FILE, [])

@app.post("/api/agendas", status_code=status.HTTP_201_CREATED)
def create_agenda(agenda: AgendaRecord):
    agendas = load_json(AGENDAS_FILE, [])
    agendas.append(agenda.dict())
    with open(AGENDAS_FILE, "w", encoding="utf-8") as f:
        json.dump(agendas, f, indent=2)
    return {"message": "Agenda created successfully."}

@app.put("/api/agendas/{agenda_id}")
def update_agenda(agenda_id: str, agenda: AgendaRecord):
    agendas = load_json(AGENDAS_FILE, [])
    for i, a in enumerate(agendas):
        if a["id"] == agenda_id:
            agendas[i] = agenda.dict()
            with open(AGENDAS_FILE, "w", encoding="utf-8") as f:
                json.dump(agendas, f, indent=2)
            return {"message": "Agenda updated successfully."}
    raise HTTPException(status_code=404, detail="Agenda not found.")

@app.delete("/api/agendas/{agenda_id}")
def delete_agenda(agenda_id: str):
    agendas = load_json(AGENDAS_FILE, [])
    initial_count = len(agendas)
    agendas = [a for a in agendas if a["id"] != agenda_id]
    if len(agendas) == initial_count:
        raise HTTPException(status_code=404, detail="Agenda not found.")
    with open(AGENDAS_FILE, "w", encoding="utf-8") as f:
        json.dump(agendas, f, indent=2)
    return {"message": "Agenda deleted."}

# CHECKINGS API
@app.get("/api/checkings", response_model=List[CheckingRecord])
def get_checkings():
    return load_json(CHECKINGS_FILE, [])

@app.post("/api/checkings", status_code=status.HTTP_201_CREATED)
def create_checking(record: CheckingRecord):
    records = load_json(CHECKINGS_FILE, [])
    records.append(record.dict())
    with open(CHECKINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": "Checking recorded successfully."}

@app.put("/api/checkings/{checking_id}")
def update_checking(checking_id: str, record: CheckingRecord):
    records = load_json(CHECKINGS_FILE, [])
    for i, r in enumerate(records):
        if r["id"] == checking_id:
            records[i] = record.dict()
            with open(CHECKINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(records, f, indent=2)
            return {"message": "Checking record updated successfully."}
    raise HTTPException(status_code=404, detail="Checking record not found.")

@app.delete("/api/checkings/{checking_id}")
def delete_checking(checking_id: str):
    records = load_json(CHECKINGS_FILE, [])
    initial_count = len(records)
    records = [r for r in records if r["id"] != checking_id]
    if len(records) == initial_count:
        raise HTTPException(status_code=404, detail="Checking record not found.")
    with open(CHECKINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": "Checking record deleted."}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
