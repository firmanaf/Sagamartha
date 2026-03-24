from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from datetime import date, datetime
import json
import os

app = FastAPI(title="Sagamartha HR Enterprise v4 Backend")

# Konfigurasi CORS agar frontend dari domain lain (Vercel/Netlify) bisa memanggil API backend ini
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # Ganti "*" dengan spesifik URL Vercel/Netlify Anda di produksi
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory storage mimicking DB for prototype
USERS_FILE = "enterprise_users_pwkv4.json"
RECORDS_FILE = "enterprise_records.json"  # Converting CSV to JSON approach for simple API
ATTENDANCE_FILE = "enterprise_attendance.json"
OVERTIME_FILE = "enterprise_overtime.json"

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
    status: str = "pending"

class AttendanceRecord(BaseModel):
    id: str
    employee_name: str
    attendance_date: str
    check_in_time: str
    check_out_time: Optional[str] = ""
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

def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

@app.get("/")
def read_root():
    return {"message": "Sagamartha HR v4 API is running."}

@app.get("/api/users", response_model=List[UserBase])
def get_users():
    users = load_json(USERS_FILE, [])
    # Filter sensitive data
    return [{"username": u["username"], "employee_name": u["employee_name"], "role": u["role"], "supervisor_name": u.get("supervisor_name", "")} for u in users]

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
    with open(ATTENDANCE_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)
    return {"message": "Attendance recorded successfully."}

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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
