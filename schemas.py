from pydantic import BaseModel
from typing import List, Dict

class EmployeeRecord(BaseModel):
    emp_no: str
    name: str
    post: str
    basic_pay: int
    total_duties: int
    attendance: Dict[str, bool]

# Naya Multi-Month Block
class MonthData(BaseModel):
    month: int
    year: int
    records: List[EmployeeRecord]

# Final Payload jo React bhejega
class ReportRequest(BaseModel):
    ministry: str
    department: str
    office: str
    night_hours_per_duty: float
    months_data: List[MonthData]