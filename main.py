from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from schemas import ReportRequest
from reports import process_and_generate_reports 

app = FastAPI(
    title="CleanReport NDA Engine",
    description="Backend API for CleanReport SaaS"
)

# 👇 CORS CONFIGURATION: Frontend ko Backend se baat karne ki permission deta hai
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def home():
    return {"status": "Success", "message": "CleanReport API is Online! 🚀"}

@app.post("/generate-reports")
def generate_reports(data: ReportRequest):
    # Calling our Brain Engine!
    result = process_and_generate_reports(data)
    return result