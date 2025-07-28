from fastapi import FastAPI, Request, Form
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from app.excel_generator import generate_task_tracker
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="app/templates")


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/generate")
async def generate(request: Request, year: int = Form(...)):
    # 1) Prepare an in-memory buffer
    buffer = BytesIO()

    # 2) Generate the workbook into the buffer
    generate_task_tracker(year, buffer)

    # 3) Stream it back to the client with a download filename
    filename = f"task_tracker_{year}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )