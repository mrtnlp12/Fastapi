from fastapi import APIRouter, File, HTTPException, UploadFile
from config.db import conn
from threading import Thread
from services import email
from services.excel import (
    generate_daily_report,
    generate_report_by_client,
    update_report,
)


router = APIRouter()


@router.post("/excel")
async def upload_excel(file: UploadFile = File(...)):
    if not file:
        raise HTTPException(status_code=400, detail="No se ha enviado un archivo")

    if not file.filename.endswith(".xlsx"):
        raise HTTPException(
            status_code=400, detail="El archivo no es un archivo de Excel"
        )

    with open("fuente.xlsx", "wb") as f:
        f.write(await file.read())
        f.close()

    metadata = conn.local.metadata.find_one()

    thread1 = Thread(
        target=generate_daily_report,
        args=(
            metadata,
            conn,
        ),
    )
    thread2 = Thread(
        target=generate_report_by_client,
        args=(
            metadata,
            conn,
        ),
    )
    thread3 = Thread(
        target=update_report,
        args=(
            metadata,
            conn,
        ),
    )

    thread1.start()
    thread2.start()
    thread3.start()

    thread1.join()
    thread2.join()
    thread3.join()

    # email.EmailService().sendConfirmationMessage()

    return {"filename": file.filename}
