from fastapi import APIRouter, HTTPException
from config.db import conn
from schemas.metadata import metadataEntity
from models.metadata import Metadata

router = APIRouter()


@router.get("/metadata")
async def getMetadata():
    metadata = metadataEntity(conn.local.metadata.find_one())

    if metadata is None:
        raise HTTPException(status_code=404, detail="Metadata not found")

    return metadata


@router.post("/metadata")
async def initializeMetadata(metadata: Metadata):
    new_metadata = metadataEntity(dict(metadata))
    conn.local.metadata.insert_one(new_metadata)
    return "Metadata initialized"
