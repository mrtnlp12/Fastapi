from pydantic import BaseModel
from typing import List


class Cliente(BaseModel):
    nombre: str
    fila: int


class Metadata(BaseModel):
    fuenteTop: int
    concentradoTop: int
    insertados: int
    clientes: List[Cliente] = []
