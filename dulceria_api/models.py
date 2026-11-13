from typing import Optional
from datetime import date
from sqlmodel import Field, SQLModel

class DulceBase(SQLModel):
    nombre: str = Field(index=True)
    marca: Optional[str] = None
    categoria: str  # Ej: Chocolates, Gomitas, Caramelos, Chicles
    precio_compra: float
    precio_venta: float
    stock: int = Field(default=0)
    stock_minimo: int = Field(default=10)
    fecha_caducidad: Optional[date] = None

class Dulce(DulceBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)

class DulceCreate(DulceBase):
    pass

class DulceUpdate(SQLModel):
    nombre: Optional[str] = None
    marca: Optional[str] = None
    categoria: Optional[str] = None
    precio_compra: Optional[float] = None
    precio_venta: Optional[float] = None
    stock: Optional[int] = None
    stock_minimo: Optional[int] = None
    fecha_caducidad: Optional[date] = None