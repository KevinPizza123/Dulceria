from contextlib import asynccontextmanager
from typing import List
from datetime import date, timedelta
from fastapi import FastAPI, Depends, HTTPException, Query
from sqlmodel import Session, select
from fastapi.responses import FileResponse
from database import create_db_and_tables, get_session
from models import Dulce, DulceCreate, DulceUpdate

@asynccontextmanager
async def lifespan(app: FastAPI):
    create_db_and_tables()
    yield

app = FastAPI(
    title="API de Inventario - Dulcería",
    description="Sistema de gestión de stock, precios y caducidades",
    version="1.0.0",
    lifespan=lifespan
)

@app.get("/", include_in_schema=False)
def servir_interfaz():
    return FileResponse("static/index.html")

@app.post("/dulces/", response_model=Dulce, status_code=201)
def crear_dulce(dulce: DulceCreate, session: Session = Depends(get_session)):
    db_dulce = Dulce.model_validate(dulce)
    session.add(db_dulce)
    session.commit()
    session.refresh(db_dulce)
    return db_dulce

@app.get("/dulces/", response_model=List[Dulce])
def listar_dulces(
    categoria: str | None = None,
    session: Session = Depends(get_session)
):
    query = select(Dulce)
    if categoria:
        query = query.where(Dulce.categoria == categoria)
    return session.exec(query).all()

@app.get("/dulces/{dulce_id}", response_model=Dulce)
def obtener_dulce(dulce_id: int, session: Session = Depends(get_session)):
    dulce = session.get(Dulce, dulce_id)
    if not dulce:
        raise HTTPException(status_code=404, detail="Dulce no encontrado")
    return dulce

@app.patch("/dulces/{dulce_id}", response_model=Dulce)
def actualizar_dulce(
    dulce_id: int, 
    dulce_data: DulceUpdate, 
    session: Session = Depends(get_session)
):
    db_dulce = session.get(Dulce, dulce_id)
    if not db_dulce:
        raise HTTPException(status_code=404, detail="Dulce no encontrado")
    
    datos_dict = dulce_data.model_dump(exclude_unset=True)
    for key, value in datos_dict.items():
        setattr(db_dulce, key, value)
        
    session.add(db_dulce)
    session.commit()
    session.refresh(db_dulce)
    return db_dulce

@app.post("/dulces/{dulce_id}/movimiento", response_model=Dulce)
def registrar_movimiento_stock(
    dulce_id: int,
    cantidad: int = Query(..., description="Positivo para entrada (compra), negativo para salida (venta)"),
    session: Session = Depends(get_session)
):
    dulce = session.get(Dulce, dulce_id)
    if not dulce:
        raise HTTPException(status_code=404, detail="Dulce no encontrado")
    
    if dulce.stock + cantidad < 0:
        raise HTTPException(status_code=400, detail="Stock insuficiente para realizar la salida")
    
    dulce.stock += cantidad
    session.add(dulce)
    session.commit()
    session.refresh(dulce)
    return dulce

@app.get("/alertas/")
def obtener_alertas(dias_caducidad: int = 30, session: Session = Depends(get_session)):
    todos = session.exec(select(Dulce)).all()
    fecha_limite = date.today() + timedelta(days=dias_caducidad)
    
    stock_bajo = [d for d in todos if d.stock <= d.stock_minimo]
    por_caducar = [
        d for d in todos 
        if d.fecha_caducidad is not None and d.fecha_caducidad <= fecha_limite
    ]
    
    return {
        "stock_bajo": stock_bajo,
        "proximos_a_caducar": por_caducar
    }

@app.delete("/dulces/{dulce_id}")
def eliminar_dulce(dulce_id: int, session: Session = Depends(get_session)):
    dulce = session.get(Dulce, dulce_id)
    if not dulce:
        raise HTTPException(status_code=404, detail="Dulce no encontrado")
    session.delete(dulce)
    session.commit()
    return {"mensaje": f"'{dulce.nombre}' eliminado del inventario"}