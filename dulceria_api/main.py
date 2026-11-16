import io
from pathlib import Path
from contextlib import asynccontextmanager
from typing import List
from datetime import date, datetime, timedelta

from fastapi import FastAPI, Depends, HTTPException, Query, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from sqlmodel import Session, select

from database import create_db_and_tables, get_session
from models import Dulce, DulceCreate, DulceUpdate

BASE_DIR = Path(__file__).resolve().parent


@asynccontextmanager
async def lifespan(app: FastAPI):
    create_db_and_tables()
    yield


app = FastAPI(
    title="API de Inventario - Dulcería",
    description="Sistema de gestión de stock, precios, caducidades, carga masiva y reportes Excel/PDF",
    version="1.3.0",
    lifespan=lifespan
)


@app.get("/", include_in_schema=False)
def servir_interfaz():
    ruta_html = BASE_DIR / "static" / "index.html"
    return FileResponse(
        ruta_html,
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
        },
    )


@app.post("/dulces/", response_model=Dulce, status_code=201)
def crear_dulce(dulce: DulceCreate, session: Session = Depends(get_session)):
    db_dulce = Dulce.model_validate(dulce)
    session.add(db_dulce)
    session.commit()
    session.refresh(db_dulce)
    return db_dulce


@app.post("/dulces/importar-excel/", status_code=201)
async def importar_dulces_excel(
    archivo: UploadFile = File(...),
    session: Session = Depends(get_session)
):
    if not archivo.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Por favor sube un archivo Excel válido (.xlsx)")

    contenido = await archivo.read()
    wb = load_workbook(filename=io.BytesIO(contenido), data_only=True)
    hoja = wb.active

    creados = 0
    errores = []

    for num_fila, fila in enumerate(hoja.iter_rows(min_row=2, values_only=True), start=2):
        if not any(fila):
            continue

        try:
            nombre, marca, categoria, precio_compra, precio_venta, stock, stock_minimo, fecha_cad = (fila + (None,) * 8)[:8]

            if not nombre or not categoria or precio_compra is None or precio_venta is None:
                errores.append(f"Fila {num_fila}: Faltan datos obligatorios (nombre, categoría o precios).")
                continue

            fecha_final = None
            if isinstance(fecha_cad, datetime):
                fecha_final = fecha_cad.date()
            elif isinstance(fecha_cad, date):
                fecha_final = fecha_cad
            elif isinstance(fecha_cad, str) and fecha_cad.strip():
                fecha_final = date.fromisoformat(fecha_cad.strip()[:10])

            nuevo_dulce = Dulce(
                nombre=str(nombre).strip(),
                marca=str(marca).strip() if marca else None,
                categoria=str(categoria).strip(),
                precio_compra=float(precio_compra),
                precio_venta=float(precio_venta),
                stock=int(stock) if stock is not None else 0,
                stock_minimo=int(stock_minimo) if stock_minimo is not None else 10,
                fecha_caducidad=fecha_final
            )
            session.add(nuevo_dulce)
            creados += 1

        except Exception as e:
            errores.append(f"Fila {num_fila}: Error de formato ({str(e)})")

    session.commit()

    return {
        "mensaje": f"Se importaron {creados} productos correctamente.",
        "importados": creados,
        "errores": errores
    }


@app.get("/dulces/exportar-excel/")
def exportar_dulces_excel(
    categoria: str | None = None,
    session: Session = Depends(get_session)
):
    query = select(Dulce)
    if categoria:
        query = query.where(Dulce.categoria == categoria)
    dulces = session.exec(query).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte de Inventario"

    encabezados = [
        "ID", "Nombre", "Marca", "Categoría",
        "Precio Compra", "Precio Venta", "Stock", "Stock Mínimo",
        "Fecha Caducidad", "Valor Costo Total", "Valor Venta Total", "Estado"
    ]
    ws.append(encabezados)

    header_fill = PatternFill(start_color="DB2777", end_color="DB2777", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    fecha_limite = date.today() + timedelta(days=30)

    for d in dulces:
        costo_total = round(d.precio_compra * d.stock, 2)
        venta_total = round(d.precio_venta * d.stock, 2)

        estados = []
        if d.stock <= d.stock_minimo:
            estados.append("Stock Bajo")
        if d.fecha_caducidad and d.fecha_caducidad <= fecha_limite:
            estados.append("Próximo a Caducar")
        estado_texto = " / ".join(estados) if estados else "OK"

        ws.append([
            d.id,
            d.nombre,
            d.marca or "Sin marca",
            d.categoria,
            d.precio_compra,
            d.precio_venta,
            d.stock,
            d.stock_minimo,
            d.fecha_caducidad.isoformat() if d.fecha_caducidad else "Sin fecha",
            costo_total,
            venta_total,
            estado_texto
        ])

    if dulces:
        fila_total = len(dulces) + 2
        ws.cell(row=fila_total, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=fila_total, column=7, value=f"=SUM(G2:G{fila_total - 1})").font = Font(bold=True)
        ws.cell(row=fila_total, column=10, value=f"=SUM(J2:J{fila_total - 1})").font = Font(bold=True)
        ws.cell(row=fila_total, column=11, value=f"=SUM(K2:K{fila_total - 1})").font = Font(bold=True)

    for col in ws.columns:
        max_len = max(len(str(celda.value or "")) for celda in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    nombre_archivo = f"reporte_inventario_{date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"}
    )


@app.get("/dulces/exportar-pdf/")
def exportar_dulces_pdf(
    categoria: str | None = None,
    session: Session = Depends(get_session)
):
    query = select(Dulce)
    if categoria:
        query = query.where(Dulce.categoria == categoria)
    dulces = session.exec(query).all()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24
    )
    styles = getSampleStyleSheet()
    elementos = []

    titulo = f"Reporte de Inventario - Dulcería ({categoria if categoria else 'Todas las categorías'})"
    elementos.append(Paragraph(titulo, styles["Title"]))
    elementos.append(Paragraph(f"Fecha de emisión: {date.today().isoformat()} | Total productos: {len(dulces)}", styles["Normal"]))
    elementos.append(Spacer(1, 14))

    datos = [[
        "ID", "Nombre", "Marca", "Categoría",
        "P. Compra", "P. Venta", "Stock", "Caducidad",
        "Costo Total", "Venta Total", "Estado"
    ]]

    fecha_limite = date.today() + timedelta(days=30)
    total_unidades = 0
    total_costo = 0.0
    total_venta = 0.0

    for d in dulces:
        costo_total = round(d.precio_compra * d.stock, 2)
        venta_total = round(d.precio_venta * d.stock, 2)
        total_unidades += d.stock
        total_costo += costo_total
        total_venta += venta_total

        estados = []
        if d.stock <= d.stock_minimo:
            estados.append("Stock Bajo")
        if d.fecha_caducidad and d.fecha_caducidad <= fecha_limite:
            estados.append("Caduca Pronto")
        estado_texto = " / ".join(estados) if estados else "OK"

        datos.append([
            str(d.id),
            d.nombre[:26],
            (d.marca or "—")[:16],
            d.categoria,
            f"${d.precio_compra:.2f}",
            f"${d.precio_venta:.2f}",
            str(d.stock),
            d.fecha_caducidad.isoformat() if d.fecha_caducidad else "—",
            f"${costo_total:.2f}",
            f"${venta_total:.2f}",
            estado_texto
        ])

    datos.append([
        "TOTAL", "", "", "", "", "",
        str(total_unidades), "",
        f"${total_costo:.2f}",
        f"${total_venta:.2f}",
        ""
    ])

    tabla = Table(datos, repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DB2777")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ALIGN", (1, 1), (2, -2), "LEFT"),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F8FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#FCE7F3")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    elementos.append(tabla)
    doc.build(elementos)
    buffer.seek(0)

    nombre_archivo = f"reporte_inventario_{date.today().isoformat()}.pdf"
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"}
    )


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