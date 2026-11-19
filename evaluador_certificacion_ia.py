import os
import json
import hashlib
from datetime import datetime
from typing import Dict, Any, List, Tuple

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas


def calcular_calificacion_imparcial(
    puntaje_controles: float,
    casos: list,
    evidencias: list,
    alertas_forenses: int = 0,
    alertas_dd: int = 0,
) -> Dict[str, Any]:
    # 1. Dimensión Controles Internos / Políticas (35%)
    dim_controles = max(0.0, min(100.0, float(puntaje_controles)))

    # 2. Dimensión Integridad Forense y Debida Diligencia (25%)
    penalizacion = (alertas_forenses * 8.0) + (alertas_dd * 12.0)
    criticos_abiertos = sum(1 for c in casos if getattr(c, "estado", "") != "Cerrado" and getattr(c, "nivel_residual", "") in ("Crítico", "Muy alto"))
    penalizacion += criticos_abiertos * 10.0
    dim_forense = max(0.0, min(100.0, 100.0 - penalizacion))

    # 3. Dimensión Gestión de Riesgos y Cierre PAC (25%)
    if not casos:
        dim_pac = 100.0 if dim_controles >= 70 else 50.0
    else:
        prom_pac = sum(float(getattr(c, "avance_pac", 0.0)) for c in casos) / len(casos)
        cerrados = sum(1 for c in casos if getattr(c, "estado", "") == "Cerrado")
        tasa_cierre = (cerrados / len(casos)) * 100.0
        dim_pac = round((prom_pac * 0.6) + (tasa_cierre * 0.4), 2)

    # 4. Dimensión Trazabilidad y Evidencia Criptográfica SHA-256 (15%)
    if len(evidencias) == 0:
        dim_evidencia = 0.0
    elif not casos:
        dim_evidencia = min(100.0, len(evidencias) * 25.0)
    else:
        ratio_ev = min(1.0, len(evidencias) / max(1, len(casos)))
        dim_evidencia = round(ratio_ev * 100.0, 2)

    nota_global = round(
        (dim_controles * 0.35) +
        (dim_forense * 0.25) +
        (dim_pac * 0.25) +
        (dim_evidencia * 0.15),
        2,
    )

    if nota_global >= 90.0:
        nivel = "NIVEL AAA — ÓPTIMO (CERTIFICADO EN MADUREZ PLENA)"
        dictamen = "APROBADO PARA CERTIFICACIÓN INTEGRAL"
    elif nota_global >= 75.0:
        nivel = "NIVEL AA — SATISFACTORIO (IMPLEMENTACIÓN AVANZADA)"
        dictamen = "APROBADO CON OBSERVACIONES MENORES"
    elif nota_global >= 60.0:
        nivel = "NIVEL B — EN DESARROLLO (IMPLEMENTACIÓN PARCIAL)"
        dictamen = "CERTIFICACIÓN CONDICIONADA A CIERRE DE PAC"
    else:
        nivel = "NIVEL C — NO CONFORME (RIESGO CRÍTICO / INSUFICIENTE)"
        dictamen = "NO CALIFICA PARA CERTIFICACIÓN"

    payload_sello = {
        "nota_global": nota_global,
        "nivel": nivel,
        "dim_controles": dim_controles,
        "dim_forense": dim_forense,
        "dim_pac": dim_pac,
        "dim_evidencia": dim_evidencia,
        "timestamp": datetime.utcnow().strftime("%Y-%m-%d %H:%M"),
    }
    hash_certificado = hashlib.sha256(json.dumps(payload_sello, sort_keys=True).encode()).hexdigest()

    return {
        "nota_global": nota_global,
        "nivel": nivel,
        "dictamen": dictamen,
        "dim_controles": round(dim_controles, 1),
        "dim_forense": round(dim_forense, 1),
        "dim_pac": round(dim_pac, 1),
        "dim_evidencia": round(dim_evidencia, 1),
        "hash_certificado": hash_certificado,
        "fecha_utc": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
    }


def generar_guia_investigacion_ia(
    modalidad: str,
    tipo_entidad: str,
    entidad: str,
    evaluacion: Dict[str, Any],
    tiene_forense: bool,
    tiene_dd: bool,
) -> List[Dict[str, str]]:
    tareas = []
    if not entidad.strip():
        tareas.append({
            "prioridad": "ALTA",
            "modulo": "Contexto Lateral",
            "accion": "Definir el nombre oficial de la entidad, proyecto o empresa y el período a evaluar.",
            "dato_requerido": "Razón social / Institución y rango de fechas.",
        })

    if evaluacion["dim_controles"] == 0:
        tareas.append({
            "prioridad": "ALTA",
            "modulo": "Control Interno y Políticas",
            "accion": f"Evaluar los controles operativos y políticas vigentes bajo estándares de {tipo_entidad}.",
            "dato_requerido": "Manuales de funciones, políticas antisoborno, segregación de funciones y conciliaciones.",
        })

    if not tiene_forense:
        tareas.append({
            "prioridad": "MEDIA",
            "modulo": "Pruebas Forenses",
            "accion": "Cargar auxiliares reales de compras, pagos, facturación o ejecución presupuestaria.",
            "dato_requerido": "Archivo CSV/Excel con proveedor, monto, fecha, NCF/comprobante y presupuesto.",
        })

    if not tiene_dd:
        tareas.append({
            "prioridad": "MEDIA",
            "modulo": "Debida Diligencia",
            "accion": "Cruzar nómina de empleados/directivos contra el registro de proveedores y accionistas.",
            "dato_requerido": "Listados de personal (cédula, teléfono, cuenta) y suplidores (RNC, representante).",
        })

    if evaluacion["dim_evidencia"] < 100:
        tareas.append({
            "prioridad": "ALTA",
            "modulo": "Evidencias SHA-256",
            "accion": "Cargar los soportes documentales (PDF, actas, estados financieros, políticas) para sellado criptográfico.",
            "dato_requerido": "Documentos originales verificables vinculados a cada caso o control.",
        })

    if evaluacion["dim_pac"] < 80:
        tareas.append({
            "prioridad": "MEDIA",
            "modulo": "Casos y Plan de Acción (PAC)",
            "accion": "Documentar las 5 dimensiones del hallazgo (Condición, Criterio, Causa, Efecto, Recomendación) y actualizar avance PAC.",
            "dato_requerido": "Responsables asignados, fechas compromiso y porcentaje de cierre real.",
        })

    return tareas


def emitir_certificado_oficial_docx(
    entidad: str,
    tipo_entidad: str,
    modalidad: str,
    periodo: str,
    evaluacion: Dict[str, Any],
    usuario: str,
    ruta_salida: str,
) -> str:
    doc = Document()
    titulo = doc.add_heading("CERTIFICACIÓN TÉCNICA DE MADUREZ, CONTROL INTERNO Y CUMPLIMIENTO", 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("SIG Antisoborno Audit Enterprise · Evaluación Imparcial Basada en Evidencia")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")
    doc.add_paragraph(f"Entidad / Proyecto Evaluado: {entidad}")
    doc.add_paragraph(f"Naturaleza Institucional: {tipo_entidad}")
    doc.add_paragraph(f"Modalidad de Evaluación: {modalidad}")
    doc.add_paragraph(f"Período Auditado / Evaluado: {periodo}")
    doc.add_paragraph(f"Fecha y Hora de Certificación: {evaluacion['fecha_utc']}")
    doc.add_paragraph(f"Evaluador / Oficial Responsable: {usuario}")

    doc.add_heading("1. Resultado de Calificación Global y Dictamen", level=1)
    p_nota = doc.add_paragraph()
    run_nota = p_nota.add_run(f"CALIFICACIÓN GLOBAL OBTENIDA: {evaluacion['nota_global']}% / 100.0%\n")
    run_nota.bold = True
    run_nota.font.size = Pt(13)
    p_nota.add_run(f"Clasificación de Madurez: {evaluacion['nivel']}\n")
    p_nota.add_run(f"Dictamen Técnico del Sistema: {evaluacion['dictamen']}")

    doc.add_heading("2. Desglose Ponderado por Dimensión Auditada", level=1)
    tabla = doc.add_table(rows=1, cols=4)
    tabla.style = "Table Grid"
    hdr = tabla.rows.cells
    hdr.text = "Dimensión Evaluada"
    hdr.text = "Ponderación"
    hdr.text = "Calificación Obtenida"
    hdr.text = "Marco de Referencia"

    filas = [
        ("Diseño y Operación de Controles Internos y Políticas", "35%", f"{evaluacion['dim_controles']}%", "COSO ERM / NOBACI / ISO 37001"),
        ("Integridad Transaccional, Forense y Debida Diligencia", "25%", f"{evaluacion['dim_forense']}%", "ISO 19011 / Ley 340-06 / Ley 155-17"),
        ("Gestión de Riesgos y Cierre de Planes de Acción (PAC)", "25%", f"{evaluacion['dim_pac']}%", "ISSAI / INTOSAI / ISO 31000"),
        ("Trazabilidad Documental y Custodia de Evidencias", "15%", f"{evaluacion['dim_evidencia']}%", "Integridad Criptográfica SHA-256"),
    ]
    for dim, peso, nota, marco in filas:
        row = tabla.add_row().cells
        row.text = dim
        row.text = peso
        row.text = nota
        row.text = marco

    doc.add_heading("3. Garantía de Imparcialidad y Sello Criptográfico", level=1)
    doc.add_paragraph(
        "La presente valoración ha sido calculada de manera automatizada, objetiva e imparcial a partir "
        "exclusivamente de los datos transaccionales, controles verificados y evidencias digitales cargadas "
        "en la plataforma, sin intervención manual en la ponderación matemática."
    )
    doc.add_paragraph(f"Código Único de Verificación (SHA-256): {evaluacion['hash_certificado']}")

    doc.save(ruta_salida)
    return ruta_salida


def emitir_certificado_oficial_pdf(
    entidad: str,
    tipo_entidad: str,
    modalidad: str,
    periodo: str,
    evaluacion: Dict[str, Any],
    usuario: str,
    ruta_salida: str,
) -> str:
    c = canvas.Canvas(ruta_salida, pagesize=LETTER)
    w, h = LETTER
    y = h - 60

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w / 2, y, "CERTIFICACIÓN DE CONTROL INTERNO, CUMPLIMIENTO Y MADUREZ")
    y -= 20
    c.setFont("Helvetica", 10)
    c.drawCentredString(w / 2, y, "SIG Antisoborno Audit Enterprise · Evaluación Imparcial 100% Trazable")
    y -= 40

    c.setFont("Helvetica-Bold", 10)
    datos = [
        f"Entidad / Proyecto: {entidad}",
        f"Sector / Naturaleza: {tipo_entidad} | Modalidad: {modalidad}",
        f"Período Evaluado: {periodo}",
        f"Fecha de Emisión: {evaluacion['fecha_utc']} | Responsable: {usuario}",
    ]
    for d in datos:
        c.drawString(50, y, d)
        y -= 18

    y -= 15
    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, y, f"CALIFICACIÓN GLOBAL: {evaluacion['nota_global']}% / 100%")
    y -= 20
    c.setFont("Helvetica-Bold", 11)
