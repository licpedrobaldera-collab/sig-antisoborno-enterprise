import io
import re
import hashlib
from datetime import datetime, timedelta
from typing import List, Dict, Any

import pandas as pd
import requests
from docx import Document

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    from PIL import Image
    import pytesseract
except ImportError:
    Image = None
    pytesseract = None


REQUISITOS_ISO_37001_2025: List[Dict[str, Any]] = [
    {
        "codigo": "REQ-01",
        "pilar": "1. Gobernanza y Tolerancia Cero",
        "requisito": "Política Antisoborno aprobada por Alta Dirección con tolerancia cero (POL-SGAS-001).",
        "evidencia_requerida": "Política firmada, acta del Consejo y registro de difusión.",
        "norma": "ISO 37001:2025 Cl. 5.2 / ISSAI 200",
        "peso": 12,
        "palabras_clave": ["tolerancia cero", "politica antisoborno", "alta direccion", "prohibir", "soborno"],
    },
    {
        "codigo": "REQ-02",
        "pilar": "2. Alcance y Partes Interesadas",
        "requisito": "Cobertura a personal interno, socios de negocio, contratistas y todas las sedes.",
        "evidencia_requerida": "Matriz de partes interesadas y cláusulas contractuales anticorrupción.",
        "norma": "ISO 37001:2025 Cl. 4.2 - 4.3 / UK Bribery Act",
        "peso": 10,
        "palabras_clave": ["socios de negocio", "contratistas", "proveedores", "alcance", "terceros"],
    },
    {
        "codigo": "REQ-03",
        "pilar": "3. Prohibición Soborno y Facilitación",
        "requisito": "Regulación contra soborno activo/pasivo y pagos de facilitación a funcionarios.",
        "evidencia_requerida": "Código de ética, manual de compras y bitácora de incidentes.",
        "norma": "ISO 37001:2025 Cl. 8.7 / FCPA / Ley 448-06",
        "peso": 12,
        "palabras_clave": ["pagos de facilitacion", "soborno activo", "pasivo", "funcionarios publicos"],
    },
    {
        "codigo": "REQ-04",
        "pilar": "4. Regalos y Donaciones Políticas",
        "requisito": "Umbrales simbólicos para cortesías y prohibición de contribuciones políticas.",
        "evidencia_requerida": "Registro de regalos/atenciones y libro mayor de donaciones.",
        "norma": "ISO 37001:2025 Cl. 8.7 / COSO Control Interno",
        "peso": 8,
        "palabras_clave": ["regalos", "hospitalidad", "donaciones", "contribuciones politicas", "simbolico"],
    },
    {
        "codigo": "REQ-05",
        "pilar": "5. Gestión de Conflictos de Interés",
        "requisito": "Declaración anual de vínculos personales, familiares o financieros.",
        "evidencia_requerida": "Formularios firmados y cruces Nómina vs. Proveedores.",
        "norma": "ISO 37001:2025 Cl. 7.2 / Ley 340-06 Art. 14 / NOBACI",
        "peso": 12,
        "palabras_clave": ["conflictos de interes", "vinculo", "familiar", "imparcialidad", "declarar"],
    },
    {
        "codigo": "REQ-06",
        "pilar": "6. Debida Diligencia y Contabilidad",
        "requisito": "Evaluación previa de terceros y prohibición de registros fuera de balance.",
        "evidencia_requerida": "Expedientes KYC/RPE, estados financieros y conciliaciones.",
        "norma": "ISO 37001:2025 Cl. 8.2 - 8.3 / Ley 155-17 / NICSP-NIIF",
        "peso": 12,
        "palabras_clave": ["debida diligencia", "transparencia financiera", "libros contables", "cuentas ocultas"],
    },
    {
        "codigo": "REQ-07",
        "pilar": "7. Oficial de Cumplimiento",
        "requisito": "Designación formal de Oficial de Cumplimiento con acceso directo al Consejo.",
        "evidencia_requerida": "Resolución de nombramiento, perfil de puesto y organigrama.",
        "norma": "ISO 37001:2025 Cl. 5.3.2 / Modelo Singapur",
        "peso": 10,
        "palabras_clave": ["oficial de cumplimiento", "independencia", "acceso directo", "organo de gobierno"],
    },
    {
        "codigo": "REQ-08",
        "pilar": "8. Canal de Denuncias y No Represalias",
        "requisito": "Canales confidenciales operativos y prohibición estricta de represalias.",
        "evidencia_requerida": "Evidencia de canal ético activo y protocolo de investigación.",
        "norma": "ISO 37001:2025 Cl. 8.9 - 8.10 / Modelo Nórdico",
        "peso": 10,
        "palabras_clave": ["canal de denuncias", "represalias", "confidencialidad", "anonimo", "buena fe"],
    },
    {
        "codigo": "REQ-09",
        "pilar": "9. Control Social, Transparencia Activa y Publicación Oficial",
        "requisito": "Publicación oficial y accesible en el portal web institucional de la Política Antisoborno, canal ético, ejecución presupuestaria, compras, nómina y estados financieros para control social.",
        "evidencia_requerida": "Enlace activo del Portal de Transparencia / Sección Ética verificable en línea y captura de publicación.",
        "norma": "ISO 37001:2025 Cl. 7.4 / Ley 200-04 / Estándar Gobierno Abierto",
        "peso": 14,
        "palabras_clave": ["transparencia", "portal", "publicacion", "compras", "presupuesto", "nomina", "estados financieros", "etica", "antisoborno"],
    },
]


def extraer_texto_multiformato(archivo_bytes: bytes, nombre_archivo: str) -> Dict[str, Any]:
    ext = nombre_archivo.lower().split(".")[-1] if "." in nombre_archivo else ""
    sha256 = hashlib.sha256(archivo_bytes).hexdigest()
    texto_extraido = ""
    df_tabular = pd.DataFrame()
    metodo = "Binario/Genérico"

    try:
        if ext == "pdf":
            metodo = "Extracción PDF"
            if pypdf is not None:
                reader = pypdf.PdfReader(io.BytesIO(archivo_bytes))
                paginas = [page.extract_text() or "" for page in reader.pages]
                texto_extraido = "\n".join(paginas)
            else:
                texto_extraido = archivo_bytes.decode("latin-1", errors="ignore")

        elif ext in ("docx", "doc"):
            metodo = "Extracción Word (.docx)"
            doc = Document(io.BytesIO(archivo_bytes))
            partes = [p.text for p in doc.paragraphs if p.text.strip()]
            for t in doc.tables:
                for row in t.rows:
                    partes.append(" | ".join(cell.text.strip() for cell in row.cells))
            texto_extraido = "\n".join(partes)

        elif ext in ("xlsx", "xls"):
            metodo = "Extracción Excel"
            df_tabular = pd.read_excel(io.BytesIO(archivo_bytes))
            texto_extraido = df_tabular.to_string(index=False)

        elif ext == "csv":
            metodo = "Extracción CSV"
            for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    df_tabular = pd.read_csv(io.BytesIO(archivo_bytes), encoding=enc, sep=None, engine="python")
                    break
                except Exception:
                    continue
            if not df_tabular.empty:
                texto_extraido = df_tabular.to_string(index=False)

        elif ext in ("png", "jpg", "jpeg", "tiff", "bmp"):
            metodo = "Imagen / OCR"
            if Image is not None and pytesseract is not None:
                img = Image.open(io.BytesIO(archivo_bytes))
                texto_extraido = pytesseract.image_to_string(img)
            else:
                texto_extraido = f"[Imagen cargada y sellada: {nombre_archivo}]"

        else:
            metodo = "Texto Plano"
            texto_extraido = archivo_bytes.decode("utf-8", errors="ignore")

    except Exception as ex:
        texto_extraido = f"[Error leyendo {nombre_archivo}: {ex}]"

    return {
        "nombre": nombre_archivo,
        "extension": ext,
        "metodo": metodo,
        "sha256": sha256,
        "texto": texto_extraido,
        "dataframe": df_tabular,
        "caracteres": len(texto_extraido),
    }


def analizar_portal_transparencia_url(url: str) -> Dict[str, Any]:
    url_limpia = url.strip()
    if not url_limpia.startswith(("http://", "https://")):
        url_limpia = "https://" + url_limpia

    try:
        headers = {"User-Agent": "SIG-Antisoborno-Audit-Bot/6.1"}
        resp = requests.get(url_limpia, headers=headers, timeout=15)
        html_raw = resp.text
        sha256_web = hashlib.sha256(html_raw.encode("utf-8", errors="ignore")).hexdigest()

        # Limpiar etiquetas HTML para extraer texto visible y enlaces
        texto_limpio = re.sub(r"<script.*?>.*?</script>", " ", html_raw, flags=re.S | re.I)
        texto_limpio = re.sub(r"<style.*?>.*?</style>", " ", texto_limpio, flags=re.S | re.I)
        texto_limpio = re.sub(r"<[^>]+>", " ", texto_limpio)
        texto_limpio = re.sub(r"\s+", " ", texto_limpio).strip()

        enlaces_encontrados = re.findall(r'href=["\'](.*?)["\']', html_raw, flags=re.I)
        docs_publicados = [e for e in enlaces_encontrados if any(ext in e.lower() for ext in (".pdf", ".xlsx", ".xls", ".csv", ".doc", "transparencia", "etica", "compras", "presupuesto", "nomina", "finanzas"))]

        return {
            "ok": resp.status_code == 200,
            "status_code": resp.status_code,
            "url": url_limpia,
            "sha256": sha256_web,
            "texto": texto_limpio,
            "total_enlaces_relevantes": len(docs_publicados),
            "muestra_enlaces": docs_publicados[:15],
            "error": "",
        }
    except Exception as ex:
        return {
            "ok": False,
            "status_code": 0,
            "url": url_limpia,
            "sha256": "",
            "texto": "",
            "total_enlaces_relevantes": 0,
            "muestra_enlaces": [],
            "error": str(ex),
        }


def evaluar_brechas_documentales_ia(textos_cargados: List[str], url_verificada: bool = False) -> pd.DataFrame:
    corpus = " ".join(textos_cargados).lower()
    reemplazos = {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ñ": "n"}
    for k, v in reemplazos.items():
        corpus = corpus.replace(k, v)

    filas = []
    for req in REQUISITOS_ISO_37001_2025:
        hits = sum(1 for kw in req["palabras_clave"] if kw in corpus)
        total_kw = len(req["palabras_clave"])

def generar_politica_y_plan_90_dias_docx(entidad,tipo_entidad,pais,portal_url,canal_correo,canal_tel,canal_web,df_brechas,ruta_salida):
    from docx import Document
    from datetime import datetime, timedelta
    doc=Document()
    doc.add_heading(chr(80)+chr(79)+chr(76)+chr(73)+chr(84)+chr(73)+chr(67)+chr(65)+chr(32)+chr(73)+chr(83)+chr(79)+chr(32)+chr(51)+chr(55)+chr(48)+chr(48)+chr(49),0)
    doc.add_paragraph(str(entidad)+chr(32)+str(tipo_entidad)+chr(32)+str(pais))
    doc.add_paragraph(str(portal_url)+chr(32)+str(canal_correo))
    doc.add_heading(chr(49)+chr(46)+chr(32)+chr(84)+chr(111)+chr(108)+chr(101)+chr(114)+chr(97)+chr(110)+chr(99)+chr(105)+chr(97)+chr(32)+chr(67)+chr(101)+chr(114)+chr(111),level=1)
    doc.add_paragraph(chr(67)+chr(111)+chr(109)+chr(112)+chr(114)+chr(111)+chr(109)+chr(105)+chr(115)+chr(111)+chr(32)+chr(100)+chr(101)+chr(32)+chr(112)+chr(114)+chr(101)+chr(118)+chr(101)+chr(110)+chr(99)+chr(105)+chr(111)+chr(110)+chr(32)+chr(121)+chr(32)+chr(116)+chr(114)+chr(97)+chr(110)+chr(115)+chr(112)+chr(97)+chr(114)+chr(101)+chr(110)+chr(99)+chr(105)+chr(97)+chr(46))
    doc.add_heading(chr(50)+chr(46)+chr(32)+chr(80)+chr(108)+chr(97)+chr(110)+chr(32)+chr(57)+chr(48)+chr(32)+chr(68)+chr(105)+chr(97)+chr(115),level=1)
    doc.add_paragraph(chr(77)+chr(101)+chr(115)+chr(32)+chr(49)+chr(44)+chr(32)+chr(50)+chr(32)+chr(121)+chr(32)+chr(51))
    doc.save(ruta_salida)
    return ruta_salida
