from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def generar_contrato_saas_docx(
    proveedor_firma: str,
    rnc_proveedor: str,
    cliente_nombre: str,
    rnc_cliente: str,
    plan_contratado: str,
    monto_usd: float,
    frecuencia_pago: str,
    usuarios_incluidos: int,
    ruta_salida: str,
) -> str:
    doc = Document()
    titulo = doc.add_heading("CONTRATO DE LICENCIA DE USO DE SOFTWARE (SaaS) Y TÉRMINOS DE SERVICIO", 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("SIG Antisoborno Audit Enterprise · Licencia Comercial de Uso")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")
    doc.add_paragraph(
        f"Entre {proveedor_firma}, sociedad comercial debidamente constituida conforme a las leyes de la "
        f"República Dominicana, titular del RNC Núm. {rnc_proveedor}, en lo adelante denominada EL PROVEEDOR; "
        f"y de la otra parte {cliente_nombre}, entidad/sociedad identificada con el RNC Núm. {rnc_cliente}, "
        f"en lo adelante denominada EL CLIENTE, se conviene y pacta el presente Contrato:"
    )

    clausulas = [
        ("PRIMERA: OBJETO DEL CONTRATO Y ALCANCE DE LA LICENCIA",
         f"EL PROVEEDOR otorga a EL CLIENTE una licencia de uso no exclusiva, intransferible y revocable para acceder y operar la plataforma tecnológica SIG Antisoborno Audit Enterprise bajo la modalidad {plan_contratado}, habilitando hasta {usuarios_incluidos} usuarios autorizados."),
        ("SEGUNDA: PROPIEDAD INTELECTUAL Y RESERVA DE DERECHOS",
         f"El código fuente, algoritmos forenses, arquitectura de datos, diseños, manuales y marcas asociadas al sistema son propiedad intelectual exclusiva de {proveedor_firma}, protegidos ante la ONDA y ONAPI. Queda estrictamente prohibida la ingeniería inversa, descompilación, sublicenciamiento o reproducción no autorizada."),
        ("TERCERA: INVERSIÓN, FORMA DE PAGO Y VIGENCIA",
         f"Por concepto de licenciamiento y soporte técnico, EL CLIENTE pagará la suma de US${monto_usd:,.2f} bajo esquema de facturación {frecuencia_pago}. La falta de pago oportuno facultará a EL PROVEEDOR a suspender temporalmente el acceso hasta la regularización del saldo."),
        ("CUARTA: CONFIDENCIALIDAD, SEGURIDAD Y CADENA DE CUSTODIA SHA-256",
         "Toda información financiera, contable, documental o de denuncias cargada por EL CLIENTE es estrictamente confidencial y pertenece exclusivamente a EL CLIENTE. El sistema garantiza la integridad de las evidencias mediante sellado criptográfico SHA-256 y bitácoras inmutables."),
        ("QUINTA: ALCANCE DE LAS EVALUACIONES Y CERTIFICACIONES TÉCNICAS",
         "Las puntuaciones de madurez, diagnósticos ISO 37001:2025, mapas de calor y certificados emitidos por el software reflejan de manera objetiva e imparcial los datos y evidencias suministrados por EL CLIENTE, sin sustituir certificaciones acreditadas por organismos externos de normalización."),
        ("SEXTA: SOPORTE TÉCNICO Y DISPONIBILIDAD (SLA)",
         "EL PROVEEDOR garantiza una disponibilidad operativa de la plataforma en la nube del 99.5% mensual y brindará soporte técnico especializado para actualizaciones normativas y resolución de incidencias."),
        ("SÉPTIMA: JURISDICCIÓN Y LEY APLICABLE",
         "Para todo lo no previsto en este contrato, las partes se someten a las leyes de la República Dominicana y a la competencia de los tribunales del Distrito Nacional."),
    ]

    for tit, texto in clausulas:
        doc.add_heading(tit, level=2)
        doc.add_paragraph(texto)

    doc.add_paragraph("")
    doc.add_paragraph(f"Hecho y firmado en dos originales de un mismo tenor y efecto, en fecha {datetime.utcnow().strftime('%Y-%m-%d')}.")
    doc.add_paragraph("")
    doc.add_paragraph(f"_______________________________                    _______________________________")
    doc.add_paragraph(f"Por EL PROVEEDOR ({proveedor_firma})               Por EL CLIENTE ({cliente_nombre})")

    doc.save(ruta_salida)
    return ruta_salida


def generar_brochure_propuesta_docx(
    firma_consultora: str,
    cliente_prospecto: str,
    sector_prospecto: str,
    precio_propuesto_usd: float,
    modalidad_oferta: str,
    contacto_comercial: str,
    ruta_salida: str,
) -> str:
    doc = Document()
    titulo = doc.add_heading("PROPUESTA EJECUTIVA Y BROCHURE DE SOLUCIÓN TECNOLÓGICA", 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("SIG Antisoborno Audit Enterprise 6.0 · Gobernanza, Auditoría Forense y Cumplimiento ISO 37001:2025")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")
    doc.add_paragraph(f"Preparado especialmente para: {cliente_prospecto} ({sector_prospecto})")
    doc.add_paragraph(f"Presentado por: {firma_consultora} | Contacto: {contacto_comercial}")
    doc.add_paragraph(f"Fecha de Presentación: {datetime.utcnow().strftime('%Y-%m-%d')}")

    doc.add_heading("1. ¿Por qué implementar SIG Antisoborno Audit Enterprise?", level=1)
    doc.add_paragraph(
        "Las organizaciones públicas y privadas enfrentan crecientes exigencias regulatorias en materia de transparencia, "
        "prevención de soborno, debida diligencia de terceros y control interno. Nuestra plataforma transforma la auditoría "
        "tradicional en un sistema continuo, automatizado y auditable al 100%."
    )

    doc.add_heading("2. Capacidades Clave de la Plataforma", level=1)
    capacidades = [
        ("Diagnóstico Automatizado ISO 37001:2025:", "Ingesta ilimitada de documentos (Word, PDF, Excel, Imágenes) con detección inteligente de brechas frente a estándares internacionales (ISO 37001, ISSAI, COSO ERM, UK Bribery Act, FCPA)."),
        ("Motor Forense Transaccional:", "Identificación automática de fraccionamiento de compras, facturas duplicadas, concentración de suplidores y desviaciones presupuestarias."),
        ("Debida Diligencia y Conflictos de Interés:", "Cruce multifactorial entre nóminas/directivos y proveedores para prevenir incompatibilidades y direccionamiento contractual."),
        ("Generador IA de Políticas y Plan a 90 Días:", "Creación inmediata de la Política Integral POL-SGAS-001 y hoja de ruta trimestral para entidades que inician desde cero."),
        ("Cadena de Custodia SHA-256 y Certificación:", "Sellado criptográfico de todas las evidencias, mapa de calor 5x5, comparativa histórica entre ciclos y emisión de Certificados Técnicos de Madurez."),
    ]
    for negrita, desc in capacidades:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(negrita + " ")
        r.bold = True
        p.add_run(desc)

    doc.add_heading("3. Propuesta Económica y Alcance", level=1)
    doc.add_paragraph(f"• Modalidad Recomendada: {modalidad_oferta}")
    doc.add_paragraph(f"• Inversión Estimada: US${precio_propuesto_usd:,.2f}")
    doc.add_paragraph("• Entregables Incluidos: Acceso a plataforma, configuración de perfiles RBAC, capacitación operativa, soporte técnico y emisión de certificaciones periódicas.")

    doc.save(ruta_salida)
    return ruta_salida
