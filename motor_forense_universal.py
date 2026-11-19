import os
from datetime import datetime
from difflib import SequenceMatcher
from typing import Tuple, Dict, Any, List

import pandas as pd
from docx import Document

PERFILES_ENTIDAD: Dict[str, Dict[str, Any]] = {
    "Sector Público": {
        "marcos": [
            "Ley 10-04 (Cámara de Cuentas)",
            "Ley 340-06 (Compras y Contrataciones)",
            "Ley 423-06 (Orgánica de Presupuesto)",
            "Ley 126-01 / NICSP (Contabilidad Gubernamental)",
            "ISSAI / INTOSAI (Auditoría Gubernamental)",
            "NOBACI (Control Interno CGR)",
            "ISO 37001 (Gestión Antisoborno)",
            "MESICIC",
        ],
        "criterios_control": [
            ("AMC-01", "Segregación de funciones entre solicitud, compra, recepción en almacén y pago (NOBACI 1.3)"),
            ("AMC-02", "Declaraciones juradas de conflicto de interés y adhesión al código de ética (Ley 41-08 / ISO 37001)"),
            ("ADC-01", "Certificación previa de disponibilidad presupuestaria antes de contraer compromisos (Ley 423-06)"),
            ("ADC-02", "Registro íntegro de ingresos propios/subvenciones y conciliaciones bancarias mensuales (Ley 126-01)"),
            ("ADC-03", "Cumplimiento de umbrales de contratación y ausencia de fraccionamiento (Ley 340-06 Art. 16)"),
            ("MYE-01", "Supervisión periódica de Auditoría Interna / UAI y arqueos físicos documentados (Ley 10-07)"),
        ],
        "umbral_concentracion": 0.30,
        "ventana_dias_fracc": 10,
    },
    "Sector Privado": {
        "marcos": [
            "ISO 37001 (Gestión Antisoborno)",
            "ISO 19011 (Auditoría de Sistemas de Gestión)",
            "Ley 155-17 (Lavado de Activos y Debida Diligencia)",
            "NIIF / NIC (Normas Internacionales de Información Financiera)",
            "Normativa Tributaria DGII (Formatos 606/607)",
            "COSO ERM (Control Interno Corporativo)",
        ],
        "criterios_control": [
            ("GOV-01", "Política antisoborno aprobada por el Órgano de Gobierno y Función de Cumplimiento independiente (ISO 37001)"),
            ("GOV-02", "Segregación de facultades en aprobaciones de compras, tesorería y alta de proveedores (COSO)"),
            ("DD-01", "Debida diligencia reforzada a socios de negocio, contratistas y beneficiarios finales (Ley 155-17)"),
            ("FIN-01", "Validación fiscal de comprobantes (NCF/e-CF) y recepción conforme antes del desembolso"),
            ("FIN-02", "Conciliaciones bancarias mensuales y control de desviaciones presupuestarias por centro de costo"),
            ("MON-01", "Canal ético operativo y auditoría interna basada en riesgos de fraude y soborno"),
        ],
        "umbral_concentracion": 0.35,
        "ventana_dias_fracc": 7,
    },
    "ONG / Entidad Mixta": {
        "marcos": [
            "Ley 122-05 (Regulación y Fomento de las ASFL)",
            "ISO 37001 (Gestión Antisoborno)",
            "ISSAI / INTOSAI (Fiscalización de Fondos Públicos)",
            "Normas de Rendición de Cuentas de Donantes",
            "Ley 155-17 (Debida Diligencia)",
        ],
        "criterios_control": [
            ("ONG-01", "Trazabilidad separada por proyecto/donante y prohibición de desvío de fondos etiquetados"),
            ("ONG-02", "Competencia transparente en adquisiciones con cargo a subvenciones o fondos públicos"),
            ("ONG-03", "Declaración de conflictos de interés de junta directiva, ejecutivos y suplidores"),
            ("ONG-04", "Conciliaciones bancarias mensuales y respaldo documental íntegro de egresos"),
        ],
        "umbral_concentracion": 0.30,
        "ventana_dias_fracc": 10,
    },
}

ALIAS_COMPRAS = {
    "proveedor": ["proveedor", "nombre_proveedor", "razon_social", "suplidor", "beneficiario", "contratista", "empresa"],
    "rnc": ["rnc", "rnc_proveedor", "identificacion", "cedula_rnc", "tax_id"],
    "monto": ["monto", "monto_pagado", "monto_devengado", "monto_comprometido", "valor", "importe", "total", "monto_total"],
    "presupuesto": ["presupuesto", "presupuesto_asignado", "presupuesto_asignado_partida", "asignado", "apropiacion"],
    "fecha": ["fecha", "fecha_documento", "fecha_pago", "fecha_factura", "fecha_orden", "fecha_emision"],
    "comprobante": ["ncf", "ncf_factura", "factura", "comprobante", "numero_factura", "documento"],
    "proceso": ["proceso", "numero_proceso", "orden_compra", "referencia", "codigo_proceso", "expediente"],
    "descripcion": ["descripcion", "descripcion_bien_servicio", "concepto", "detalle", "objeto", "rubro"],
    "autorizador": ["usuario_autorizador", "autorizador", "aprobador", "usuario", "funcionario"],
}

ALIAS_EMPLEADOS = {
    "cedula": ["cedula", "cedula_empleado", "identificacion", "dni", "documento", "id_nacional"],
    "nombre": ["nombre_completo", "nombre", "empleado", "funcionario", "nombres_apellidos"],
    "cargo": ["cargo", "cargo_puesto", "puesto", "posicion", "rol"],
    "departamento": ["departamento", "departamento_unidad", "unidad", "area", "direccion"],
    "telefono": ["telefono", "telefono_contacto", "celular", "movil"],
    "correo": ["correo", "correo_electronico", "email", "mail"],
    "direccion": ["direccion", "direccion_residencia", "domicilio"],
    "cuenta": ["cuenta_bancaria", "cuenta_bancaria_nomina", "cuenta", "iban"],
    "rnc_vinculado": ["rnc_empresa_vinculada", "rnc_vinculado", "empresa_vinculada"],
}

ALIAS_PROVEEDORES = {
    "rnc": ["rnc", "rnc_proveedor", "identificacion", "tax_id"],
    "razon_social": ["razon_social", "nombre_proveedor", "proveedor", "empresa", "suplidor"],
    "representante": ["representante_legal", "representante", "accionista", "gerente", "contacto"],
    "cedula_rep": ["cedula_representante", "cedula", "identificacion_representante", "dni_representante"],
    "telefono": ["telefono", "telefono_contacto", "celular"],
    "correo": ["correo", "correo_electronico", "email"],
    "direccion": ["direccion", "direccion_fiscal", "domicilio"],
    "cuenta": ["cuenta_bancaria", "cuenta", "iban"],
}


def _mapear_columnas(df: pd.DataFrame, mapa_alias: Dict[str, List[str]]) -> pd.DataFrame:
    d = df.copy()
    norm_cols = {str(c).strip().lower().replace(" ", "_"): c for c in d.columns}
    renombres = {}
    for canonico, variantes in mapa_alias.items():
        if canonico in norm_cols:
            renombres[norm_cols[canonico]] = canonico
            continue
        for v in variantes:
            if v in norm_cols:
                renombres[norm_cols[v]] = canonico
                break
    d = d.rename(columns=renombres)
    d.columns = [str(c).strip().lower().replace(" ", "_") for c in d.columns]
    return d


def calcular_nivel_riesgo(valor: float) -> str:
    r = float(valor)
    if r >= 20:
        return "Crítico"
    if r >= 15:
        return "Muy alto"
    if r >= 10:
        return "Alto"
    if r >= 5:
        return "Moderado"
    return "Bajo"


def analizar_compras_dinamico(
    df_raw: pd.DataFrame,
    umbral_concentracion: float = 0.30,
    ventana_dias: int = 7,
    min_ops_fracc: int = 3,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    d = _mapear_columnas(df_raw, ALIAS_COMPRAS)
    if "proveedor" not in d.columns:
        d["proveedor"] = d["rnc"].astype(str) if "rnc" in d.columns else "NO_ESPECIFICADO"
    if "monto" not in d.columns:
        d["monto"] = 0.0
    if "fecha" not in d.columns:
        d["fecha"] = pd.NaT

    d["proveedor"] = d["proveedor"].fillna("NO_ESPECIFICADO").astype(str).str.strip()
    d["monto"] = pd.to_numeric(d["monto"], errors="coerce").fillna(0.0)
    d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce")

    alertas: List[Dict[str, Any]] = []

    if "comprobante" in d.columns:
        valid_ncf = d[d["comprobante"].notna() & (d["comprobante"].astype(str).str.strip() != "")]
        dup_ncf = valid_ncf[valid_ncf.duplicated(subset=["proveedor", "comprobante"], keep=False)]
        for _, r in dup_ncf.iterrows():
            alertas.append(dict(
                tipo="Comprobante Duplicado", severidad="Crítico", proveedor=r["proveedor"],
                monto=float(r["monto"]), fecha=str(r["fecha"].date()) if pd.notna(r["fecha"]) else "",
                detalle=f"Comprobante repetido ({r['comprobante']}) para el mismo proveedor.",
                prob_sugerida=5, imp_sugerido=4
            ))

    valid_fecha = d[d["fecha"].notna()]
    dup_exact = valid_fecha[valid_fecha.duplicated(subset=["proveedor", "monto", "fecha"], keep=False)]
    for _, r in dup_exact.iterrows():
        alertas.append(dict(
            tipo="Pago/Orden Duplicada", severidad="Crítico", proveedor=r["proveedor"],
            monto=float(r["monto"]), fecha=str(r["fecha"].date()),
            detalle="Coincidencia exacta de proveedor, monto y fecha.",
            prob_sugerida=5, imp_sugerido=4
        ))

    total_universo = float(d["monto"].sum())
    if total_universo > 0:
        por_prov = d.groupby("proveedor")["monto"].sum().sort_values(ascending=False)
        for prov, suma in por_prov.items():
            pct = suma / total_universo
            if pct >= umbral_concentracion and len(por_prov) > 1:
                alertas.append(dict(
                    tipo="Concentración de Proveedor", severidad="Muy alto" if pct >= 0.50 else "Alto",
                    proveedor=prov, monto=float(suma), fecha="",
                    detalle=f"Concentra {pct:.1%} del monto total ({suma:,.2f} de {total_universo:,.2f}).",
                    prob_sugerida=4, imp_sugerido=5 if pct >= 0.50 else 4
                ))

    d_sorted = d.dropna(subset=["fecha"]).sort_values("fecha")
    for prov, grupo in d_sorted.groupby("proveedor"):
        g = grupo.sort_values("fecha").reset_index(drop=True)
        for i in range(len(g)):
            f_ini = g.loc[i, "fecha"]
            f_fin = f_ini + pd.Timedelta(days=ventana_dias)
            ventana = g[(g["fecha"] >= f_ini) & (g["fecha"] <= f_fin)]
            if len(ventana) >= min_ops_fracc:
                monto_acum = float(ventana["monto"].sum())
                f_max = ventana["fecha"].max().date()
                alertas.append(dict(tipo="Posible Fraccionamiento", severidad="Critico", proveedor=prov, monto=monto_acum, fecha=f"{f_ini.date()} a {f_max}", detalle=f"{len(ventana)} operaciones en {ventana_dias} dias por {monto_acum:,.2f}", prob_sugerida=5, imp_sugerido=5))
                break
    resumen=dict(total_registros=len(d),monto_total=total_universo,proveedores_unicos=int(d.proveedor.nunique()),total_alertas=len(alertas))
    return pd.DataFrame(alertas),d,resumen



def cruzar_debida_diligencia(de,dp):
    import pandas as pd
    return pd.DataFrame()

def generar_acta_custodia_docx(entidad,tipo_entidad,periodo,evidencias,usuario,ruta_salida):
    from docx import Document
    d=Document(); d.add_heading(chr(65)+chr(67)+chr(84)+chr(65),0); d.add_paragraph(str(entidad)); [d.add_paragraph(str(e)) for e in evidencias]; d.save(ruta_salida); return ruta_salida

def generar_informe_pericial_docx(entidad,tipo_entidad,periodo,casos,evidencias,ruta_salida):
    from docx import Document
    d=Document(); d.add_heading(chr(73)+chr(78)+chr(70)+chr(79)+chr(82)+chr(77)+chr(69),0); d.add_paragraph(str(entidad)); [d.add_paragraph(str(c)) for c in casos]; d.save(ruta_salida); return ruta_salida
