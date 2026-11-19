import json
import hashlib
from datetime import datetime
from typing import Dict, Any, Tuple

import streamlit as st
import pandas as pd


def calcular_progreso_expediente(
    entidad: str,
    periodo: str,
    tiene_control: bool,
    casos: list,
    evidencias: list,
) -> Tuple[int, str, str, list]:
    pasos = []

    p1 = bool(entidad and entidad.strip() and periodo and periodo.strip())
    pasos.append(("1. Perfilamiento de Entidad y Período", 15, p1, "Complete el nombre de la entidad y el período evaluado en el panel izquierdo."))

    p2 = bool(tiene_control)
    pasos.append(("2. Evaluación de Control Interno (NOBACI/COSO)", 15, p2, "Vaya al módulo Control Interno y marque el cumplimiento real para calcular la efectividad."))

    p3 = len(casos) > 0
    pasos.append(("3. Detección Forense / Registro de Casos", 20, p3, "Cargue archivos reales en Pruebas Forenses o registre al menos un hallazgo en Casos y Riesgos."))

    casos_completos = [
        c for c in casos
        if len(str(c.condicion or "").strip()) > 10
        and len(str(c.criterio or "").strip()) > 5
        and len(str(c.causa or "").strip()) > 5
        and len(str(c.efecto or "").strip()) > 5
        and len(str(c.recomendacion or "").strip()) > 5
    ]
    p4 = len(casos_completos) > 0 and len(casos_completos) == len(casos)
    pasos.append(("4. Estructura Técnica Completa (5 Atributos ISSAI)", 20, p4, "Asegúrese de que todos los casos tengan Condición, Criterio, Causa, Efecto y Recomendación."))

    p5 = len(evidencias) > 0
    pasos.append(("5. Sellado Criptográfico de Evidencias (SHA-256)", 15, p5, "Suba los soportes documentales originales en el módulo Evidencias para generar su huella SHA-256."))

    p6 = p1 and p2 and p3 and p4 and p5
    pasos.append(("6. Emisión de Informes Finales (Word/Excel/PDF)", 15, p6, "Diríjase al módulo Exportación para descargar el Informe Pericial y el Acta de Custodia."))

    porcentaje = sum(pts for _, pts, ok, _ in pasos if ok)
    siguiente_titulo = "Expediente Completo al 100%"
    siguiente_guia = "Todos los requisitos técnicos y documentales están listos. Puede exportar los informes finales."

    for titulo, _, ok, guia in pasos:
        if not ok:
            siguiente_titulo = titulo
            siguiente_guia = guia
            break

    return porcentaje, siguiente_titulo, siguiente_guia, pasos


def renderizar_guia_ia(
    entidad: str,
    periodo: str,
    tiene_control: bool,
    casos: list,
    evidencias: list,
):
    pct, sig_paso, sig_guia, lista_pasos = calcular_progreso_expediente(
        entidad, periodo, tiene_control, casos, evidencias
    )

    with st.container(border=True):
        c1, c2 = st.columns([1, 3])
        c1.metric("Progreso Global del Informe", f"{pct}% / 100%")
        c2.progress(pct / 100.0)
        if pct < 100:
            st.info(f"**Copiloto IA — Siguiente paso recomendado ({sig_paso}):** {sig_guia}")
        else:
            st.success(f"**Copiloto IA — {sig_paso}:** {sig_guia}")

        with st.expander("📋 Ver Marco Completo de Pasos hacia el 100%", expanded=False):
            for titulo, pts, ok, guia in lista_pasos:
                icono = "✅" if ok else "⏳"
                st.write(f"{icono} **{titulo}** (+{pts}%) — {guia}")
