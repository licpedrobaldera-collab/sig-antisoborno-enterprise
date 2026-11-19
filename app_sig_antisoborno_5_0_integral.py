import os
import streamlit as st
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, Text, Boolean
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime

VERSION = "6.0.0"
APP_TITLE = "SIG Antisoborno Audit Enterprise 6.0.0 SaaS & Big-4 Suite"
DB_PATH = os.getenv("SIG_DB_PATH", "sig_antisoborno_6_0_enterprise.db")
EXPORT_DIR = "exportaciones_6_0"
EVIDENCE_DIR = "evidencias_6_0"
os.makedirs(EXPORT_DIR, exist_ok=True)
os.makedirs(EVIDENCE_DIR, exist_ok=True)

st.set_page_config(page_title=APP_TITLE, page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")

engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()

class Caso(Base):
    __tablename__ = "casos"
    id = Column(Integer, primary_key=True); codigo = Column(String, unique=True, index=True); titulo = Column(String)
    departamento = Column(String); proceso = Column(String); responsable = Column(String); marco = Column(String)
    estado = Column(String, default="Abierto"); probabilidad = Column(Integer); impacto = Column(Integer)
    riesgo_inherente = Column(Float); nivel_inherente = Column(String); efectividad_control = Column(Float, default=0)
    riesgo_residual = Column(Float); nivel_residual = Column(String); avance_pac = Column(Float, default=0)
    condicion = Column(Text, default=""); criterio = Column(Text, default=""); causa = Column(Text, default=""); efecto = Column(Text, default="")
    recomendacion = Column(Text, default=""); monto_inversion = Column(Float, default=0.0); cuota_mensual = Column(Float, default=0.0)
    fecha_efectividad = Column(String, default=""); creado_por = Column(String, default="admin")
    fecha_creacion = Column(DateTime, default=datetime.utcnow); fecha_actualizacion = Column(DateTime, default=datetime.utcnow); eliminado = Column(Boolean, default=False)

class CicloEvaluacion(Base):
    __tablename__ = "ciclos_evaluacion"
    id = Column(Integer, primary_key=True); codigo_ciclo = Column(String, unique=True); entidad = Column(String)
    tipo_entidad = Column(String); frecuencia = Column(String); periodo = Column(String)
    nota_global = Column(Float); nota_controles = Column(Float); nota_forense = Column(Float)
    nota_pac = Column(Float); nota_evidencias = Column(Float); dictamen = Column(String)
    hash_sha256 = Column(String); usuario = Column(String); fecha = Column(DateTime, default=datetime.utcnow)

class Historico(Base):
    __tablename__ = "historico"
    id = Column(Integer, primary_key=True); codigo = Column(String, index=True); usuario = Column(String); accion = Column(String)
    estado_anterior = Column(String, default=""); estado_nuevo = Column(String, default=""); riesgo_anterior = Column(String, default="")
    riesgo_actual = Column(String, default=""); detalle = Column(Text, default=""); hash_sha256 = Column(String); timestamp = Column(DateTime, default=datetime.utcnow)

class Evidencia(Base):
    __tablename__ = "evidencias"
    id = Column(Integer, primary_key=True); codigo_caso = Column(String, index=True); nombre = Column(String); tipo = Column(String)
    sha256 = Column(String); ruta = Column(String); usuario = Column(String); descripcion = Column(Text, default=""); fecha = Column(DateTime, default=datetime.utcnow)

class Denuncia(Base):
    __tablename__ = "denuncias"
    id = Column(Integer, primary_key=True); ticket = Column(String, unique=True); pin = Column(String); categoria = Column(String)
    departamento = Column(String); descripcion = Column(Text); monto = Column(Float, default=0); anonima = Column(Boolean, default=True)
    estado = Column(String, default="Recibida"); evidencia_sha256 = Column(String, default=""); fecha = Column(DateTime, default=datetime.utcnow)

class Conocimiento(Base):
    __tablename__ = "conocimiento"
    id = Column(Integer, primary_key=True); codigo = Column(String, unique=True); pais = Column(String); caso = Column(String); periodo = Column(String)
    monto_usd = Column(Float, default=0); sector = Column(String); tipologia = Column(Text); patrones = Column(String); norma = Column(String)
    version = Column(Integer, default=1); activo = Column(Boolean, default=True); actualizado = Column(DateTime, default=datetime.utcnow)

Base.metadata.create_all(engine)

from interfaz_enterprise_6_0 import ejecutar_suite_enterprise
ejecutar_suite_enterprise(SessionLocal, Caso, CicloEvaluacion, Historico, Evidencia, Denuncia, Conocimiento, EXPORT_DIR, EVIDENCE_DIR, VERSION)

