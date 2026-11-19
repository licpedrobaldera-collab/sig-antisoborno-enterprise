# SIG Antisoborno Audit Enterprise 5.0.0 Integral

Aplicación Streamlit integral para gestión de riesgos antisoborno, casos, histórico, evidencias SHA-256, pruebas forenses, debida diligencia, canal ético, base de conocimiento, dashboard configurable y exportación.

## Instalación Windows
```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Acceso inicial
- admin / Admin123!
- oficial / Oficial123!
- auditor / Auditor123!

Cambiar contraseñas antes de producción mediante:
- SIG_ADMIN_PASSWORD
- SIG_OFICIAL_PASSWORD
- SIG_AUDITOR_PASSWORD

## Riesgo
Probabilidad 1-5 × Impacto 1-5. El riesgo residual se calcula ajustando por efectividad del control.

## Producción
La versión incluida usa SQLite para despliegue simple/local. Para un entorno institucional se recomienda migrar la capa de persistencia a PostgreSQL y gestionar secretos mediante el proveedor de hosting.
