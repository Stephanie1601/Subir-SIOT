# app.py
import os
import io
import time
import json
import base64
from pathlib import Path

import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Carga SIOT → Pipefy", page_icon="📤", layout="wide")

# ---------- THEME / CSS ----------
st.markdown('''
<style>
/* Fuente general Arial para toda la app */
html, body, [class*="css"] {
    font-family: 'Arial', sans-serif !important;
}

/* Contenedor principal más abajo para que no tape el logo */
.block-container {
    max-width: 1200px;
    padding-top: 4.2rem !important;
    margin-top: 0 !important;
}

/* Botones */
div.stButton > button {
    border-radius: 12px;
    padding: 0.6rem 1rem;
    font-weight: 600;
    font-family: 'Arial', sans-serif !important;
}

/* Sidebar */
.stSidebar, .sidebar .sidebar-content {
    background: linear-gradient(180deg, #fafafa, #f0f0f0);
}

/* Tarjetas KPI */
.kpi {
    padding: 12px 16px;
    border-radius: 14px;
    background: #ffffff;
    box-shadow: 0 3px 12px rgba(0,0,0,.06);
    border: 1px solid #eee;
}

/* Secciones */
.section {
    padding: 16px;
    border-radius: 16px;
    background: #ffffff;
    border: 1px solid #ececec;
    box-shadow: 0 3px 16px rgba(0,0,0,.05);
}

/* Encabezados globales (fallback) */
h1, h2, h3 { 
    font-family: 'Arial', sans-serif !important;
    font-weight: 800;
}

/* ESTILOS ESPECÍFICOS DE LOGIN */
.title-io {               /* 🚇 Instrucción Operacional de Trabajos */
    font-size: 2.6rem;    /* MÁS GRANDE */
    line-height: 1.2;
    margin: 0.75rem 0 0.25rem 0;
    color: #FF9016;       /* Naranja EOMMT */
    font-weight: 800;
}
.subtitle-login {         /* 🔐 Ingreso al sistema */
    font-size: 1.6rem;    /* MÁS PEQUEÑO */
    line-height: 1.2;
    margin: 0.5rem 0 0.75rem 0;
    color: #444444;
    font-weight: 800;
}

/* Espaciador superior fino */
.header-spacer { height: 15px; }

/* Badge modo automático */
.badge {
    display:inline-block; padding:4px 8px; border-radius:999px;
    background:#eef6ff; color:#185adb; font-size:.85rem; font-weight:700;
    border:1px solid #d6e8ff;
}

small.help { color: #666; }
</style>
''', unsafe_allow_html=True)

# ---------- SIMPLE AUTH ----------
AUTH_USERS = json.loads(os.environ.get("AUTH_USERS_JSON", os.getenv("AUTH_USERS_JSON", '{"admin":"admin"}')))

# ---------- LOGO HELPERS ----------
def _find_logo_bytes() -> bytes | None:
    candidates = [
        Path("/mnt/data/06ccb9c2-ca99-49b6-a58e-9452a7e6a452.png"),
        Path("/mnt/data/Logo EOMMT.png"),
        Path(__file__).parent / "Logo EOMMT.png",
        Path("Logo EOMMT.png"),
        Path("logo_eommt.png"),
        Path("assets/Logo EOMMT.png"),
    ]
    for p in candidates:
        try:
            if p.exists():
                return p.read_bytes()
        except Exception:
            pass
    return None

def render_logo_center(width_px: int = 220):
    img_bytes = _find_logo_bytes()
    if not img_bytes:
        st.info("No se encontró el logo de EOMMT en el servidor.")
        return
    b64 = base64.b64encode(img_bytes).decode("ascii")
    st.markdown(
        f"""
        <div style="text-align:center; margin: 10px 0 8px 0;">
            <img src="data:image/png;base64,{b64}" width="{width_px}" />
        </div>
        """,
        unsafe_allow_html=True
    )

def render_logo_sidebar(width_px: int = 160):
    img_bytes = _find_logo_bytes()
    if not img_bytes:
        return
    b64 = base64.b64encode(img_bytes).decode("ascii")
    st.sidebar.markdown(
        f"""
        <div style="text-align:center; margin: 6px 0 10px 0;">
            <img src="data:image/png;base64,{b64}" width="{width_px}" />
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------- LOGIN ----------
def login_view():
    left, center, right = st.columns([1, 1, 1])
    with center:
        st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
        render_logo_center(width_px=200)
        st.markdown('<h2 class="title-io">🚇 Instrucción Operacional de Trabajos</h2>', unsafe_allow_html=True)
        st.markdown('<h3 class="subtitle-login">🔐 Ingreso al sistema</h3>', unsafe_allow_html=True)
        st.write("Por favor ingresa tus credenciales para continuar:")

        user = st.text_input("Usuario", key="login_user", placeholder="Escribe tu usuario")
        pwd  = st.text_input("Contraseña", key="login_pwd", type="password", placeholder="Escribe tu contraseña")
        ok = st.button("Ingresar", key="btn_login", use_container_width=True)

        if ok:
            if user in AUTH_USERS and AUTH_USERS.get(user) == pwd:
                st.session_state['auth_user'] = user
                st.success("✅ Acceso concedido.")
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")
    return 'auth_user' in st.session_state

def require_auth():
    if 'auth_user' in st.session_state:
        return True
    ok = login_view()
    if not ok:
        st.stop()
    return True

def logout_button():
    with st.sidebar:
        if st.button("🚪 Cerrar sesión", type="secondary", use_container_width=True):
            st.session_state.pop('auth_user', None)
            st.success("Sesión cerrada.")
            st.rerun()

# ---------- PIPEFY / EXCEL HELPERS ----------
# Mapeo AUTOMÁTICO: etiqueta de columna (encabezado Excel) -> field_id de Pipefy
LABEL_TO_FIELD_ID = {
    "IOT": "siot",
    "EMPRESA": "empresa",
    "CCU": "ccu",
    "INTEGRANTES DE CUADRILLA": "integrantes_de_cuadrilla",
    "CONTACTO CCU": "c_dula_ccu",
    "CANTÓN / ESTACIÓN": "cant_n_estaci_n",
    "ZONA DE ESTACIÓN": "zona_de_trabajo",
    "DESCRIPCIÓN DE ACTIVIDAD": "descripci_n_de_actividad",
    "FECHA DE INICIO": "fecha_de_inicio",
    "FECHA DE FIN": "fecha_de_fin",
    "HORA DE INICIO": "hora_de_inicio",
    "HORA DE FIN": "hora_de_fin",
    "TIPO DE JORNADA": "tipo_de_jornada",
    "TIPO DE MANTENIMIENTO / INSPECCIÓN": "tipo_de_mantenimiento",
    "N° REGISTRO DE FALLA": "registro_de_incidente",
    "CATEGORÍA DE RIESGO": "categor_a_de_riesgo",
    "CATEGORÍA DE TRABAJOS": "categor_a_de_trabajos",
    "DESENERGIZACIONES": "desenergizaci_n",
    "VEHÍCULO": "veh_culo",
    "ILUMINACIÓN PARCIAL": "iluminaci_n_parcia",
    "SEÑALETICA PROPIA": "se_aletica_propia",
    "R1": "r1_1",
    "R2": "r2_1",
    "P1": "p1",
    "P3": "p3",
    "E1": "e1",
    "V3": "v3",
    "P6": "copy_of_se_aletica_propia",
    "P7": "copy_of_r1",
    "P8": "copy_of_p3",
    "BLOQUEO DE VÍA": "bloqueo_de_v_a_1",
    "DESDE": "bloqueo_desde",
    "HASTA": "hasta",
    "Seleccionar etiqueta": "seleccionar_etiqueta",
    "CORREO ELECTRÓNICO DEL SOLICITANTE": "correo_electr_nico_del_solicitante",
}

def read_excel_header_row(uploaded_bytes: bytes, header_row_index_1based: int = 8) -> pd.DataFrame:
    """
    Lee SIEMPRE tomando la fila 'header_row_index_1based' como encabezados.
    - header_row_index_1based=8 -> header=7 en pandas.
    """
    bio = io.BytesIO(uploaded_bytes)
    # Tomamos SIEMPRE la primera hoja (es lo más estable para este caso)
    df = pd.read_excel(bio, engine="openpyxl", header=header_row_index_1based - 1)
    return df

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # eliminar columnas Unnamed
    df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]
    # limpiar encabezados
    df.columns = [str(c).strip() for c in df.columns]
    # eliminar filas completamente vacías
    df = df.dropna(how="all")
    # normalizar strings
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def build_fields_attributes(row: dict, mapping: dict) -> list:
    attrs = []
    for col, field_id in mapping.items():
        if not field_id:
            continue
        value = row.get(col)
        if value is None:
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        # Formateo simple para fechas (Pipefy acepta string ISO o dd/mm/yyyy según config)
        if hasattr(value, "strftime"):
            value = value.strftime("%Y-%m-%d")
        attrs.append({"field_id": field_id, "field_value": value})
    return attrs

def pipefy_create_card(pipe_id: int, fields_attrs: list, token: str):
    url = "https://api.pipefy.com/graphql"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    mutation = """
    mutation($input: CreateCardInput!) {
      createCard(input: $input) {
        card { id }
      }
    }
    """
    variables = {"input": {"pipe_id": pipe_id, "fields_attributes": fields_attrs}}
    try:
        resp = requests.post(url, headers=headers, json={"query": mutation, "variables": variables}, timeout=60)
    except Exception as e:
        return False, None, [{"message": str(e)}], str(e)

    ok = (resp.status_code == 200)
    data = {}
    try:
        data = resp.json()
    except Exception:
        pass
    errors = data.get("errors")
    card_id = None
    if data.get("data") and data["data"].get("createCard"):
        card_id = data["data"]["createCard"]["card"]["id"]
    return ok and (errors is None) and (card_id is not None), card_id, errors, resp.text

# ---------- SECRETS / VARS ----------
def get_secret(name, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return os.getenv(name, default)

PIPE_ID_ENV = get_secret("PIPEFY_PIPE_ID")
TOKEN_ENV   = get_secret("PIPEFY_TOKEN")
DRY_RUN_ENV = str(get_secret("PIPEFY_DRY_RUN", "false")).lower() == "true"
AUTO_MODE   = bool(PIPE_ID_ENV and TOKEN_ENV)

# ---------- APP ----------
if require_auth():
    # Branding sidebar + logout
    render_logo_sidebar(width_px=150)
    logout_button()

    # Panel lateral sólo si no hay secrets
    if not AUTO_MODE:
        with st.sidebar:
            st.subheader("🔧 Configuración Pipefy")
            pipe_id = st.text_input("Pipe ID", placeholder="Ej. 123456789")
            token = st.text_input("API Token", type="password", placeholder="Token secreto de Pipefy")
            dry_run = st.toggle("Simular (no crea tarjetas)", value=True, help="Haz pruebas antes de subir definitivamente.")
    else:
        with st.sidebar:
            st.markdown(f"<span class='badge'>Modo automático (secrets)</span>", unsafe_allow_html=True)
            st.write(f"Pipe ID: **{PIPE_ID_ENV}**")
            st.write("Token: **••••••••**")
            st.write(f"Dry run: **{DRY_RUN_ENV}**")
        pipe_id = str(PIPE_ID_ENV)
        token   = str(TOKEN_ENV)
        dry_run = DRY_RUN_ENV

    # Logo y títulos
    st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
    render_logo_center(width_px=220)

    st.title("📤 SIOT → Pipefy")
    st.caption("Sube tu archivo Excel. Tomamos la **fila 8** como encabezados y creamos tarjetas desde la **fila 9** hasta la última con **EMPRESA**.")

    # Uploader
    up = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

    if up is not None:
        content = up.read()
        # 1) Leer con encabezado en fila 8
        df = read_excel_header_row(content, header_row_index_1based=8)
        df = clean_dataframe(df)

        # 2) Quedarse sólo con filas desde la 9 y hasta última con EMPRESA
        if "EMPRESA" not in df.columns:
            st.error("No se encontró la columna 'EMPRESA' en el archivo. Verifica que el encabezado esté exactamente en la fila 8.")
            st.stop()

        # El df ya inicia en fila 8 como encabezado; df.iloc[1:] empieza fila 9
        df_from_9 = df.iloc[1:].copy()

        # recortar hasta última con EMPRESA
        mask_empresa = df_from_9["EMPRESA"].notna() & (df_from_9["EMPRESA"].astype(str).str.strip() != "")
        if not mask_empresa.any():
            st.error("No hay datos a partir de la fila 9 en la columna 'EMPRESA'.")
            st.stop()
        last_idx = df_from_9.index[mask_empresa].max()
        df_data = df_from_9.loc[df_from_9.index.min(): last_idx].copy()

        st.subheader("👀 Vista previa (desde fila 9 hasta última con EMPRESA)")
        st.dataframe(df_data.head(50), use_container_width=True)

        # 3) Mapeo automático: tomar sólo columnas presentes en el Excel
        auto_mapping = {col: LABEL_TO_FIELD_ID[col] for col in df_data.columns if col in LABEL_TO_FIELD_ID}

        # Info de mapeo
        st.markdown("**Columnas mapeadas automáticamente:** " + ", ".join(auto_mapping.keys()) if auto_mapping else "No se pudo mapear ninguna columna automáticamente.")
        missing_cols = [c for c in df_data.columns if c not in auto_mapping and not str(c).startswith("Unnamed")]
        if missing_cols:
            st.caption("Columnas sin mapeo (no se enviarán a Pipefy): " + ", ".join(missing_cols))

        # 4) KPIs
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi'><b>Columnas mapeadas</b><br>{len(auto_mapping)}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi'><b>Filas totales (desde fila 9)</b><br>{len(df_from_9)}</div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><b>Filas a subir</b><br>{len(df_data)}</div>", unsafe_allow_html=True)

        st.markdown("---")

        if not pipe_id or not token:
            st.error("Faltan credenciales de Pipefy. Define `PIPEFY_PIPE_ID` y `PIPEFY_TOKEN` en *secrets* o complétalos en el panel lateral.")
            st.stop()

        # 5) Iniciar proceso AUTOMÁTICO
        st.info(f"Iniciando proceso {'(simulación)' if dry_run else ''} con Pipe ID {pipe_id}…")

        try:
            pipe_id_int = int(pipe_id)
        except:
            st.error("Pipe ID debe ser numérico.")
            st.stop()

        progress = st.progress(0.0, text="Iniciando…")
        logs = []
        ok_count, fail_count, skipped = 0, 0, 0
        total = len(df_data)

        for i, (_, row) in enumerate(df_data.iterrows(), start=1):
            row_dict = row.to_dict()
            fields = build_fields_attributes(row_dict, auto_mapping)

            if not fields:
                skipped += 1
                # i + 8 -> índice real en Excel (por el encabezado en 8 y arranque en 9)
                logs.append({"estado": "omitida", "razon": "Sin campos mapeados con datos", "fila_excel": i + 8})
                progress.progress(i/total, text=f"Omitida fila Excel {i+8} (sin datos mapeados)")
                continue

            if DRY_RUN_ENV if AUTO_MODE else dry_run:
                ok_count += 1
                logs.append({"estado": "simulada", "campos": fields, "fila_excel": i + 8})
            else:
                ok, card_id, errors, raw = pipefy_create_card(pipe_id_int, fields, token)
                if ok and card_id:
                    ok_count += 1
                    logs.append({"estado": "ok", "card_id": card_id, "fila_excel": i + 8})
                else:
                    fail_count += 1
                    logs.append({"estado": "error", "fila_excel": i + 8, "detalle": errors or raw})
                    time.sleep(0.4)

            time.sleep(0.15)
            progress.progress(i/total, text=f"Procesadas {i}/{total} (desde fila Excel 9)")

        st.success(f"Proceso terminado. Éxitos: {ok_count} • Fallos: {fail_count} • Omitidas: {skipped}")
        st.download_button(
            "📥 Descargar log (JSON)",
            data=json.dumps(logs, ensure_ascii=False, indent=2),
            file_name="resultado_pipefy.json",
            mime="application/json"
        )
