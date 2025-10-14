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

# ---------- HELPERS DE EXCEL / PIPEFY ----------
def read_excel_table(uploaded_bytes: bytes, table_name: str = "SIOT") -> pd.DataFrame:
    """
    Busca una tabla 'SIOT'; si no existe, usa la primera hoja.
    """
    bio = io.BytesIO(uploaded_bytes)
    wb = load_workbook(bio, data_only=True, read_only=True)

    for ws in wb.worksheets:
        tables = getattr(ws, "_tables", {}) or {}
        for t in tables.values():
            if t.name and t.name.lower() == table_name.lower():
                ref = t.ref  # "A1:K300"
                start, end = ref.split(":")
                start_col = ''.join(filter(str.isalpha, start))
                start_row = int(''.join(filter(str.isdigit, start)))
                end_col = ''.join(filter(str.isalpha, end))
                end_row = int(''.join(filter(str.isdigit, end)))
                min_col = column_index_from_string(start_col)
                max_col = column_index_from_string(end_col)
                data = []
                for r in ws.iter_rows(min_row=start_row, max_row=end_row,
                                      min_col=min_col, max_col=max_col):
                    data.append([cell.value for cell in r])
                if not data:
                    return pd.DataFrame()
                header = [str(h).strip() if h is not None else "" for h in data[0]]
                body = data[1:]
                return pd.DataFrame(body, columns=header)

    # Respaldo: primera hoja
    bio.seek(0)
    xls = pd.ExcelFile(bio, engine="openpyxl")
    first = xls.sheet_names[0]
    return pd.read_excel(bio, sheet_name=first, engine="openpyxl")

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df = df.dropna(how="all")
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

# ---------- LECTURA DE SECRETS / VARS ----------
def get_secret(name, default=None):
    try:
        return st.secrets[name]  # Streamlit Secrets
    except Exception:
        return os.getenv(name, default)  # fallback a env var si no está en secrets

PIPE_ID_ENV = get_secret("PIPEFY_PIPE_ID")
TOKEN_ENV   = get_secret("PIPEFY_TOKEN")
DRY_RUN_ENV = str(get_secret("PIPEFY_DRY_RUN", "false")).lower() == "true"
AUTO_MODE   = bool(PIPE_ID_ENV and TOKEN_ENV)

# ---------- APP ----------
if require_auth():
    # Branding sidebar + botón de cerrar sesión
    render_logo_sidebar(width_px=150)
    logout_button()

    # Mostrar panel de configuración SOLO si NO hay secrets/vars
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

    # Logo centrado arriba
    st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
    render_logo_center(width_px=220)

    st.title("📤 SIOT → Pipefy")
    st.caption("Sube tu archivo Excel, detectamos la tabla **SIOT**, y creamos una tarjeta por cada fila con datos.")

    # ---------- CARGA DE ARCHIVO ----------
    up = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

    if up is not None:
        content = up.read()
        df = read_excel_table(content, "SIOT")
        if df is None or df.empty:
            st.error("No se encontró la tabla 'SIOT' ni datos en la primera hoja. Verifica tu archivo.")
            st.stop()

        df = clean_dataframe(df)

        # ----------- SOLO DESDE FILA 9 -----------
        df_data = df.iloc[8:].copy()  # fila 9 (1-indexado) => iloc[8:]
        if df_data.empty:
            st.error("No hay datos a partir de la fila 9.")
            st.stop()

        st.subheader("👀 Vista previa (desde fila 9)")
        st.dataframe(df_data.head(50), use_container_width=True)

        # ---------- Mapeo columnas → field_id ----------
        st.subheader("🧭 Mapeo de columnas → campos de Pipefy")
        default_mapping_literal = json.dumps({str(c): "" for c in df.columns}, ensure_ascii=False, indent=2)

        mapping_json = st.text_area(
            "Pega aquí el JSON de mapeo (formato: {'ColumnaExcel': 'field_id'})",
            value=default_mapping_literal,
            height=260
        )

        try:
            mapping = json.loads(mapping_json)
            for c in df.columns:
                mapping.setdefault(str(c), "")
        except Exception as e:
            st.error(f"JSON inválido en el mapeo: {e}")
            st.stop()

        st.markdown(f"**Filas detectadas con datos (desde fila 9):** {len(df_data)}")

        # KPIs
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi'><b>Columnas</b><br>{len(df_data.columns)}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi'><b>Filas totales</b><br>{len(df)}</div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><b>Filas a subir (≥ fila 9)</b><br>{len(df_data)}</div>", unsafe_allow_html=True)

        st.markdown("---")

        if not pipe_id or not token:
            st.error("Faltan credenciales de Pipefy. Define `PIPEFY_PIPE_ID` y `PIPEFY_TOKEN` en *secrets* o complétalos en el panel lateral.")
            st.stop()

        # Iniciar proceso AUTOMÁTICO (sin botón)
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
            fields = build_fields_attributes(row_dict, mapping)

            if not fields:
                skipped += 1
                logs.append({"estado": "omitida", "razon": "Sin campos con datos", "fila_excel": i + 8})
                progress.progress(i/total, text=f"Omitida fila Excel {i+8} (sin datos mapeados)")
                continue

            if dry_run:
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
            progress.progress(i/total, text=f"Procesadas {i} / {total} (desde fila Excel 9)")

        st.success(f"Proceso terminado. Éxitos: {ok_count} • Fallos: {fail_count} • Omitidas/Simuladas: {skipped}")
        st.download_button(
            "📥 Descargar log (JSON)",
            data=json.dumps(logs, ensure_ascii=False, indent=2),
            file_name="resultado_pipefy.json",
            mime="application/json"
        )
