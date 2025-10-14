# app.py
import os
import io
import re
import time
import json
import base64
import unicodedata
from pathlib import Path
from datetime import time as dtime, datetime, date

import requests
import pandas as pd
import streamlit as st

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Carga SIOT → Pipefy", page_icon="📤", layout="wide")

# ---------- THEME / CSS ----------
st.markdown('''
<style>
html, body, [class*="css"] { font-family: 'Arial', sans-serif !important; }
.block-container { max-width: 1200px; padding-top: 4.2rem !important; margin-top: 0 !important; }
div.stButton > button { border-radius: 12px; padding: 0.6rem 1rem; font-weight: 600; font-family: 'Arial', sans-serif !important; }
.stSidebar, .sidebar .sidebar-content { background: linear-gradient(180deg, #fafafa, #f0f0f0); }
.kpi { padding: 12px 16px; border-radius: 14px; background: #ffffff; box-shadow: 0 3px 12px rgba(0,0,0,.06); border: 1px solid #eee; }
.section { padding: 16px; border-radius: 16px; background: #ffffff; border: 1px solid #ececec; box-shadow: 0 3px 16px rgba(0,0,0,.05); }
h1, h2, h3 { font-family: 'Arial', sans-serif !important; font-weight: 800; }
.title-io { font-size: 2.6rem; line-height: 1.2; margin: .75rem 0 .25rem 0; color: #FF9016; font-weight: 800; }
.subtitle-login { font-size: 1.6rem; line-height: 1.2; margin: .5rem 0 .75rem 0; color: #444; font-weight: 800; }
.header-spacer { height: 15px; }
.badge { display:inline-block; padding:4px 8px; border-radius:999px; background:#eef6ff; color:#185adb; font-size:.85rem; font-weight:700; border:1px solid #d6e8ff; }
small.help { color: #666; }
</style>
''', unsafe_allow_html=True)

# ---------- SIMPLE AUTH ----------
AUTH_USERS = json.loads(os.environ.get("AUTH_USERS_JSON", os.getenv("AUTH_USERS_JSON", '{"admin":"admin"}')))

# ---------- LOGO ----------
def _find_logo_bytes() -> bytes | None:
    for p in [
        Path("/mnt/data/06ccb9c2-ca99-49b6-a58e-9452a7e6a452.png"),
        Path("/mnt/data/Logo EOMMT.png"),
        Path(__file__).parent / "Logo EOMMT.png",
        Path("Logo EOMMT.png"),
        Path("logo_eommt.png"),
        Path("assets/Logo EOMMT.png"),
    ]:
        try:
            if p.exists(): return p.read_bytes()
        except Exception: pass
    return None

def render_logo_center(width_px: int = 220):
    img = _find_logo_bytes()
    if not img: return
    b64 = base64.b64encode(img).decode("ascii")
    st.markdown(f"""<div style="text-align:center; margin: 10px 0 8px 0;">
        <img src="data:image/png;base64,{b64}" width="{width_px}" /></div>""", unsafe_allow_html=True)

def render_logo_sidebar(width_px: int = 160):
    img = _find_logo_bytes()
    if not img: return
    b64 = base64.b64encode(img).decode("ascii")
    st.sidebar.markdown(f"""<div style="text-align:center; margin: 6px 0 10px 0;">
        <img src="data:image/png;base64,{b64}" width="{width_px}" /></div>""", unsafe_allow_html=True)

# ---------- LOGIN ----------
def login_view():
    _, c, _ = st.columns([1,1,1])
    with c:
        st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
        render_logo_center(200)
        st.markdown('<h2 class="title-io">🚇 Instrucción Operacional de Trabajos</h2>', unsafe_allow_html=True)
        st.markdown('<h3 class="subtitle-login">🔐 Ingreso al sistema</h3>', unsafe_allow_html=True)
        st.write("Por favor ingresa tus credenciales para continuar:")
        user = st.text_input("Usuario", key="login_user", placeholder="Escribe tu usuario")
        pwd  = st.text_input("Contraseña", key="login_pwd", type="password", placeholder="Escribe tu contraseña")
        if st.button("Ingresar", key="btn_login", use_container_width=True):
            if user in AUTH_USERS and AUTH_USERS.get(user) == pwd:
                st.session_state['auth_user'] = user
                st.success("✅ Acceso concedido.")
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")
    return 'auth_user' in st.session_state

def require_auth():
    if 'auth_user' in st.session_state: return True
    ok = login_view()
    if not ok: st.stop()
    return True

def logout_button():
    with st.sidebar:
        if st.button("🚪 Cerrar sesión", type="secondary", use_container_width=True):
            st.session_state.pop('auth_user', None)
            st.success("Sesión cerrada.")
            st.rerun()

# ---------- NORMALIZACIÓN ----------
def normalize_label(s: str) -> str:
    if s is None: return ""
    s = str(s).replace("\xa0", " ")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s.upper()

# Sinónimos para encabezados largos → etiqueta corta esperada
HEADER_SYNONYMS = {
    "CCU (COORDINADOR DE CUADRILLA) NOMBRE APELLIDO": "CCU",
    "CCU (COORDINADOR DE CUADRILLA)": "CCU",
    "INTEGRANTES DEL EQUIPO DE CUADRILLA (NOMBRE APELLIDO - NÚMERO DE CÉDULA)": "INTEGRANTES DE CUADRILLA",
    "INTEGRANTES DEL EQUIPO DE CUADRILLA (NOMBRE APELLIDO - NUMERO DE CEDULA)": "INTEGRANTES DE CUADRILLA",
}

# ---------- Mapeo Excel -> Pipefy ----------
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
    "SELECCIONAR ETIQUETA": "seleccionar_etiqueta",
    "CORREO ELECTRÓNICO DEL SOLICITANTE": "correo_electr_nico_del_solicitante",
}
# Normalizado → field_id (incluye sinónimos)
NORM_TO_FIELD_ID = {}
for k, v in LABEL_TO_FIELD_ID.items():
    NORM_TO_FIELD_ID[normalize_label(k)] = v
for long_key, short_key in HEADER_SYNONYMS.items():
    NORM_TO_FIELD_ID[normalize_label(long_key)] = LABEL_TO_FIELD_ID[short_key]

# ---------- LECTURA EXCEL CON DETECCIÓN DE ENCABEZADO ----------
def read_excel_detect_header(uploaded_bytes: bytes, search_rows: int = 30) -> tuple[pd.DataFrame, int]:
    """
    Lee la PRIMERA hoja sin encabezado y detecta la fila de encabezados buscando 'EMPRESA'
    (normalizado) y al menos otro label conocido. Devuelve (df_con_encabezados, idx_fila_header_1based).
    """
    bio = io.BytesIO(uploaded_bytes)
    raw = pd.read_excel(bio, engine="openpyxl", header=None)

    header_row = None
    for i in range(min(search_rows, len(raw))):
        row_vals = [normalize_label(x) for x in raw.iloc[i].tolist()]
        if not any(row_vals): 
            continue
        # ¿Contiene EMPRESA u otro label fuerte?
        hits = sum(1 for val in row_vals if val in NORM_TO_FIELD_ID or val in ("EMPRESA", "CCU"))
        if hits >= 1 and ("EMPRESA" in row_vals):
            header_row = i
            break

    if header_row is None:
        # fallback: buscar la primera fila con más de 3 celdas no vacías
        for i in range(min(search_rows, len(raw))):
            row_vals = [str(x).strip() for x in raw.iloc[i].tolist() if pd.notna(x) and str(x).strip() != ""]
            if len(row_vals) >= 3:
                header_row = i
                break

    if header_row is None:
        # último fallback: asumir fila 8
        header_row = 7

    # Construir DF con esa fila como encabezado
    headers = raw.iloc[header_row].tolist()
    df = raw.iloc[header_row+1:].copy()
    df.columns = headers
    return df, header_row + 1  # 1-based índice de encabezado

# ---------- FORMATEO / PIPEFY ----------
def _fmt_value_for_pipefy(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, dtime):
        return value.strftime("%H:%M")
    return value

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
        value = _fmt_value_for_pipefy(value)
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
    variables = {"input": {"pipe_id": int(pipe_id), "fields_attributes": fields_attrs}}
    try:
        resp = requests.post(url, headers=headers, json={"query": mutation, "variables": variables}, timeout=60)
    except Exception as e:
        return False, None, [{"message": str(e)}], str(e)

    ok = (resp.status_code == 200)
    data = {}
    try: data = resp.json()
    except Exception: pass
    errors = data.get("errors")
    card_id = data.get("data", {}).get("createCard", {}).get("card", {}).get("id")
    return ok and (errors is None) and (card_id is not None), card_id, errors, resp.text

# ---------- SECRETS / VARS ----------
def get_secret(name, default=None):
    try: return st.secrets[name]
    except Exception: return os.getenv(name, default)

PIPE_ID_ENV = get_secret("PIPEFY_PIPE_ID")
TOKEN_ENV   = get_secret("PIPEFY_TOKEN")
DRY_RUN_ENV = str(get_secret("PIPEFY_DRY_RUN", "false")).lower() == "true"
AUTO_MODE   = bool(PIPE_ID_ENV and TOKEN_ENV)

# ---------- APP ----------
if require_auth():
    render_logo_sidebar(150)
    logout_button()

    if not AUTO_MODE:
        with st.sidebar:
            st.subheader("🔧 Configuración Pipefy")
            pipe_id = st.text_input("Pipe ID", placeholder="Ej. 123456789")
            token = st.text_input("API Token", type="password", placeholder="Token secreto de Pipefy")
            dry_run = st.toggle("Simular (no crea tarjetas)", value=True, help="Haz pruebas antes de subir definitivamente.")
    else:
        with st.sidebar:
            st.markdown("<span class='badge'>Modo automático (secrets)</span>", unsafe_allow_html=True)
            st.write(f"Pipe ID: **{PIPE_ID_ENV}**"); st.write("Token: **••••••••**"); st.write(f"Dry run: **{DRY_RUN_ENV}**")
        pipe_id = str(PIPE_ID_ENV); token = str(TOKEN_ENV); dry_run = DRY_RUN_ENV

    st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
    render_logo_center(220)

    st.title("📤 SIOT → Pipefy")
    st.caption("Detecta automáticamente la fila de encabezados (busca **EMPRESA**) y crea tarjetas desde la fila siguiente hasta la última con **EMPRESA**.")

    up = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

    if up is not None:
        content = up.read()

        # 1) Detectar encabezado y construir DF
        df_raw, header_row_1based = read_excel_detect_header(content, search_rows=30)

        # Limpiar columnas 'Unnamed' y espacios
        df = df_raw.copy()
        df = df.loc[:, [c for c in df.columns if str(c).strip() != "" and not str(c).startswith("Unnamed")]]
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
        df = df.dropna(how="all")

        # Mapa normalizado de columnas originales
        orig_cols = list(df.columns)
        norm_cols = [normalize_label(c) for c in orig_cols]
        norm_to_orig = dict(zip(norm_cols, orig_cols))

        # 2) Validar EMPRESA
        if "EMPRESA" not in norm_cols:
            st.error("No se encontró la columna 'EMPRESA' en la fila de encabezado detectada.\n\nEncabezados detectados: " + ", ".join([str(c) for c in orig_cols]))
            st.stop()
        emp_col = norm_to_orig["EMPRESA"]

        # 3) Limitar hasta la última fila con EMPRESA no vacía
        mask_emp = df[emp_col].astype(str).str.strip().replace({"None": "", "nan": ""}) != ""
        if not mask_emp.any():
            st.error("No hay datos debajo del encabezado en 'EMPRESA'.")
            st.stop()
        last_idx = df.index[mask_emp].max()
        df_data = df.loc[df.index.min(): last_idx].copy()

        st.subheader(f"👀 Vista previa (encabezado en fila {header_row_1based}; datos desde fila {header_row_1based+1})")
        st.dataframe(df_data.head(50), use_container_width=True)

        # 4) Mapeo AUTOMÁTICO: usar sinónimos y normalización
        auto_mapping = {}
        for ncol, orig in zip(norm_cols, orig_cols):
            if ncol in NORM_TO_FIELD_ID:
                auto_mapping[orig] = NORM_TO_FIELD_ID[ncol]

        st.markdown("**Columnas mapeadas automáticamente:** " + (", ".join(auto_mapping.keys()) if auto_mapping else "ninguna"))
        missing = [c for c in orig_cols if c not in auto_mapping and not str(c).startswith("Unnamed")]
        if missing:
            st.caption("Columnas sin mapeo (no se enviarán): " + ", ".join(missing))

        # 5) KPIs
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='kpi'><b>Columnas mapeadas</b><br>{len(auto_mapping)}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='kpi'><b>Filas totales debajo del encabezado</b><br>{len(df)}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='kpi'><b>Filas a subir</b><br>{len(df_data)}</div>", unsafe_allow_html=True)

        st.markdown("---")

        if not pipe_id or not token:
            st.error("Faltan credenciales de Pipefy. Define `PIPEFY_PIPE_ID` y `PIPEFY_TOKEN` en *secrets* o complétalos en el panel lateral.")
            st.stop()

        st.info(f"Iniciando proceso {'(simulación)' if dry_run else ''} con Pipe ID {pipe_id}…")

        try:
            int(pipe_id)
        except:
            st.error("Pipe ID debe ser numérico.")
            st.stop()

        progress = st.progress(0.0, text="Iniciando…")
        logs = []
        ok_count = fail_count = skipped = 0
        total = len(df_data)

        for i, (_, row) in enumerate(df_data.iterrows(), start=1):
            row_dict = row.to_dict()
            fields = build_fields_attributes(row_dict, auto_mapping)

            if not fields:
                skipped += 1
                logs.append({"estado": "omitida", "razon": "Sin campos mapeados con datos", "fila_excel": header_row_1based + i})
                progress.progress(i/total, text=f"Omitida fila Excel {header_row_1based + i} (sin datos mapeados)")
                continue

            if dry_run:
                ok_count += 1
                logs.append({"estado": "simulada", "campos": fields, "fila_excel": header_row_1based + i})
            else:
                ok, card_id, errors, raw = pipefy_create_card(pipe_id, fields, token)
                if ok and card_id:
                    ok_count += 1
                    logs.append({"estado": "ok", "card_id": card_id, "fila_excel": header_row_1based + i})
                else:
                    fail_count += 1
                    logs.append({"estado": "error", "fila_excel": header_row_1based + i, "detalle": errors or raw})
                    time.sleep(0.4)

            time.sleep(0.15)
            progress.progress(i/total, text=f"Procesadas {i}/{total} (desde fila Excel {header_row_1based+1})")

        st.success(f"Proceso terminado. Éxitos: {ok_count} • Fallos: {fail_count} • Omitidas: {skipped}")
        st.download_button(
            "📥 Descargar log (JSON)",
            data=json.dumps(logs, ensure_ascii=False, indent=2),
            file_name="resultado_pipefy.json",
            mime="application/json"
        )
