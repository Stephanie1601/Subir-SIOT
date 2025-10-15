# app.py
import os
import io
import re
import time
import json
import base64
from pathlib import Path
from datetime import datetime
import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# ===============================
# CONFIG PÁGINA / ESTILO
# ===============================
st.set_page_config(page_title="Carga SIOT → Pipefy", page_icon="📤", layout="wide")
st.markdown('''
<style>
html, body, [class*="css"] { font-family: 'Arial', sans-serif !important; }
.block-container { max-width: 1200px; padding-top: 4.0rem !important; margin-top: 0 !important; }
div.stButton > button { border-radius: 12px; padding: 0.6rem 1rem; font-weight: 600; font-family: 'Arial', sans-serif !important; }
.stSidebar, .sidebar .sidebar-content { background: linear-gradient(180deg, #fafafa, #f0f0f0); }
.kpi { padding: 12px 16px; border-radius: 14px; background: #ffffff; box-shadow: 0 3px 12px rgba(0,0,0,.06); border: 1px solid #eee; }
h1, h2, h3 { font-family: 'Arial', sans-serif !important; font-weight: 800; }
.header-spacer { height: 12px; }
.badge { display:inline-block; padding:4px 8px; border-radius:999px; background:#eef6ff; color:#185adb; font-size:.85rem; font-weight:700; border:1px solid #d6e8ff; }
small.help { color: #666; }
</style>
''', unsafe_allow_html=True)

# ===============================
# LOGO (centrado / sidebar)
# ===============================
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
        except Exception:
            pass
    return None

def render_logo_center(width_px: int = 220):
    img = _find_logo_bytes()
    if not img: return
    b64 = base64.b64encode(img).decode("ascii")
    st.markdown(f"""<div style="text-align:center; margin: 8px 0 10px 0;">
        <img src="data:image/png;base64,{b64}" width="{width_px}" /></div>""", unsafe_allow_html=True)

def render_logo_sidebar(width_px: int = 160):
    img = _find_logo_bytes()
    if not img: return
    b64 = base64.b64encode(img).decode("ascii")
    st.sidebar.markdown(f"""<div style="text-align:center; margin: 6px 0 10px 0;">
        <img src="data:image/png;base64,{b64}" width="{width_px}" /></div>""", unsafe_allow_html=True)

# ===============================
# AUTENTICACIÓN SIMPLE
# ===============================
AUTH_USERS = json.loads(os.environ.get("AUTH_USERS_JSON", os.getenv("AUTH_USERS_JSON", '{"admin":"admin"}')))

def login_view():
    _, c, _ = st.columns([1,1,1])
    with c:
        render_logo_center(200)
        st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
        st.markdown("## 🚇 Instrucción Operacional de Trabajos")
        st.markdown("### 🔐 Ingreso al sistema")
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

# ===============================
# SECRETS / PIPEFY
# ===============================
PIPEFY_API_URL = "https://api.pipefy.com/graphql"

def get_secret(name, default=None):
    try: return st.secrets[name]
    except Exception: return os.getenv(name, default)

PIPEFY_TOKEN = get_secret("PIPEFY_TOKEN", "")
PIPE_ID      = int(str(get_secret("PIPEFY_PIPE_ID", "0")) or "0")

# ===============================
# LECTURA EXCEL: tabla SIOT primero, fallback por encabezado "EMPRESA"
# ===============================
def read_excel_table_siot(uploaded_bytes: bytes, table_name: str = "SIOT") -> pd.DataFrame:
    """
    1) Intenta extraer la tabla estructurada `SIOT` usando openpyxl (read_only=False).
    2) Si no la encuentra, fallback: detecta la fila de encabezados donde esté 'EMPRESA'.
    """
    bio = io.BytesIO(uploaded_bytes)
    wb = load_workbook(bio, data_only=True, read_only=False)

    # Intentar leer tablas
    for ws in wb.worksheets:
        tables = ws.tables or {}
        for t in tables.values():
            if (t.name or "").strip().lower() == table_name.lower():
                # Rango de la tabla
                start, end = t.ref.split(":")
                start_col = ''.join(filter(str.isalpha, start))
                start_row = int(''.join(filter(str.isdigit, start)))
                end_col = ''.join(filter(str.isalpha, end))
                end_row = int(''.join(filter(str.isdigit, end)))
                min_col = column_index_from_string(start_col)
                max_col = column_index_from_string(end_col)

                data = []
                for r in ws.iter_rows(min_row=start_row, max_row=end_row,
                                      min_col=min_col, max_col=max_col, values_only=True):
                    data.append(list(r))
                if not data:
                    return pd.DataFrame()

                header = [str(h).strip() if h is not None else "" for h in data[0]]
                body = data[1:]
                df = pd.DataFrame(body, columns=header)
                # Limpieza
                df = df.loc[:, [c for c in df.columns if str(c).strip() != "" and not str(c).startswith("Unnamed")]]
                for c in df.columns:
                    if df[c].dtype == object:
                        df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
                df = df.dropna(how="all")
                return df

    # Fallback: primera hoja, buscar fila de encabezados con "EMPRESA"
    bio.seek(0)
    raw = pd.read_excel(bio, engine="openpyxl", sheet_name=0, header=None)
    header_row = None
    for i in range(min(40, len(raw))):
        row_vals = [str(x).strip().upper() if pd.notna(x) else "" for x in raw.iloc[i].tolist()]
        if "EMPRESA" in row_vals:
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()

    headers = raw.iloc[header_row].tolist()
    df = raw.iloc[header_row+1:].copy()
    df.columns = [str(c).strip() for c in headers]
    df = df.loc[:, [c for c in df.columns if c and not str(c).startswith("Unnamed")]]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.dropna(how="all")
    return df

# ===============================
# === LÓGICA DE SUBIDA (TU SEGUNDO CÓDIGO) ===
# ===============================
def _fmt_date(val):
    """Devuelve fecha en formato YYYY-MM-DD si es posible; si no, None o str(val)."""
    if val is None:
        return None
    try:
        if isinstance(val, float) and pd.isna(val):
            return None
    except Exception:
        pass
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return s

def _add_field(fields, field_id, value):
    """Agrega campo si el valor no está vacío."""
    if value is None:
        return
    try:
        if isinstance(value, float) and pd.isna(value):
            return
    except Exception:
        pass
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return
    fields.append({"field_id": field_id, "field_value": s})

def _parse_multi(val):
    """Convierte una celda en lista (para checklists o multiselect)."""
    if val is None:
        return None
    try:
        if isinstance(val, float) and pd.isna(val):
            return None
    except Exception:
        pass
    if isinstance(val, (list, tuple, set)):
        out = [str(x).strip() for x in val if str(x).strip() not in ("", "nan")]
        return out or None
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None
    parts = [p.strip() for p in s.replace(",", ";").split(";") if p.strip()]
    return parts or None

def _add_field_list(fields, field_id, value):
    items = _parse_multi(value)
    if items:
        fields.append({"field_id": field_id, "field_value": items})

def _fetch_labels_map(token: str, pipe_id: int) -> dict:
    """Devuelve un dict {nombre: id} de las etiquetas del pipe."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    q = {
        "query": """
        query($id: ID!){
          pipe(id: $id){
            labels { id name }
          }
        }
        """,
        "variables": {"id": pipe_id}
    }
    try:
        r = requests.post(PIPEFY_API_URL, headers=headers, json=q, timeout=30)
        data = r.json()
        labels = data.get("data", {}).get("pipe", {}).get("labels", []) if r.status_code == 200 else []
        return { lbl["name"]: lbl["id"] for lbl in labels if "id" in lbl and "name" in lbl }
    except Exception:
        return {}

def _add_label_select(fields, field_id, value, labels_map, report_missing):
    """Agrega label_select usando IDs. `value` puede ser str con ';' o ',' o lista."""
    items = _parse_multi(value)
    if not items:
        return
    ids = []
    for name in items:
        name = str(name).strip()
        if not name:
            continue
        if name in labels_map:
            ids.append(labels_map[name])
        else:
            report_missing.append(name)
    if ids:
        fields.append({"field_id": field_id, "field_value": ids})

# ========== crear tarjeta ==========
def pipefy_create_card(token: str, pipe_id: int, fields_attrs: list, title: str):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    mutation = {
        "query": """
        mutation($input: CreateCardInput!) {
          createCard(input: $input) {
            card { id title }
          }
        }
        """,
        "variables": {
            "input": {
                "pipe_id": pipe_id,
                "title": title,
                "fields_attributes": fields_attrs,
            }
        },
    }
    try:
        response = requests.post(PIPEFY_API_URL, headers=headers, json=mutation, timeout=40)
        if response.status_code != 200:
            return False, f"HTTP {response.status_code}: {response.text}"
        result = response.json()
        if "errors" in result:
            return False, str(result["errors"])
        return True, result.get("data", {}).get("createCard", {}).get("card", {}).get("id")
    except Exception as e:
        return False, str(e)

# ===============================
# APP
# ===============================
if require_auth():
    render_logo_sidebar(150)
    logout_button()

    with st.sidebar:
        st.subheader("🔧 Configuración Pipefy")
        token = st.text_input("API Token (st.secrets/ENV)", value=PIPEFY_TOKEN, type="password")
        pipe_id = st.text_input("Pipe ID", value=str(PIPE_ID))
        dry_run = st.toggle("Simular (no crea tarjetas)", value=True, help="Prueba sin crear tarjetas reales.")

    st.markdown('<div class="header-spacer"></div>', unsafe_allow_html=True)
    render_logo_center(220)

    st.title("📤 SIOT → Pipefy")
    st.caption("Se lee la **tabla `SIOT`** (si existe). Si no, se detecta la fila de encabezados con **EMPRESA** y se crea desde allí. Se sube hasta la última fila con **EMPRESA** no vacía.")

    up = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

    if up is not None:
        content = up.read()
        df = read_excel_table_siot(content, "SIOT")

        if df is None or df.empty:
            st.error("No se logró leer datos. Asegúrate de que la tabla se llame **SIOT** o que exista una fila de encabezados con **EMPRESA**.")
            st.stop()

        # Validar y recortar hasta la última fila con EMPRESA
        if "EMPRESA" not in df.columns:
            st.error("No se encontró la columna **EMPRESA** en los encabezados detectados.")
            st.stop()

        mask_emp = df["EMPRESA"].astype(str).str.strip().replace({"None": "", "nan": ""}) != ""
        if not mask_emp.any():
            st.error("No hay filas con EMPRESA.")
            st.stop()
        last_idx = df.index[mask_emp].max()
        df_data = df.loc[df.index.min(): last_idx].copy()

        st.subheader("👀 Vista previa")
        st.dataframe(df_data.head(50), use_container_width=True)

        # Columnas “relevantes” para considerar si una fila está vacía (opcional)
        columnas_relevantes = [
            "EMPRESA","CCU","INTEGRANTES DE CUADRILLA","CONTACTO CCU",
            "FECHA DE INICIO","FECHA DE FIN","CANTÓN / ESTACIÓN","ZONA DE ESTACIÓN",
            "DESCRIPCIÓN DE ACTIVIDAD","HORA DE INICIO","HORA DE FIN",
            "TIPO DE JORNADA","TIPO DE MANTENIMIENTO / INSPECCIÓN","N° REGISTRO DE FALLA",
            "CATEGORÍA DE RIESGO","CATEGORÍA DE TRABAJOS","DESENERGIZACIONES",
            "VEHÍCULO","ILUMINACIÓN PARCIAL","SEÑALETICA PROPIA","R1","R2","P1","P3","E1","V3","P6","P7","P8",
            "BLOQUEO DE VÍA","DESDE","HASTA","Seleccionar etiqueta","CORREO ELECTRÓNICO DEL SOLICITANTE"
        ]
        def _row_is_empty(row, cols):
            for col in cols:
                if col in row and str(row[col]).strip().lower() not in ("", "nan", "none"):
                    return False
            return True

        mask_validas = ~df_data.apply(lambda r: _row_is_empty(r, columnas_relevantes), axis=1)
        df_validas = df_data[mask_validas].copy()
        df_vacias  = df_data[~mask_validas].copy()

        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='kpi'><b>Columnas</b><br>{len(df_validas.columns)}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='kpi'><b>Filas con datos</b><br>{len(df_validas)}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='kpi'><b>Filas vacías omitidas</b><br>{len(df_vacias)}</div>", unsafe_allow_html=True)

        with st.expander("Ver filas omitidas (no se subirán)"):
            st.dataframe(df_vacias, use_container_width=True)

        if not token or not pipe_id.strip().isdigit():
            st.error("Configura correctamente el **Token** y el **Pipe ID**.")
            st.stop()

        pipe_id_int = int(pipe_id)
        labels_map  = _fetch_labels_map(token, pipe_id_int)

        if st.button("🚀 Subir a Pipefy", type="primary", use_container_width=True):
            missing_labels = []
            creadas = 0
            errores = 0

            progress = st.progress(0.0, text="Iniciando…")
            total = len(df_validas)

            for i, (_, row) in enumerate(df_validas.iterrows(), start=1):
                fields = []
                # === campos texto/fecha/select ===
                _add_field(fields, "empresa", row.get("EMPRESA"))
                _add_field(fields, "ccu", row.get("CCU"))
                _add_field(fields, "integrantes_de_cuadrilla", row.get("INTEGRANTES DE CUADRILLA"))
                _add_field(fields, "c_dula_ccu", row.get("CONTACTO CCU"))
                _add_field(fields, "fecha_de_inicio", _fmt_date(row.get("FECHA DE INICIO")))
                _add_field(fields, "fecha_de_fin", _fmt_date(row.get("FECHA DE FIN")))
                _add_field(fields, "cant_n_estaci_n", row.get("CANTÓN / ESTACIÓN"))
                _add_field(fields, "zona_de_trabajo", row.get("ZONA DE ESTACIÓN"))
                _add_field(fields, "descripci_n_de_actividad", row.get("DESCRIPCIÓN DE ACTIVIDAD"))
                _add_field(fields, "hora_de_inicio", row.get("HORA DE INICIO"))
                _add_field(fields, "hora_de_fin", row.get("HORA DE FIN"))
                _add_field(fields, "tipo_de_jornada", row.get("TIPO DE JORNADA"))
                _add_field(fields, "tipo_de_mantenimiento", row.get("TIPO DE MANTENIMIENTO / INSPECCIÓN"))
                _add_field(fields, "registro_de_incidente", row.get("N° REGISTRO DE FALLA"))
                _add_field(fields, "categor_a_de_riesgo", row.get("CATEGORÍA DE RIESGO"))
                _add_field(fields, "categor_a_de_trabajos", row.get("CATEGORÍA DE TRABAJOS"))
                _add_field(fields, "desenergizaci_n", row.get("DESENERGIZACIONES"))

                # === checklists / multiselect ===
                _add_field_list(fields, "veh_culo", row.get("VEHÍCULO"))
                _add_field_list(fields, "iluminaci_n_parcia", row.get("ILUMINACIÓN PARCIAL"))
                _add_field_list(fields, "se_aletica_propia", row.get("SEÑALETICA PROPIA"))
                _add_field_list(fields, "r1_1", row.get("R1"))
                _add_field_list(fields, "r2_1", row.get("R2"))
                _add_field_list(fields, "p1", row.get("P1"))
                _add_field_list(fields, "p3", row.get("P3"))
                _add_field_list(fields, "e1", row.get("E1"))
                _add_field_list(fields, "v3", row.get("V3"))
                _add_field_list(fields, "copy_of_se_aletica_propia", row.get("P6"))
                _add_field_list(fields, "copy_of_r1", row.get("P7"))
                _add_field_list(fields, "copy_of_p3", row.get("P8"))

                # === bloqueo de vía ===
                _add_field_list(fields, "bloqueo_de_v_a_1", row.get("BLOQUEO DE VÍA"))
                _add_field(fields, "bloqueo_desde", row.get("DESDE"))
                _add_field(fields, "hasta", row.get("HASTA"))

                # === etiquetas (label_select) ===
                _add_label_select(fields, "seleccionar_etiqueta", row.get("Seleccionar etiqueta"), labels_map, missing_labels)

                # === email solicitante ===
                _add_field(fields, "correo_electr_nico_del_solicitante", row.get("CORREO ELECTRÓNICO DEL SOLICITANTE"))

                title = str(row.get("EMPRESA") or f"Fila {i}")

                if dry_run:
                    creadas += 1
                else:
                    ok, info = pipefy_create_card(token, pipe_id_int, fields, title)
                    if ok:
                        creadas += 1
                    else:
                        errores += 1
                        st.error(f"❌ Error en fila {i}: {info}")

                time.sleep(0.10)
                progress.progress(i/total, text=f"Procesadas {i}/{total}")

            if missing_labels:
                st.warning("Estas etiquetas NO existen en el Pipe y se omitieron: " + ", ".join(sorted(set(missing_labels))))
            st.success(f"Proceso terminado. Tarjetas {'simuladas' if dry_run else 'creadas'}: {creadas} • Errores: {errores}")
