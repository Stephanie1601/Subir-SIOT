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
.block-container {
    max-width: 1200px;
    padding-top: 0rem !important;   /* eliminar espacio superior */
    margin-top: 0 !important;
}
div.stButton > button {
    border-radius: 12px;
    padding: 0.6rem 1rem;
    font-weight: 600;
}
.stSidebar, .sidebar .sidebar-content {
    background: linear-gradient(180deg, #fafafa, #f0f0f0);
}
.kpi {
    padding: 12px 16px;
    border-radius: 14px;
    background: #ffffff;
    box-shadow: 0 3px 12px rgba(0,0,0,.06);
    border: 1px solid #eee;
}
.section {
    padding: 16px;
    border-radius: 16px;
    background: #ffffff;
    border: 1px solid #ececec;
    box-shadow: 0 3px 16px rgba(0,0,0,.05);
}
h1, h2, h3 { font-weight: 800; }
small.help { color: #666; }
</style>
''', unsafe_allow_html=True)

# ---------- SIMPLE AUTH ----------
AUTH_USERS = json.loads(os.environ.get("AUTH_USERS_JSON", os.getenv("AUTH_USERS_JSON", '{"admin":"admin"}')))

# ---------- LOGO HELPERS ----------
def _find_logo_bytes() -> bytes | None:
    """
    Intenta cargar el logo desde rutas conocidas.
    Devuelve bytes del PNG si lo encuentra; en caso contrario None.
    """
    candidates = [
        Path("/mnt/data/06ccb9c2-ca99-49b6-a58e-9452a7e6a452.png"),  # ruta confirmada
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
    """
    Renderiza el logo SIEMPRE: si hay archivo lo incrusta como data URL Base64.
    Esto evita problemas de rutas en Streamlit Cloud.
    """
    img_bytes = _find_logo_bytes()
    if not img_bytes:
        st.info("No se encontró el logo de EOMMT en el servidor.")
        return
    b64 = base64.b64encode(img_bytes).decode("ascii")
    st.markdown(
        f"""
        <div style="text-align:center; margin: 6px 0 4px 0;">
            <img src="data:image/png;base64,{b64}" width="{width_px}" />
        </div>
        """,
        unsafe_allow_html=True
    )

def render_logo_sidebar(width_px: int = 160):
    """
    Igual que arriba pero pensado para la barra lateral.
    """
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
        render_logo_center(width_px=200)
        st.markdown("### 🚇 Instrucción Operacional de Trabajos")
        st.markdown("## 🔐 Ingreso al sistema")
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

# ---------- HELPERS ----------
def read_excel_table(uploaded_bytes: bytes, table_name: str = "SIOT") -> pd.DataFrame:
    """
    Busca una tabla estructurada llamada 'SIOT' en el archivo; si no la encuentra,
    usa la primera hoja como respaldo.
    """
    bio = io.BytesIO(uploaded_bytes)
    wb = load_workbook(bio, data_only=True, read_only=True)

    # 1) Intentar encontrar tabla estructurada "SIOT"
    for ws in wb.worksheets:
        tables = getattr(ws, "_tables", {}) or {}
        for t in tables.values():
            if t.name and t.name.lower() == table_name.lower():
                ref = t.ref  # p.ej. "A1:K300"
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
                df = pd.DataFrame(body, columns=header)
                return df

    # 2) Respaldo: primera hoja
    bio.seek(0)
    xls = pd.ExcelFile(bio, engine="openpyxl")
    first = xls.sheet_names[0]
    df = pd.read_excel(bio, sheet_name=first, engine="openpyxl")
    return df

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Limpieza básica: nombres de columnas, elimina filas/columnas vacías, trim a strings."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df = df.dropna(how="all")
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def build_fields_attributes(row: dict, mapping: dict) -> list:
    """
    Convierte row (dict de la fila Excel) en fields_attributes que Pipefy espera:
    [{"field_id": "...", "field_value": ...}, ...]
    Solo incluye columnas mapeadas y con valor no vacío.
    """
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
    """
    Llama a la mutación createCard de Pipefy (GraphQL).
    Devuelve (ok, card_id, errors, raw_response_text)
    """
    url = "https://api.pipefy.com/graphql"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    mutation = """
    mutation($input: CreateCardInput!) {
      createCard(input: $input) {
        card { id }
      }
    }
    """
    variables = {
        "input": {
            "pipe_id": pipe_id,
            "fields_attributes": fields_attrs
        }
    }
    try:
        resp = requests.post(
            url,
            headers=headers,
            json={"query": mutation, "variables": variables},
            timeout=60
        )
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

# ---------- APP ----------
if require_auth():
    # Logo en barra lateral (branding)
    render_logo_sidebar(width_px=150)

    with st.sidebar:
        st.subheader("🔧 Configuración Pipefy")
        pipe_id = st.text_input("Pipe ID", placeholder="Ej. 123456789")
        token = st.text_input("API Token", type="password", placeholder="Token secreto de Pipefy")
        dry_run = st.toggle("Simular (no crea tarjetas)", value=True,
                            help="Haz pruebas antes de subir definitivamente.")

    # Logo centrado en la página principal
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
        st.subheader("👀 Vista previa")
        st.dataframe(df.head(50), use_container_width=True)

        # ---------- Mapeo columnas → field_id ----------
        st.subheader("🧭 Mapeo de columnas → campos de Pipefy")

        default_mapping_literal = json.dumps(
            {str(c): "" for c in df.columns},
            ensure_ascii=False,
            indent=2
        )

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

        df_upload = df.dropna(how="all")
        st.markdown(f"**Filas detectadas con datos:** {len(df_upload)}")

        with st.expander("🔎 Filtro opcional"):
            cols = st.multiselect(
                "Columnas a mostrar en el resumen",
                options=list(df_upload.columns),
                default=list(df_upload.columns)[:6]
            )
            st.dataframe(df_upload[cols].head(100), use_container_width=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi'><b>Columnas</b><br>{len(df_upload.columns)}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi'><b>Filas totales</b><br>{len(df)}</div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><b>Filas a subir</b><br>{len(df_upload)}</div>", unsafe_allow_html=True)

        st.markdown("---")

        btn = st.button(
            "🚀 Subir a Pipefy",
            type="primary",
            use_container_width=True,
            disabled=(not pipe_id or not token)
        )

        if btn:
            if not pipe_id or not token:
                st.error("Completa el Pipe ID y el API Token.")
                st.stop()

            try:
                pipe_id_int = int(pipe_id)
            except:
                st.error("Pipe ID debe ser numérico.")
                st.stop()

            progress = st.progress(0.0, text="Iniciando...")
            logs = []
            ok_count, fail_count, skipped = 0, 0, 0
            total = len(df_upload)

            for i, (_, row) in enumerate(df_upload.iterrows(), start=1):
                row_dict = row.to_dict()
                fields = build_fields_attributes(row_dict, mapping)

                if not fields:
                    skipped += 1
                    logs.append({"estado": "omitida", "razon": "Sin campos con datos", "fila": i})
                    progress.progress(i/total, text=f"Omitida fila {i} (sin datos mapeados)")
                    continue

                if dry_run:
                    ok_count += 1
                    logs.append({"estado": "simulada", "campos": fields, "fila": i})
                else:
                    ok, card_id, errors, raw = pipefy_create_card(pipe_id_int, fields, token)
                    if ok and card_id:
                        ok_count += 1
                        logs.append({"estado": "ok", "card_id": card_id, "fila": i})
                    else:
                        fail_count += 1
                        logs.append({"estado": "error", "fila": i, "detalle": errors or raw})
                        time.sleep(0.4)

                time.sleep(0.15)
                progress.progress(i/total, text=f"Procesadas {i} / {total}")

            st.success(f"Proceso terminado. Éxitos: {ok_count} • Fallos: {fail_count} • Omitidas/Simuladas: {skipped}")
            st.download_button(
                "📥 Descargar log (JSON)",
                data=json.dumps(logs, ensure_ascii=False, indent=2),
                file_name="resultado_pipefy.json",
                mime="application/json"
            )
else:
    st.stop()
