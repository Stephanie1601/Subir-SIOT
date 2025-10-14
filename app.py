from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import io
import pandas as pd
import streamlit as st

def read_excel_table(uploaded_bytes: bytes, table_name: str = "SIOT") -> pd.DataFrame:
    """
    Lee la tabla estructurada `table_name` (p.ej. 'SIOT').
    - Usa ws.tables (no _tables).
    - Debe abrir el workbook con read_only=False para que existan las tablas.
    - Si no la encuentra, hace fallback buscando una fila de encabezado con 'EMPRESA'.
    """
    bio = io.BytesIO(uploaded_bytes)
    # IMPORTANTE: read_only=False para que openpyxl cargue las tablas
    wb = load_workbook(bio, data_only=True, read_only=False)

    # Mostrar qué tablas ve (útil para depurar)
    vistos = []
    for ws in wb.worksheets:
        try:
            for t in (ws.tables or {}).values():
                vistos.append(f"{ws.title}:{t.name}:{t.ref}")
        except Exception:
            pass
    if vistos:
        st.info("Tablas detectadas: " + " | ".join(vistos))
    else:
        st.info("No se detectaron tablas con openpyxl.ws.tables (revisaré fallback por encabezados).")

    # 1) Intentar encontrar la tabla exacta (case-insensitive, ignora espacios)
    target = table_name.strip().lower()
    for ws in wb.worksheets:
        tables = ws.tables or {}  # dict name->Table
        for t in tables.values():
            tname = (t.name or "").strip().lower()
            if tname == target:
                # t.ref como 'B10:AH500'
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

                # limpieza suave
                df = df.loc[:, [c for c in df.columns if str(c).strip() != "" and not str(c).startswith("Unnamed")]]
                for c in df.columns:
                    if df[c].dtype == object:
                        df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
                df = df.dropna(how="all")
                return df

    # 2) Fallback: primera hoja, buscar la fila con 'EMPRESA' como encabezado
    bio.seek(0)
    raw = pd.read_excel(bio, engine="openpyxl", sheet_name=0, header=None)
    header_row = None
    # busca la palabra EMPRESA en las primeras ~40 filas
    for i in range(min(40, len(raw))):
        row_vals = [str(x).strip().upper() if pd.notna(x) else "" for x in raw.iloc[i].tolist()]
        if "EMPRESA" in row_vals:
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()  # ni tabla ni encabezado claro

    headers = raw.iloc[header_row].tolist()
    df = raw.iloc[header_row+1:].copy()
    df.columns = headers
    # limpieza
    df = df.loc[:, [c for c in df.columns if str(c).strip() != "" and not str(c).startswith("Unnamed")]]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.dropna(how="all")
    return df
