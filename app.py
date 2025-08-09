import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")

st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y la plantilla del distribuidor. Se actualizan RESUMEN, ANEXO, HISTORIAL (mes), RESUMEN {MES} y CARTERA {MES} sobre la plantilla, preservando el resto de hojas.")

# ---------- Helpers ----------
def normalize_dn(series: pd.Series) -> pd.Series:
    out = series.astype(str).str.replace(r"\.0$", "", regex=True)
    def fix(x: str) -> str:
        try:
            s = str(x)
            if "e+" in s.lower():
                return str(int(float(s)))
            return s.split(".")[0]
        except Exception:
            return s
    return out.apply(fix)

def classify_row(row: pd.Series) -> str:
    # Reglas de clasificación (MVP); ajustables si cambian los costos de paquete
    tipo = str(row.get("TIPO", "")).upper()
    costo = row.get("COSTO PAQUETE", np.nan)
    if "MOB" in tipo:
        return "MBB"
    # HBB por costos comunes
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    # MiFi por costos comunes
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    return "MBB"

def cartera_pct_mbb(n_altas_mes: int) -> float:
    if n_altas_mes < 50:
        return 0.03
    elif n_altas_mes < 300:
        return 0.05
    elif n_altas_mes < 500:
        return 0.07
    elif n_altas_mes < 1000:
        return 0.08
    else:
        return 0.10

def get_col(df: pd.DataFrame, candidates) -> str:
    """Devuelve el nombre de la primera columna existente en df de la lista candidates."""
    for c in candidates:
        if c in df.columns:
            return c
    # fallback: intenta buscar por lower/strip
    df_map = {str(c).strip().lower(): c for c in df.columns}
    for c in candidates:
        key = str(c).strip().lower()
        if key in df_map:
            return df_map[key]
    raise KeyError(f"No se encontró ninguna de estas columnas: {candidates}")

def calc_report(df_tot: pd.DataFrame, df_rec: pd.DataFrame, dist_name: str, year: int, month: int) -> BytesIO:
    month_start = pd.Timestamp(year, month, 1)
    month_end   = month_start + pd.offsets.MonthEnd(1)

    # Normalización copias
    df_tot = df_tot.copy()
    df_rec = df_rec.copy()

    # Columnas claves (tolerantes a espacios/variantes)
    col_fecha_tot = get_col(df_tot, ["FECHA"])
    col_dn_tot    = get_col(df_tot, ["DN"])
    col_plan      = get_col(df_tot, ["PLAN"])
    col_costo     = get_col(df_tot, ["COSTO PAQUETE", "COSTO PAQ", "COSTO"])
    col_tipo      = get_col(df_tot, ["TIPO"])
    col_dist      = get_col(df_tot, ["DISTRIBUIDOR ", "DISTRIBUIDOR"])

    col_fecha_rec = get_col(df_rec, ["FECHA"])
    col_dn_rec    = get_col(df_rec, ["DN"])
    col_monto_rec = get_col(df_rec, ["MONTO", "MONTO RECARGA", "RECARGA"])
    col_fpago     = get_col(df_rec, ["FORMA DE PAGO", "FP", "PAGO"])

    # Tipos de datos
    df_tot[col_fecha_tot] = pd.to_datetime(df_tot[col_fecha_tot], errors="coerce")
    df_rec[col_fecha_rec] = pd.to_datetime(df_rec[col_fecha_rec], errors="coerce")

    df_tot["DN_NORM"] = normalize_dn(df_tot[col_dn_tot])
    df_rec["DN_NORM"] = normalize_dn(df_rec[col_dn_rec])

    # Filtro por distribuidor
    mask_dist = df_tot[col_dist].astype(str).str.strip().str.lower() == dist_name.strip().lower()
    tot_dist = df_tot[mask_dist].copy()
    dns_dist = set(tot_dist["DN_NORM"].dropna())

    # Activaciones del mes
    altas_mes = tot_dist[(tot_dist[col_fecha_tot] >= month_start) & (tot_dist[col_fecha_tot] <= month_end)].copy()

    # Recargas del mes (del universo de ese distribuidor)
    rec_month = df_rec[(df_rec[col_fecha_rec] >= month_start) & (df_rec[col_fecha_rec] <= month_end)].copy()
    rec_month_dist = rec_month[rec_month["DN_NORM"].isin(dns_dist)].copy()

    # Clasificación producto
    # nos aseguramos de que existan columnas que requiere classify_row
    if "TIPO" not in tot_dist.columns:
        tot_dist["TIPO"] = tot_dist[col_tipo]
    if "COSTO PAQUETE" not in tot_dist.columns:
        tot_dist["COSTO PAQUETE"] = tot_dist[col_costo]
    tot_dist["PRODUCTO"] = tot_dist.apply(classify_row, axis=1)

    # ----- Reglas de negocio -----
    n_altas = altas_mes["DN_NORM"].nunique()
    pct_mbb = cartera_pct_mbb(int(n_altas))

    min_mbb  = 35
    min_mifi = 110
    min_hbb  = 99

    # Suma de recargas por línea en el mes
    rec_by_dn = (
        rec_month_dist.groupby("DN_NORM", as_index=False)[col_monto_rec]
        .sum()
        .rename(columns={col_monto_rec: "RECARGA_TOTAL_MES"})
    )

    # ANEXO
    anexo = (
        tot_dist[[col_dn_tot, "DN_NORM", col_fecha_tot, col_plan, "COSTO PAQUETE", "PRODUCTO"]]
        .rename(columns={col_dn_tot: "DN", col_fecha_tot: "FECHA"})
        .merge(rec_by_dn, on="DN_NORM", how="left")
    )
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    def elegible(row):
        if row["PRODUCTO"] == "MBB":
            return row["RECARGA_TOTAL_MES"] >= min_mbb
        elif row["PRODUCTO"] == "MiFi":
            return row["RECARGA_TOTAL_MES"] >= min_mifi
        elif row["PRODUCTO"] == "HBB":
            return row["RECARGA_TOTAL_MES"] >= min_hbb
        return False

    anexo["ELEGIBLE_CARTERA"] = anexo.apply(elegible, axis=1)

    def pct_aplicado(row):
        if row["PRODUCTO"] == "MBB":
            return pct_mbb
        elif row["PRODUCTO"] in ("MiFi", "HBB"):
            return 0.05  # base 5% M1–12
        return 0.0

    anexo["% CARTERA APLICADA"] = anexo.apply(pct_aplicado, axis=1)
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"], anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"], 0.0
    ).round(2)

    # RESUMEN (una fila)
    resumen = pd.DataFrame([{
        "Distribuidor": dist_name,
        "Mes": f'{month_start.strftime("%B").capitalize()} {year}',
        "Altas del mes": int(n_altas),
        "Recargas totales del mes ($)": round(float(rec_month_dist[col_monto_rec].sum() if not rec_month_dist.empty else 0.0), 2),
        "Porcentaje Cartera aplicado (MBB)": float(pct_mbb),
        "Comisión Cartera total ($)": round(float(anexo["COMISION_CARTERA_$"].sum()), 2),
    }])

    # RESUMEN MES (por producto + total)
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
        .agg({
            "DN_NORM": "nunique",
            "RECARGA_TOTAL_MES": "sum",
            "COMISION_CARTERA_$": "sum"
        })
        .rename(columns={
            "DN_NORM": "Lineas",
            "RECARGA_TOTAL_MES": "Recarga_Mes_$",
            "COMISION_CARTERA_$": "Comision_Mes_$"
        })
    )

    total_row = pd.DataFrame([{
        "PRODUCTO": "TOTAL",
        "Lineas": resumen_mes["Lineas"].sum(),
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
        "Comision_Mes_$": resumen_mes["Comision_Mes_$"].sum()
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # HISTORIAL ACTIVACIONES (solo mes)
    hist = (
        altas_mes[[col_fecha_tot, "DN_NORM", col_plan, col_costo]]
        .rename(columns={col_fecha_tot: "FECHA", "DN_NORM": "DN", col_costo: "COSTO PAQUETE"})
        .sort_values("FECHA")
    )

    # CARTERA MES (detalle recargas del mes)
    rec_det = rec_month_dist.copy()
    if not rec_det.empty:
        rec_det["ELEGIBLE_MBB"] = rec_det[col_monto_rec] >= min_mbb
        rec_det = (
            rec_det[[col_fecha_rec, "DN_NORM", col_plan, col_monto_rec, col_fpago, "ELEGIBLE_MBB"]]
            .rename(columns={col_fecha_rec: "FECHA", "DN_NORM": "DN", col_monto_rec: "MONTO", col_fpago: "FORMA DE PAGO"})
            .sort_values("FECHA")
        )
    else:
        rec_det = pd.DataFrame(columns=["FECHA", "DN", "PLAN", "MONTO", "FORMA DE PAGO", "ELEGIBLE_MBB"])

    # Export preliminar (para re-leer y volcar en la plantilla)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        anexo.to_excel(writer, sheet_name="ANEXO", index=False)
        hist.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)
        resumen_mes.to_excel(writer, sheet_name=f'RESUMEN {month_start.strftime("%B").upper()} {year}', index=False)
        rec_det.to_excel(writer, sheet_name=f'CARTERA {month_start.strftime("%B").upper()} {year}', index=False)
    output.seek(0)
    return output

def write_like_template(template_bytes, new_sheets: dict) -> BytesIO:
    """
    Carga la plantilla (archivo Excel con historial y formato),
    reemplaza/crea solo las hojas pasadas en new_sheets y devuelve BytesIO final.
    """
    import openpyxl
    from openpyxl.utils.dataframe import dataframe_to_rows

    # Cargar plantilla
    if isinstance(template_bytes, BytesIO):
        template_bytes.seek(0)
        wb = openpyxl.load_workbook(template_bytes)
    else:
        wb = openpyxl.load_workbook(template_bytes)

    # Reemplazar/crear las hojas de mes
    for sh_name, df in new_sheets.items():
        if sh_name in wb.sheetnames:
            ws_old = wb[sh_name]
            wb.remove(ws_old)
        ws = wb.create_sheet(sh_name)
        # Volcar DataFrame (sin estilos; la plantilla manda para el resto de hojas)
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)
        # Ajuste simple de ancho
        for col_cells in ws.columns:
            try:
                col_letter = col_cells[0].column_letter
                max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
                ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)
            except Exception:
                pass

    # Guardar a bytes
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ---------- UI ----------
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("1) Base mensual (VT Reporte Comercial...)", type=["xlsx"])
    plantilla_file = st.file_uploader("2) Plantilla/Historial del distribuidor (.xlsx)", type=["xlsx"])
    st.caption("Base: debe tener 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4). La plantilla es el XLSX del distribuidor con todas sus pestañas.")

with col2:
    st.write("Parámetros")
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.divider()
st.markdown("**Paso a paso:** 1) Sube los **dos** archivos. 2) Verifica distribuidor/año/mes. 3) Clic en **Generar reporte**. 4) Descarga el Excel final.")

if base_file is not None and plantilla_file is not None and st.button("Generar reporte"):
    try:
        # Validación rápida de hojas base
        xls = pd.ExcelFile(base_file, engine="openpyxl")
        if "Desgloce Totales" not in xls.sheet_names or "Desgloce Recarga" not in xls.sheet_names:
            st.error("El archivo base debe contener las hojas 'Desgloce Totales' y 'Desgloce Recarga'.")
        else:
            # Lee con encabezados correctos
            df_tot = pd.read_excel(base_file, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
            df_rec = pd.read_excel(base_file, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")

            # Calcula hojas del mes (en memoria)
            buf_calc = calc_report(df_tot, df_rec, dist.strip(), int(year), int(month))
            month_str = datetime(int(year), int(month), 1).strftime("%B").upper()

            # Relee las 5 hojas calculadas para volcarlas en la plantilla
            with pd.ExcelFile(buf_calc, engine="openpyxl") as xf:
                df_resumen  = pd.read_excel(xf, sheet_name="RESUMEN")
                df_anexo    = pd.read_excel(xf, sheet_name="ANEXO")
                df_hist     = pd.read_excel(xf, sheet_name="HISTORIAL DE ACTIVACIONES")
                df_res_mes  = pd.read_excel(xf, sheet_name=f"RESUMEN {month_str} {int(year)}")
                df_cart_mes = pd.read_excel(xf, sheet_name=f"CARTERA {month_str} {int(year)}")

            new_sheets = {
                "RESUMEN": df_resumen,
                "ANEXO": df_anexo,
                "HISTORIAL DE ACTIVACIONES": df_hist,
                f"RESUMEN {month_str} {int(year)}": df_res_mes,
                f"CARTERA {month_str} {int(year)}": df_cart_mes
            }

            final_buf = write_like_template(plantilla_file, new_sheets)
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {month_str} {int(year)}.xlsx"

            st.success("✅ Reporte armado sobre la plantilla del distribuidor.")
            st.download_button(
                "⬇️ Descargar Excel final",
                data=final_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.exception(e)
