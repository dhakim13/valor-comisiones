import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# ---------- Config ----------
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("MVP • Carga la base mensual y elige distribuidor/mes. Exporta un Excel con RESUMEN, ANEXO, HISTORIAL (mes), RESUMEN MES y CARTERA MES.")

# ---------- Helpers ----------
def normalize_dn(series: pd.Series) -> pd.Series:
    out = series.astype(str).str.replace(r"\.0$", "", regex=True)
    def fix(x: str) -> str:
        try:
            if "e+" in x.lower():
                return str(int(float(x)))
            return x.split(".")[0]
        except Exception:
            return x
    return out.apply(fix)

def classify_row(row: pd.Series) -> str:
    tipo = str(row.get("TIPO", "")).upper()
    costo = row.get("COSTO PAQUETE", np.nan)
    if "MOB" in tipo:
        return "MBB"
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
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

def calc_report(df_tot: pd.DataFrame, df_rec: pd.DataFrame, dist_name: str, year: int, month: int) -> BytesIO:
    month_start = pd.Timestamp(year, month, 1)
    month_end = month_start + pd.offsets.MonthEnd(1)

    # Normalización
    df_tot = df_tot.copy()
    df_rec = df_rec.copy()
    df_tot["FECHA"] = pd.to_datetime(df_tot["FECHA"], errors="coerce")
    df_rec["FECHA"] = pd.to_datetime(df_rec["FECHA"], errors="coerce")
    df_tot["DN_NORM"] = normalize_dn(df_tot["DN"])
    df_rec["DN_NORM"] = normalize_dn(df_rec["DN"])

    # Filtro por distribuidor (la columna trae un espacio final en el nombre original)
    dist_col = "DISTRIBUIDOR "
    if dist_col not in df_tot.columns:
        # fallback por si en algún archivo viene sin espacio
        dist_col = "DISTRIBUIDOR"

    mask_dist = df_tot[dist_col].astype(str).str.strip().str.lower() == dist_name.lower().strip()
    tot_dist = df_tot[mask_dist].copy()
    dns_dist = set(tot_dist["DN_NORM"].dropna())

    # Activaciones del mes
    altas_mes = tot_dist[(tot_dist["FECHA"] >= month_start) & (tot_dist["FECHA"] <= month_end)].copy()

    # Recargas del mes (del universo de ese distribuidor)
    rec_month = df_rec[(df_rec["FECHA"] >= month_start) & (df_rec["FECHA"] <= month_end)].copy()
    rec_month_dist = rec_month[rec_month["DN_NORM"].isin(dns_dist)].copy()

    # Clasificación producto
    tot_dist["PRODUCTO"] = tot_dist.apply(classify_row, axis=1)

    # ----- Reglas -----
    n_altas = altas_mes["DN_NORM"].nunique()
    pct_mbb = cartera_pct_mbb(n_altas)
    min_mbb = 35
    min_mifi = 110
    min_hbb = 99

    # Suma de recargas por línea en el mes
    rec_by_dn = (
        rec_month_dist.groupby("DN_NORM", as_index=False)["MONTO"]
        .sum()
        .rename(columns={"MONTO": "RECARGA_TOTAL_MES"})
    )

    # ANEXO
    anexo = tot_dist[["DN", "DN_NORM", "FECHA", "PLAN", "COSTO PAQUETE", "PRODUCTO"]].merge(
        rec_by_dn, on="DN_NORM", how="left"
    )
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # Elegibilidad cartera
    def elegible(row: pd.Series) -> bool:
        if row["PRODUCTO"] == "MBB":
            return row["RECARGA_TOTAL_MES"] >= min_mbb
        elif row["PRODUCTO"] == "MiFi":
            return row["RECARGA_TOTAL_MES"] >= min_mifi
        elif row["PRODUCTO"] == "HBB":
            return row["RECARGA_TOTAL_MES"] >= min_hbb
        return False

    anexo["ELEGIBLE_CARTERA"] = anexo.apply(elegible, axis=1)

    # % aplicado
    def pct_aplicado(row: pd.Series) -> float:
        if row["PRODUCTO"] == "MBB":
            return pct_mbb
        elif row["PRODUCTO"] in ("MiFi", "HBB"):
            return 0.05  # M1–12 5%
        return 0.0

    anexo["% CARTERA APLICADA"] = anexo.apply(pct_aplicado, axis=1)
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"],
        anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"],
        0.0,
    ).round(2)

    # RESUMEN
    resumen = pd.DataFrame(
        [
            {
                "Distribuidor": dist_name,
                "Mes": f'{month_start.strftime("%B").capitalize()} {year}',
                "Altas del mes": int(n_altas),
                "Recargas totales del mes ($)": round(rec_month_dist["MONTO"].sum(), 2),
                "Porcentaje Cartera aplicado (MBB)": pct_mbb,
                "Comisión Cartera total ($)": round(anexo["COMISION_CARTERA_$"].sum(), 2),
            }
        ]
    )

    # RESUMEN MES
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
        .agg(
            DN_NORM=("DN_NORM", "nunique"),
            RECARGA_TOTAL_MES=("RECARGA_TOTAL_MES", "sum"),
            COMISION_CARTERA__=("COMISION_CARTERA_$", "sum"),
        )
        .rename(
            columns={
                "DN_NORM": "Lineas",
                "RECARGA_TOTAL_MES": "Recarga_Mes_$",
                "COMISION_CARTERA__": "Comision_Mes_$",
            }
        )
    )

    total_row = pd.DataFrame(
        [
            {
                "PRODUCTO": "TOTAL",
                "Lineas": resumen_mes["Lineas"].sum(),
                "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
                "Comision_Mes_$": resumen_mes["Comision_Mes_$"].sum(),
            }
        ]
    )
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # HISTORIAL ACTIVACIONES (solo mes)
    hist = (
        altas_mes[["FECHA", "DN_NORM", "PLAN", "COSTO PAQUETE"]]
        .rename(columns={"DN_NORM": "DN"})
        .sort_values("FECHA")
    )

    # CARTERA MES (detalle recargas)
    rec_det = rec_month_dist.copy()
    rec_det["ELEGIBLE_MBB"] = rec_det["MONTO"] >= min_mbb
    rec_det = (
        rec_det[["FECHA", "DN_NORM", "PLAN", "MONTO", "FORMA DE PAGO", "ELEGIBLE_MBB"]]
        .rename(columns={"DN_NORM": "DN"})
        .sort_values("FECHA")
    )

    # Export to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        anexo.to_excel(writer, sheet_name="ANEXO", index=False)
        hist.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)
        resumen_mes.to_excel(
            writer, sheet_name=f'RESUMEN {month_start.strftime("%B").upper()} {year}', index=False
        )
        rec_det.to_excel(
            writer, sheet_name=f'CARTERA {month_start.strftime("%B").upper()} {year}', index=False
        )
    output.seek(0)
    return output

# ---------- UI ----------
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial...)", type=["xlsx"])
    st.caption("Debe contener 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")

with col2:
    st.write("Parámetros")
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

# --- DEBUG PANEL opcional ---
with st.expander("🔧 Verificación rápida (debug)", expanded=False):
    dbg_file = st.file_uploader("Sube el MISMO archivo aquí para inspección", type=["xlsx"], key="dbg")
    if dbg_file:
        try:
            x = pd.ExcelFile(dbg_file, engine="openpyxl")
            st.write("Hojas encontradas:", x.sheet_names)
            try:
                t_preview = pd.read_excel(dbg_file, sheet_name='Desgloce Totales', header=1, nrows=5, engine="openpyxl")
                r_preview = pd.read_excel(dbg_file, sheet_name='Desgloce Recarga', header=3, nrows=5, engine="openpyxl")
                st.write("Totales (5 filas):"); st.dataframe(t_preview)
                st.write("Recargas (5 filas):"); st.dataframe(r_preview)
            except Exception as e:
                st.warning("No se pudo leer con encabezados esperados (2/4).")
                st.exception(e)
        except Exception as e:
            st.exception(e)

# Botón principal
if base_file is not None and st.button("Generar reporte"):
    try:
        xls = pd.ExcelFile(base_file, engine="openpyxl")
        if "Desgloce Totales" not in xls.sheet_names or "Desgloce Recarga" not in xls.sheet_names:
            st.error("El archivo base debe contener las hojas 'Desgloce Totales' y 'Desgloce Recarga'.")
        else:
            df_tot = pd.read_excel(base_file, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
            df_rec = pd.read_excel(base_file, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")
            buf = calc_report(df_tot, df_rec, dist, int(year), int(month))
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {year}.xlsx"
            st.success("Reporte generado.")
            st.download_button(
                "⬇️ Descargar Excel",
                data=buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.exception(e)
