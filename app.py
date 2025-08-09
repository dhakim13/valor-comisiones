import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# --------------------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------------------
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("MVP • Carga la base mensual y (opcional) el historial del distribuidor. Exporta un Excel con RESUMEN, ANEXO, HISTORIAL (mes), RESUMEN MES y CARTERA MES.")

# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
def normalize_dn(series: pd.Series) -> pd.Series:
    """Normaliza DNs: quita .0, notación científica y deja texto sin decimales."""
    out = series.astype(str).str.replace(r'\.0$', '', regex=True)
    def fix(x: str) -> str:
        try:
            xl = x.lower()
            if 'e+' in xl or 'e-' in xl:
                return str(int(float(x)))
            return x.split('.')[0]
        except Exception:
            return x
    return out.apply(fix)

def classify_row(row: pd.Series) -> str:
    """Clasifica producto usando TIPO (prefiere 'MOB' -> MBB) y Costo de paquete."""
    tipo = str(row.get('TIPO', '')).upper()
    costo = row.get('COSTO PAQUETE', np.nan)

    if 'MOB' in tipo:
        return 'MBB'
    # Mapas por costos (ajustables si cambian)
    if pd.notna(costo):
        try:
            c = float(costo)
        except Exception:
            c = np.nan
        if c in [99, 115, 349, 399, 439, 500]:
            return 'HBB'
        if c in [110, 120, 160, 245, 375, 480, 620]:
            return 'MiFi'
    # Default
    return 'MBB'

def cartera_pct_mbb(n_altas_mes: int) -> float:
    """Porcentaje de cartera MBB según volumen de altas del mes."""
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

def safe_get_col(df: pd.DataFrame, candidates, default=None):
    """Devuelve el primer nombre de columna existente en df de la lista 'candidates'."""
    for c in candidates:
        if c in df.columns:
            return c
    return default

def load_historial_distribuidor(file) -> pd.DataFrame:
    """
    Lee la hoja HISTORIAL DE ACTIVACIONES del reporte del distribuidor (formateado),
    usando header=7 (fila 8 visible). Devuelve DataFrame con columnas estandarizadas:
    DN (str), FECHA_ACTIVACION (datetime).
    """
    try:
        hist = pd.read_excel(file, sheet_name="HISTORIAL DE ACTIVACIONES", header=7, engine="openpyxl")
    except Exception:
        # fallback sin engine explícito
        hist = pd.read_excel(file, sheet_name="HISTORIAL DE ACTIVACIONES", header=7)

    # Normalizar nombres
    hist.columns = [str(c).strip() for c in hist.columns]

    # Detectar columnas
    dn_col = safe_get_col(hist, ["DN", "DN "])
    fa_col = safe_get_col(hist, ["FECHA DE ACTIVACION", "Fecha de Activacion", "FECHA_ACTIVACION"])

    # Filtrar filas válidas
    if dn_col is None:
        return pd.DataFrame(columns=["DN", "FECHA_ACTIVACION"])
    hist = hist[[dn_col] + ([fa_col] if fa_col else [])].copy()

    # Renombrar
    hist = hist.rename(columns={dn_col: "DN"})
    if fa_col:
        hist = hist.rename(columns={fa_col: "FECHA_ACTIVACION"})
        hist["FECHA_ACTIVACION"] = pd.to_datetime(hist["FECHA_ACTIVACION"], errors="coerce")
    else:
        hist["FECHA_ACTIVACION"] = pd.NaT

    # Normalizar DN
    hist["DN_NORM"] = normalize_dn(hist["DN"])
    hist = hist.dropna(subset=["DN_NORM"])
    return hist

# --------------------------------------------------------------------------------------
# Cálculo principal
# --------------------------------------------------------------------------------------
def calc_report(df_tot: pd.DataFrame,
                df_rec: pd.DataFrame,
                dist_name: str,
                year: int,
                month: int,
                df_hist: pd.DataFrame | None = None) -> BytesIO:

    # Ventana mes natural
    month_start = pd.Timestamp(year, month, 1)
    month_end   = pd.Timestamp(year, month, 1) + pd.offsets.MonthEnd(1)

    # Copias
    df_tot = df_tot.copy()
    df_rec = df_rec.copy()

    # Normalización de fechas
    # En "Desgloce Totales" la fecha de activación suele estar en 'FECHA'
    if "FECHA" in df_tot.columns:
        df_tot["FECHA"] = pd.to_datetime(df_tot["FECHA"], errors="coerce")
    else:
        # Intento alterno por si el nombre varía
        alt = safe_get_col(df_tot, ["FECHA DE ACTIVACION", "Fecha", "Fecha de Activacion"])
        if alt:
            df_tot["FECHA"] = pd.to_datetime(df_tot[alt], errors="coerce")
        else:
            df_tot["FECHA"] = pd.NaT

    # En "Desgloce Recarga" la fecha suele estar en 'FECHA'
    if "FECHA" in df_rec.columns:
        df_rec["FECHA"] = pd.to_datetime(df_rec["FECHA"], errors="coerce")
    else:
        alt = safe_get_col(df_rec, ["Fecha", "FECHA_RECARGA"])
        if alt:
            df_rec["FECHA"] = pd.to_datetime(df_rec[alt], errors="coerce")
        else:
            df_rec["FECHA"] = pd.NaT

    # Normalización DN
    dn_tot_col = safe_get_col(df_tot, ["DN", "DN "])
    dn_rec_col = safe_get_col(df_rec, ["DN", "DN "])

    if dn_tot_col is None or dn_rec_col is None:
        raise ValueError("No encuentro la columna 'DN' en las hojas base.")

    df_tot["DN_NORM"] = normalize_dn(df_tot[dn_tot_col])
    df_rec["DN_NORM"] = normalize_dn(df_rec[dn_rec_col])

    # Columna de distribuidor en Totales (a veces viene con espacio final)
    dist_col = safe_get_col(df_tot, ["DISTRIBUIDOR ", "DISTRIBUIDOR", "Distribuidor"])
    if dist_col is None:
        raise ValueError("No encuentro la columna 'DISTRIBUIDOR' en 'Desgloce Totales'.")

    # Filtro por distribuidor (case-insensitive, trim)
    mask_dist = df_tot[dist_col].astype(str).str.strip().str.lower() == dist_name.strip().lower()
    tot_dist = df_tot[mask_dist].copy()

    # Si viene historial del distribuidor, añadimos sus DN al universo
    dns_tot = set(tot_dist["DN_NORM"].dropna())
    if df_hist is not None and not df_hist.empty:
        dns_hist = set(df_hist["DN_NORM"].dropna())
    else:
        dns_hist = set()

    dns_universo = dns_tot.union(dns_hist)

    # Altas del mes (desde Totales)
    altas_mes = tot_dist[(tot_dist["FECHA"] >= month_start) & (tot_dist["FECHA"] <= month_end)].copy()

    # Recargas del mes (cruzadas al universo de ese distribuidor)
    rec_month = df_rec[(df_rec["FECHA"] >= month_start) & (df_rec["FECHA"] <= month_end)].copy()
    rec_month_dist = rec_month[rec_month["DN_NORM"].isin(dns_universo)].copy()

    # Clasificación de producto en tot_dist
    if "PRODUCTO" not in tot_dist.columns:
        tot_dist["PRODUCTO"] = tot_dist.apply(classify_row, axis=1)

    # ----- Reglas (MVP) -----
    # MBB
    n_altas = int(altas_mes["DN_NORM"].nunique())
    pct_mbb = cartera_pct_mbb(n_altas)
    min_mbb = 35

    # MiFi / HBB mínimos
    min_mifi = 110
    min_hbb = 99

    # Suma de recargas por línea en el mes
    if "MONTO" not in rec_month_dist.columns:
        # Intento alterno
        monto_col = safe_get_col(rec_month_dist, ["MONTO $", "Monto", "IMPORTE"])
        if monto_col:
            rec_month_dist = rec_month_dist.rename(columns={monto_col: "MONTO"})
        else:
            rec_month_dist["MONTO"] = 0.0

    rec_by_dn = (
        rec_month_dist.groupby("DN_NORM", as_index=False)["MONTO"]
        .sum()
        .rename(columns={"MONTO": "RECARGA_TOTAL_MES"})
    )

    # ANEXO: partimos de tot_dist (DN, fechas, plan, costo, producto) y unimos recargas del mes
    plan_col = safe_get_col(tot_dist, ["PLAN", "Paquete", "PAQUETE"])
    costo_col = safe_get_col(tot_dist, ["COSTO PAQUETE", "Costo Paquete", "COSTO_PAQUETE"])

    cols_for_anexo = ["DN_NORM", "FECHA", "PRODUCTO"]
    show_dn = dn_tot_col if dn_tot_col != "DN_NORM" else "DN"
    if show_dn not in tot_dist.columns:
        tot_dist["DN"] = tot_dist["DN_NORM"]
        show_dn = "DN"

    anexo_base_cols = [show_dn, "DN_NORM", "FECHA"]
    if plan_col:  anexo_base_cols.append(plan_col)
    if costo_col: anexo_base_cols.append(costo_col)
    anexo_base_cols.append("PRODUCTO")

    anexo = (
        tot_dist[anexo_base_cols]
        .merge(rec_by_dn, on="DN_NORM", how="left")
        .copy()
    )
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # Elegibilidad por producto
    def elegible(row):
        p = row["PRODUCTO"]
        rec = row["RECARGA_TOTAL_MES"]
        if p == "MBB":
            return rec >= min_mbb
        elif p == "MiFi":
            return rec >= min_mifi
        elif p == "HBB":
            return rec >= min_hbb
        return False

    anexo["ELEGIBLE_CARTERA"] = anexo.apply(elegible, axis=1)

    # % aplicado
    def pct_aplicado(row):
        if row["PRODUCTO"] == "MBB":
            return pct_mbb
        elif row["PRODUCTO"] in ("MiFi", "HBB"):
            return 0.05  # 5% M1–12 en ambos (base)
        return 0.0

    anexo["% CARTERA APLICADA"] = anexo.apply(pct_aplicado, axis=1)
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"],
        anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"],
        0.0
    ).round(2)

    # RESUMEN (una fila)
    total_recargas_mes = float(rec_month_dist["MONTO"].sum()) if "MONTO" in rec_month_dist.columns else 0.0
    resumen = pd.DataFrame([{
        "Distribuidor": dist_name,
        "Mes": f'{month_start.strftime("%B").capitalize()} {year}',
        "Altas del mes": n_altas,
        "Recargas totales del mes ($)": round(total_recargas_mes, 2),
        "Porcentaje Cartera aplicado (MBB)": pct_mbb,
        "Comisión Cartera total ($)": round(float(anexo["COMISION_CARTERA_$"].sum()), 2)
    }])

    # RESUMEN MES (por producto) usando dict en agg (evita el error de kwargs)
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

    # HISTORIAL ACTIVACIONES (solo mes, desde Totales)
    hist_cols = ["FECHA", "DN_NORM"]
    if plan_col:  hist_cols.append(plan_col)
    if costo_col: hist_cols.append(costo_col)
    hist = (
        altas_mes[hist_cols]
        .rename(columns={"DN_NORM": "DN"})
        .sort_values("FECHA")
        .reset_index(drop=True)
    )

    # CARTERA MES (detalle recargas del mes)
    forma_col = safe_get_col(rec_month_dist, ["FORMA DE PAGO", "Forma de Pago", "FORMA_PAGO"])
    plan_rec_col = safe_get_col(rec_month_dist, ["PLAN", "PAQUETE", "Paquete"])

    rec_det_cols = ["FECHA", "DN_NORM", "MONTO"]
    if plan_rec_col: rec_det_cols.append(plan_rec_col)
    if forma_col:     rec_det_cols.append(forma_col)

    rec_det = rec_month_dist[rec_det_cols].copy()
    rec_det["ELEGIBLE_MBB"] = rec_det["MONTO"] >= min_mbb
    rec_det = (
        rec_det
        .rename(columns={"DN_NORM": "DN"})
        .sort_values("FECHA")
        .reset_index(drop=True)
    )

    # Export a Excel (en memoria)
    fname_month = month_start.strftime("%B").upper()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        anexo.to_excel(writer, sheet_name="ANEXO", index=False)
        hist.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)
        resumen_mes.to_excel(writer, sheet_name=f"RESUMEN {fname_month} {year}", index=False)
        rec_det.to_excel(writer, sheet_name=f"CARTERA {fname_month} {year}", index=False)
    output.seek(0)
    return output

# --------------------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    base_file = st.file_uploader("1) Base mensual (VT Reporte Comercial…)", type=["xlsx"])
    st.caption("Debe traer: 'Desgloce Totales' (encabezado en fila 2) y 'Desgloce Recarga' (encabezado en fila 4).")

with col2:
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.markdown("**Opcional:** sube el reporte del distribuidor para leer `HISTORIAL DE ACTIVACIONES` y ampliar el universo de líneas.")
hist_file = st.file_uploader("2) Reporte del distribuidor (opcional)", type=["xlsx"], key="hist")

if base_file and st.button("Generar reporte"):
    try:
        # Leer base mensual
        try:
            xls = pd.ExcelFile(base_file, engine="openpyxl")
        except Exception:
            xls = pd.ExcelFile(base_file)

        if ("Desgloce Totales" not in xls.sheet_names) or ("Desgloce Recarga" not in xls.sheet_names):
            st.error("El archivo base debe contener las hojas 'Desgloce Totales' y 'Desgloce Recarga'.")
        else:
            # Encabezados: Totales(header=1), Recarga(header=3)
            try:
                df_tot = pd.read_excel(base_file, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
                df_rec = pd.read_excel(base_file, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")
            except Exception:
                df_tot = pd.read_excel(base_file, sheet_name="Desgloce Totales", header=1)
                df_rec = pd.read_excel(base_file, sheet_name="Desgloce Recarga", header=3)

            # Historial opcional
            df_hist = None
            if hist_file is not None:
                df_hist = load_historial_distribuidor(hist_file)

            # Calcular
            buf = calc_report(
                df_tot=df_tot,
                df_rec=df_rec,
                dist_name=dist,
                year=int(year),
                month=int(month),
                df_hist=df_hist
            )

            # Descargar
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {year}.xlsx"
            st.success("✅ Reporte generado.")
            st.download_button(
                "⬇️ Descargar Excel",
                data=buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.exception(e)
