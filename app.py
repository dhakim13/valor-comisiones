import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y la plantilla/historial. Genera un Excel clonando encabezados y calculando comisiones (Cartera +M2, 1ra Recarga, Portabilidad).")

# =========================
# Helpers
# =========================

def normalize_dn(series: pd.Series) -> pd.Series:
    out = series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    def fix(x):
        try:
            if "e+" in x.lower():
                return str(int(float(x)))
            return x.split(".")[0]
        except:
            return x
    return out.apply(fix)

def first_existing_col(df: pd.DataFrame, candidates) -> str:
    # Busca por igualdad (case-insensitive y strip) y también por "contiene"
    colmap = {c.strip().upper(): c for c in df.columns if isinstance(c, str)}
    for cand in candidates:
        key = cand.strip().upper()
        if key in colmap:
            return colmap[key]
    # intento por "contiene"
    up_cols = {c: c.upper() for c in df.columns if isinstance(c, str)}
    for cand in candidates:
        needle = cand.strip().upper()
        for original, uppered in up_cols.items():
            if needle in uppered:
                return original
    raise KeyError(f"No encontré ninguna de las columnas {candidates} en: {list(df.columns)}")

def safe_to_datetime(s: pd.Series):
    return pd.to_datetime(s, errors="coerce")

def month_diff(d_when: pd.Timestamp, d_activation: pd.Timestamp) -> int:
    """Diferencia en meses calendario (M=0, M1=1, M2=2, ...)."""
    if pd.isna(d_when) or pd.isna(d_activation):
        return np.nan
    return (d_when.year - d_activation.year) * 12 + (d_when.month - d_activation.month)

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

def classify_producto(tipo_val, costo_paquete_val):
    tipo = str(tipo_val or "").upper()
    try:
        costo = float(costo_paquete_val)
    except:
        costo = np.nan
    if "MOB" in tipo:
        return "MBB"
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    return "MBB"

def read_sheet_with_header_detection(xls: pd.ExcelFile, sheet_name: str, header_candidates=("DN","MSISDN","NUMERO"), search_rows: int = 50) -> pd.DataFrame:
    """
    Lee una hoja de Excel detectando automáticamente la fila de encabezados.
    Busca en las primeras `search_rows` filas una que contenga alguno(s) de `header_candidates`.
    """
    # 1) Intenta header=0 directo
    df0 = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if not all(str(c).startswith("Unnamed:") for c in df0.columns):
        return df0

    # 2) Carga sin encabezado y detecta fila
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    max_row = min(search_rows, len(raw))
    found_row = None
    for r in range(max_row):
        row_vals = raw.iloc[r].astype(str).str.strip().str.upper().tolist()
        # si esta fila contiene alguno de los candidatos como encabezado, la tomamos
        if any(cand.strip().upper() in row_vals for cand in header_candidates):
            found_row = r
            break

    if found_row is None:
        # fallback: devuelve como estaba (sin encabezados), el caller tendrá que manejarlo
        return pd.read_excel(xls, sheet_name=sheet_name, header=0)

    df = pd.read_excel(xls, sheet_name=sheet_name, header=found_row)
    # Quita columnas completamente vacías
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    return df

def clone_columns_like(template_df: pd.DataFrame, output_df: pd.DataFrame) -> pd.DataFrame:
    tpl_cols = list(template_df.columns)
    result = {}
    for c in tpl_cols:
        if c in output_df.columns:
            result[c] = output_df[c]
        else:
            result[c] = ""
    return pd.DataFrame(result)

# =========================
# Core
# =========================

def calc_report(
    xls_base: pd.ExcelFile,
    xls_tpl: pd.ExcelFile,
    year: int,
    month: int,
    dist_filtro: str
) -> BytesIO:

    month_start = pd.Timestamp(year, month, 1)
    month_end = month_start + pd.offsets.MonthEnd(1)

    # --------- Base mensual ---------
    df_tot = pd.read_excel(xls_base, sheet_name="Desgloce Totales", header=1)
    df_rec = pd.read_excel(xls_base, sheet_name="Desgloce Recarga", header=3)

    # Normalizaciones
    col_fecha_tot = first_existing_col(df_tot, ["FECHA", "FECHA (alta)", "FECHA ALTA", "ALTA"])
    col_dn_tot = first_existing_col(df_tot, ["DN", "MSISDN", "NUMERO"])
    col_dist = first_existing_col(df_tot, ["DISTRIBUIDOR ", "DISTRIBUIDOR"])
    col_plan = first_existing_col(df_tot, ["PLAN"])
    col_tipo = first_existing_col(df_tot, ["TIPO"])
    col_costo = first_existing_col(df_tot, ["COSTO PAQUETE", "COSTO", "PAQUETE"])

    df_tot = df_tot.copy()
    df_tot[col_fecha_tot] = safe_to_datetime(df_tot[col_fecha_tot])
    df_tot["DN_NORM"] = normalize_dn(df_tot[col_dn_tot])
    df_tot["DISTRIBUIDOR_NORM"] = df_tot[col_dist].astype(str).str.strip().str.lower()

    dist_norm = str(dist_filtro or "").strip().lower()
    tot_dist = df_tot[df_tot["DISTRIBUIDOR_NORM"] == dist_norm].copy()
    if tot_dist.empty:
        raise ValueError(f"No hay registros para el distribuidor '{dist_filtro}' en 'Desgloce Totales'.")

    tot_dist["PRODUCTO"] = tot_dist.apply(lambda r: classify_producto(r.get(col_tipo), r.get(col_costo)), axis=1)

    altas_mes = tot_dist[(tot_dist[col_fecha_tot] >= month_start) & (tot_dist[col_fecha_tot] <= month_end)].copy()
    n_altas_mes = int(altas_mes["DN_NORM"].nunique())

    # --------- Recargas ---------
    col_fecha_rec = first_existing_col(df_rec, ["FECHA", "FECHA RECARGA"])
    col_dn_rec = first_existing_col(df_rec, ["DN", "MSISDN", "NUMERO"])
    col_monto_rec = first_existing_col(df_rec, ["MONTO", "IMPORTE", "CARGO"])
    col_plan_rec = None
    for c in ["PLAN", "PAQUETE"]:
        try:
            col_plan_rec = first_existing_col(df_rec, [c])
            break
        except:
            pass

    df_rec = df_rec.copy()
    df_rec[col_fecha_rec] = safe_to_datetime(df_rec[col_fecha_rec])
    df_rec["DN_NORM"] = normalize_dn(df_rec[col_dn_rec])

    rec_month = df_rec[(df_rec[col_fecha_rec] >= month_start) & (df_rec[col_fecha_rec] <= month_end)].copy()
    dns_dist = set(tot_dist["DN_NORM"].dropna())
    rec_month_dist = rec_month[rec_month["DN_NORM"].isin(dns_dist)].copy()

    # --------- PLANTILLA: HISTORIAL DE ACTIVACIONES ---------
    hist_sheet_name = "HISTORIAL DE ACTIVACIONES"
    if hist_sheet_name not in xls_tpl.sheet_names:
        raise ValueError(f"Tu plantilla no tiene la hoja '{hist_sheet_name}'.")

    # DETECCIÓN DE ENCABEZADO AQUÍ 👇
    hist_tpl = read_sheet_with_header_detection(
        xls_tpl,
        hist_sheet_name,
        header_candidates=("DN","DN PORTADO","FECHA (ALTA)","FECHA ALTA","FECHA")
    )

    # Columnas clave en historial (tolerante)
    col_dn_hist = first_existing_col(hist_tpl, ["DN", "MSISDN", "NUMERO"])
    # fecha de alta en historial (si no, la tomamos de tot_dist)
    try:
        col_fecha_hist = first_existing_col(hist_tpl, ["FECHA (alta)", "FECHA ALTA", "FECHA", "ALTA"])
    except:
        col_fecha_hist = None
    # DN PORTADO
    try:
        col_portado_hist = first_existing_col(hist_tpl, ["DN PORTADO", "PORTADO", "PORTABILIDAD"])
    except:
        col_portado_hist = None

    hist_tpl = hist_tpl.copy()
    hist_tpl["DN_NORM"] = normalize_dn(hist_tpl[col_dn_hist])
    if col_fecha_hist:
        hist_tpl[col_fecha_hist] = safe_to_datetime(hist_tpl[col_fecha_hist])

    enrich_df = pd.DataFrame({"DN_NORM": hist_tpl["DN_NORM"]})
    if col_fecha_hist:
        enrich_df["FECHA_ALTA_HIST"] = hist_tpl[col_fecha_hist].values
    if col_portado_hist:
        enrich_df["ES_PORTADO"] = hist_tpl[col_portado_hist].astype(str).str.strip().ne("")

    tot_dist = tot_dist.merge(enrich_df, on="DN_NORM", how="left")
    tot_dist["FECHA_ALTA_BASE"] = tot_dist["FECHA_ALTA_HIST"]
    tot_dist.loc[tot_dist["FECHA_ALTA_BASE"].isna(), "FECHA_ALTA_BASE"] = tot_dist[col_fecha_tot]
    if "ES_PORTADO" not in tot_dist.columns:
        tot_dist["ES_PORTADO"] = False

    # ------- Mes índice para recargas (+M2) -------
    rec_month_dist = rec_month_dist.merge(
        tot_dist[["DN_NORM", "FECHA_ALTA_BASE", col_plan, col_tipo, col_costo, "PRODUCTO"]],
        on="DN_NORM",
        how="left"
    )
    rec_month_dist["MES_INDEX"] = rec_month_dist.apply(
        lambda r: month_diff(pd.Timestamp(year, month, 1), r["FECHA_ALTA_BASE"]), axis=1
    )
    rec_month_dist["ES_M2"] = rec_month_dist["MES_INDEX"].eq(2)

    # ------- 1ra recarga de la vida (en el mes) -------
    first_rec = df_rec.sort_values(col_fecha_rec).groupby("DN_NORM", as_index=False)[col_fecha_rec].min()
    first_rec = first_rec.rename(columns={col_fecha_rec: "FIRST_RECARGA"})
    rec_month_dist = rec_month_dist.merge(first_rec, on="DN_NORM", how="left")
    rec_month_dist["ES_PRIMERA_RECARGA_DEL_MES"] = rec_month_dist["FIRST_RECARGA"].between(month_start, month_end, inclusive="both")

    # ------- Portabilidad (altas portadas del mes) $30 -------
    altas_mes = altas_mes.merge(tot_dist[["DN_NORM", "ES_PORTADO"]], on="DN_NORM", how="left")
    n_portadas_mes = int(altas_mes["ES_PORTADO"].fillna(False).sum())
    comision_porta_total = n_portadas_mes * 30.0

    # ------- Cartera -------
    min_mbb, min_mifi, min_hbb = 35, 110, 99
    pct_mbb = cartera_pct_mbb(n_altas_mes)

    rec_by_dn = rec_month_dist.groupby("DN_NORM", as_index=False)[col_monto_rec].sum().rename(columns={col_monto_rec: "RECARGA_TOTAL_MES"})

    anexo = tot_dist[[col_dn_tot, "DN_NORM", col_fecha_tot, col_plan, col_costo, "PRODUCTO"]].merge(rec_by_dn, on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    def elegible(monto, prod):
        if prod == "MBB":
            return monto >= min_mbb
        elif prod == "MiFi":
            return monto >= min_mifi
        elif prod == "HBB":
            return monto >= min_hbb
        return False
    anexo["ELEGIBLE_CARTERA"] = anexo.apply(lambda r: elegible(r["RECARGA_TOTAL_MES"], r["PRODUCTO"]), axis=1)

    def pct_aplicado(prod):
        if prod == "MBB":
            return pct_mbb
        elif prod in ("MiFi", "HBB"):
            return 0.05
        return 0.0
    anexo["% CARTERA APLICADA"] = anexo["PRODUCTO"].apply(pct_aplicado)

    dn_es_m2 = rec_month_dist.groupby("DN_NORM", as_index=False)["ES_M2"].max()
    anexo = anexo.merge(dn_es_m2, on="DN_NORM", how="left").rename(columns={"ES_M2": "ES_M2_MES"})
    anexo["ES_M2_MES"] = anexo["ES_M2_MES"].fillna(False)

    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"] & anexo["ES_M2_MES"],
        (anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"]).round(2),
        0.0
    )

    dn_first_in_month = rec_month_dist.groupby("DN_NORM", as_index=False)["ES_PRIMERA_RECARGA_DEL_MES"].max()
    anexo = anexo.merge(dn_first_in_month, on="DN_NORM", how="left")
    anexo["ES_PRIMERA_RECARGA_DEL_MES"] = anexo["ES_PRIMERA_RECARGA_DEL_MES"].fillna(False)
    anexo["COMISION_1RA_RECARGA_$"] = np.where(anexo["ES_PRIMERA_RECARGA_DEL_MES"], 15.0, 0.0)

    comision_cartera_total = float(anexo["COMISION_CARTERA_$"].sum())
    comision_first_total = float(anexo["COMISION_1RA_RECARGA_$"].sum())
    recargas_totales_mes = float(rec_month_dist[col_monto_rec].sum())
    comision_total_mes = comision_cartera_total + comision_first_total + comision_porta_total

    # ------- HISTORIAL DE ACTIVACIONES (mes) -------
    hist_mes = tot_dist[(tot_dist[col_fecha_tot] >= month_start) & (tot_dist[col_fecha_tot] <= month_end)].copy()
    hist_mes_out = hist_mes[[col_fecha_tot, "DN_NORM", col_plan, col_costo]].rename(columns={
        col_fecha_tot: "FECHA",
        "DN_NORM": "DN",
        col_plan: "PLAN",
        col_costo: "COSTO PAQUETE"
    }).sort_values("FECHA")

    # ------- CARTERA MES (detalle) -------
    rec_det = rec_month_dist.copy()
    cartera_mes_out = rec_det[[col_fecha_rec, "DN_NORM", col_plan if col_plan_rec is None else col_plan_rec, col_monto_rec]].rename(columns={
        col_fecha_rec: "FECHA",
        "DN_NORM": "DN",
        (col_plan if col_plan_rec is None else col_plan_rec): "PLAN",
        col_monto_rec: "MONTO"
    }).sort_values("FECHA")

    # ------- RESUMEN MES (por producto) -------
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
        .agg({
            "DN_NORM": "nunique",
            "RECARGA_TOTAL_MES": "sum",
            "COMISION_CARTERA_$": "sum",
            "COMISION_1RA_RECARGA_$": "sum"
        })
        .rename(columns={
            "DN_NORM": "Lineas",
            "RECARGA_TOTAL_MES": "Recarga_Mes_$",
            "COMISION_CARTERA_$": "Comision_Cartera_$",
            "COMISION_1RA_RECARGA_$": "Comision_1ra_Rec_$"
        })
    )
    if not resumen_mes.empty:
        total_row = pd.DataFrame([{
            "PRODUCTO": "TOTAL",
            "Lineas": resumen_mes["Lineas"].sum(),
            "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
            "Comision_Cartera_$": resumen_mes["Comision_Cartera_$"].sum(),
            "Comision_1ra_Rec_$": resumen_mes["Comision_1ra_Rec_$"].sum()
        }])
        resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # =========================
    # Clonar encabezados de la PLANTILLA
    # =========================
    tpl_resumen = read_sheet_with_header_detection(xls_tpl, "RESUMEN", header_candidates=("DISTRIBUIDOR","MES"))
    resumen_calc = pd.DataFrame([{
        "Distribuidor": dist_filtro,
        "Mes": f"{month_start.strftime('%B').capitalize()} {year}",
        "Altas del mes": n_altas_mes,
        "Recargas totales del mes ($)": round(recargas_totales_mes, 2),
        "Porcentaje Cartera aplicado (MBB)": cartera_pct_mbb(n_altas_mes),
        "Comisión Cartera total ($)": round(comision_cartera_total, 2),
        "Comisión 1ra recarga ($)": round(comision_first_total, 2),
        "Comisión Portabilidad ($)": round(comision_porta_total, 2),
        "Comisión TOTAL del mes ($)": round(comision_total_mes, 2)
    }])
    resumen_out = clone_columns_like(tpl_resumen, resumen_calc)

    tpl_anexo = read_sheet_with_header_detection(xls_tpl, "ANEXO", header_candidates=("DN","PLAN"))
    anexo_out_base = anexo.rename(columns={
        col_fecha_tot: "FECHA ALTA",
        col_plan: "PLAN",
        col_costo: "COSTO PAQUETE",
        col_dn_tot: "DN (raw)"
    })
    anexo_out_base["DN"] = anexo_out_base["DN_NORM"]
    anexo_out_base["ES_M2 (tercer mes)"] = anexo_out_base["ES_M2_MES"]
    anexo_out_base["% CARTERA"] = anexo_out_base["% CARTERA APLICADA"]
    anexo_out_base["RECARGA MES ($)"] = anexo_out_base["RECARGA_TOTAL_MES"]
    anexo_out_base["COMISION MES ($)"] = anexo_out_base["COMISION_CARTERA_$"] + anexo_out_base["COMISION_1RA_RECARGA_$"]
    anexo_out = clone_columns_like(tpl_anexo, anexo_out_base)

    tpl_hist = read_sheet_with_header_detection(xls_tpl, "HISTORIAL DE ACTIVACIONES", header_candidates=("DN","FECHA","PLAN"))
    hist_out = clone_columns_like(tpl_hist, hist_mes_out)

    month_name_up = month_start.strftime("%B").upper()
    sheet_resumen_mes = f"RESUMEN {month_name_up} {year}"
    sheet_cartera_mes = f"CARTERA {month_name_up} {year}"

    if sheet_resumen_mes in xls_tpl.sheet_names:
        tpl_resumen_mes = read_sheet_with_header_detection(xls_tpl, sheet_resumen_mes, header_candidates=("PRODUCTO","LINEAS"))
    else:
        tpl_resumen_mes = tpl_resumen.copy()
    resumen_mes_out = clone_columns_like(tpl_resumen_mes, resumen_mes)

    if sheet_cartera_mes in xls_tpl.sheet_names:
        tpl_cartera_mes = read_sheet_with_header_detection(xls_tpl, sheet_cartera_mes, header_candidates=("FECHA","DN","PLAN","MONTO"))
    else:
        tpl_cartera_mes = pd.DataFrame(columns=["FECHA","DN","PLAN","MONTO"])
    cartera_mes_out = clone_columns_like(tpl_cartera_mes, cartera_mes_out)

    # =========================
    # Exportar
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen_out.to_excel(writer, sheet_name="RESUMEN", index=False)
        anexo_out.to_excel(writer, sheet_name="ANEXO", index=False)
        hist_out.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)
        resumen_mes_out.to_excel(writer, sheet_name=sheet_resumen_mes, index=False)
        cartera_mes_out.to_excel(writer, sheet_name=sheet_cartera_mes, index=False)
    output.seek(0)
    return output

# =========================
# UI
# =========================

col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial... .xlsx)", type=["xlsx"])
    st.caption("Debe traer: 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")
with col2:
    tpl_file = st.file_uploader("Plantilla / Historial del distribuidor (.xlsx)", type=["xlsx"])
    dist = st.text_input("Distribuidor (exacto como aparece en la base)", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.markdown("---")

if base_file and tpl_file and st.button("Generar reporte"):
    try:
        xls_base = pd.ExcelFile(base_file, engine="openpyxl")
        xls_tpl = pd.ExcelFile(tpl_file, engine="openpyxl")
        buf = calc_report(xls_base, xls_tpl, int(year), int(month), dist_filtro=dist.strip())
        fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {year}.xlsx"
        st.success("✅ Reporte generado.")
        st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)

