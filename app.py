import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y la plantilla/historial. Genera un Excel clonando los encabezados de tu plantilla y calculando comisiones (Cartera +M2, 1ra Recarga, Portabilidad).")

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
    cols = {c.strip().upper(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().upper()
        if key in cols:
            return cols[key]
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

    # Heurística acordada
    if "MOB" in tipo:
        return "MBB"
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    return "MBB"

def clone_columns_like(template_df: pd.DataFrame, output_df: pd.DataFrame) -> pd.DataFrame:
    """Reordena/renombra columnas del output para que coincidan (en nombre y orden) con template_df."""
    tpl_cols = list(template_df.columns)
    # Intentamos mapear por nombre exacto; si una columna del template no existe en output, la creamos vacía
    result = {}
    for c in tpl_cols:
        if c in output_df.columns:
            result[c] = output_df[c]
        else:
            # Si no existe, la creamos como vacío
            result[c] = ""
    return pd.DataFrame(result)

def calc_report(
    xls_base: pd.ExcelFile,
    xls_tpl: pd.ExcelFile,
    year: int,
    month: int,
    dist_filtro: str
) -> BytesIO:

    month_start = pd.Timestamp(year, month, 1)
    month_end = month_start + pd.offsets.MonthEnd(1)

    # --------- Leer base mensual ---------
    # Hojas y encabezados según acordado
    df_tot = pd.read_excel(xls_base, sheet_name="Desgloce Totales", header=1)
    df_rec = pd.read_excel(xls_base, sheet_name="Desgloce Recarga", header=3)

    # Normalizaciones
    # Columnas críticas en totales
    col_fecha_tot = first_existing_col(df_tot, ["FECHA", "FECHA (alta)", "FECHA ALTA", "ALTA"])
    col_dn_tot = first_existing_col(df_tot, ["DN", "MSISDN", "NUMERO"])
    col_dist = first_existing_col(df_tot, ["DISTRIBUIDOR ", "DISTRIBUIDOR", "DISTRIBUIDOR"])
    col_plan = first_existing_col(df_tot, ["PLAN"])
    col_tipo = first_existing_col(df_tot, ["TIPO"])
    col_costo = first_existing_col(df_tot, ["COSTO PAQUETE", "COSTO", "PAQUETE"])

    df_tot = df_tot.copy()
    df_tot[col_fecha_tot] = safe_to_datetime(df_tot[col_fecha_tot])
    df_tot["DN_NORM"] = normalize_dn(df_tot[col_dn_tot])
    df_tot["DISTRIBUIDOR_NORM"] = df_tot[col_dist].astype(str).str.strip().str.lower()

    # Filtro por distribuidor
    dist_norm = str(dist_filtro or "").strip().lower()
    tot_dist = df_tot[df_tot["DISTRIBUIDOR_NORM"] == dist_norm].copy()
    if tot_dist.empty:
        raise ValueError(f"No hay registros para el distribuidor '{dist_filtro}' en 'Desgloce Totales'.")

    # Clasificación de producto
    tot_dist["PRODUCTO"] = tot_dist.apply(lambda r: classify_producto(r.get(col_tipo), r.get(col_costo)), axis=1)

    # Activaciones del mes (para métricas del mes y portabilidad)
    altas_mes = tot_dist[(tot_dist[col_fecha_tot] >= month_start) & (tot_dist[col_fecha_tot] <= month_end)].copy()
    n_altas_mes = int(altas_mes["DN_NORM"].nunique())

    # --------- Recargas (base) ---------
    col_fecha_rec = first_existing_col(df_rec, ["FECHA", "FECHA RECARGA"])
    col_dn_rec = first_existing_col(df_rec, ["DN", "MSISDN", "NUMERO"])
    col_monto_rec = first_existing_col(df_rec, ["MONTO", "IMPORTE", "CARGO"])
    # plan en recarga puede venir o no
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

    # Recargas del mes
    rec_month = df_rec[(df_rec[col_fecha_rec] >= month_start) & (df_rec[col_fecha_rec] <= month_end)].copy()

    # Limitar universo de recargas del mes a líneas del distribuidor
    dns_dist = set(tot_dist["DN_NORM"].dropna())
    rec_month_dist = rec_month[rec_month["DN_NORM"].isin(dns_dist)].copy()

    # --------- HISTORIAL (desde tu plantilla) para Portabilidad y fechas de alta “oficiales” ---------
    # Buscamos hoja exacta "HISTORIAL DE ACTIVACIONES" (nombre de tu plantilla)
    hist_sheet_name = "HISTORIAL DE ACTIVACIONES"
    if hist_sheet_name not in xls_tpl.sheet_names:
        raise ValueError(f"Tu plantilla no tiene la hoja '{hist_sheet_name}'.")

    hist_tpl = pd.read_excel(xls_tpl, sheet_name=hist_sheet_name, header=0)
    if hist_tpl.empty:
        # si viene vacía, seguimos sin portabilidad
        hist_tpl = pd.DataFrame(columns=["DN", "DN PORTADO", "FECHA (alta)"])

    # Detectar columnas clave con tolerancia de nombres
    col_dn_hist = first_existing_col(hist_tpl, ["DN", "MSISDN", "NUMERO"])
    # fecha de alta en historial
    try:
        col_fecha_hist = first_existing_col(hist_tpl, ["FECHA (alta)", "FECHA ALTA", "FECHA", "ALTA"])
    except:
        # si no existe, usamos la alta de tot_dist
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

    # Merge “fecha alta oficial” y “portado” a universo del distribuidor
    # Si historial no trae fecha/portado, derivamos lo que podamos
    key_merge = ["DN_NORM"]
    enrich_cols = {}
    if col_fecha_hist:
        enrich_cols["FECHA_ALTA_HIST"] = hist_tpl[col_fecha_hist]
    if col_portado_hist:
        # Convertimos a flag: no vacío => True
        enrich_cols["ES_PORTADO"] = hist_tpl[col_portado_hist].astype(str).str.strip().ne("")

    enrich_df = pd.DataFrame({"DN_NORM": hist_tpl["DN_NORM"]})
    for k, series in enrich_cols.items():
        enrich_df[k] = series.values

    tot_dist = tot_dist.merge(enrich_df, on="DN_NORM", how="left")
    # Definimos FECHA_ALTA base a usar
    tot_dist["FECHA_ALTA_BASE"] = tot_dist["FECHA_ALTA_HIST"]
    tot_dist.loc[tot_dist["FECHA_ALTA_BASE"].isna(), "FECHA_ALTA_BASE"] = tot_dist[col_fecha_tot]

    # ------- Identificar M2 (tercer mes) por recarga -------
    # Para cada recarga del mes, buscamos la fecha de alta de esa línea (FECHA_ALTA_BASE) en tot_dist
    rec_month_dist = rec_month_dist.merge(
        tot_dist[["DN_NORM", "FECHA_ALTA_BASE", col_plan, col_tipo, col_costo, "PRODUCTO"]],
        on="DN_NORM",
        how="left"
    )

    # Mes index para cada recarga
    rec_month_dist["MES_INDEX"] = rec_month_dist.apply(
        lambda r: month_diff(pd.Timestamp(year, month, 1), r["FECHA_ALTA_BASE"]), axis=1
    )
    # +M2 == 2
    rec_month_dist["ES_M2"] = rec_month_dist["MES_INDEX"].eq(2)

    # ------- 1ra recarga (en toda la vida) dentro del mes -------
    # Tomamos la primera fecha de recarga de cada DN en todo el historial df_rec
    first_rec = df_rec.sort_values(col_fecha_rec).groupby("DN_NORM", as_index=False)[col_fecha_rec].min()
    first_rec = first_rec.rename(columns={col_fecha_rec: "FIRST_RECARGA"})
    rec_month_dist = rec_month_dist.merge(first_rec, on="DN_NORM", how="left")
    rec_month_dist["ES_PRIMERA_RECARGA_DEL_MES"] = rec_month_dist["FIRST_RECARGA"].between(month_start, month_end, inclusive="both")

    # ------- Portabilidad (pago por alta portada en el mes) -------
    # Regla: $30 por alta portada en el mes de cálculo (alta dentro del mes y ES_PORTADO true)
    # Tomamos altas del mes (altas_mes) y miramos ES_PORTADO
    if "ES_PORTADO" not in tot_dist.columns:
        tot_dist["ES_PORTADO"] = False
    altas_mes = altas_mes.merge(tot_dist[["DN_NORM", "ES_PORTADO"]], on="DN_NORM", how="left")
    n_portadas_mes = int(altas_mes["ES_PORTADO"].fillna(False).sum())
    comision_porta_total = n_portadas_mes * 30.0

    # ------- Cálculo de cartera por producto -------
    # Mínimos
    min_mbb = 35
    min_mifi = 110
    min_hbb = 99

    # % cartera MBB por volumen de altas del mes (del distribuidor)
    pct_mbb = cartera_pct_mbb(n_altas_mes)

    # Suma total de recargas del mes por DN
    rec_by_dn = rec_month_dist.groupby("DN_NORM", as_index=False)[col_monto_rec].sum().rename(columns={col_monto_rec: "RECARGA_TOTAL_MES"})

    # ANEXO base
    anexo = tot_dist[[col_dn_tot, "DN_NORM", col_fecha_tot, col_plan, col_costo, "PRODUCTO"]].merge(rec_by_dn, on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # Elegibilidad por producto (para cartera)
    def elegible(monto, prod):
        if prod == "MBB":
            return monto >= min_mbb
        elif prod == "MiFi":
            return monto >= min_mifi
        elif prod == "HBB":
            return monto >= min_hbb
        return False

    anexo["ELEGIBLE_CARTERA"] = anexo.apply(lambda r: elegible(r["RECARGA_TOTAL_MES"], r["PRODUCTO"]), axis=1)

    # % aplicado por producto
    def pct_aplicado(prod):
        if prod == "MBB":
            return pct_mbb
        elif prod in ("MiFi", "HBB"):
            return 0.05
        return 0.0

    anexo["% CARTERA APLICADA"] = anexo["PRODUCTO"].apply(pct_aplicado)

    # Base +M2: de las recargas del mes, solo las que están en M2 (tercer mes)
    # Para calcular comisiones correctas, unimos la info M2 por DN proporcionalmente:
    # Nota: si quieres prorrateo por múltiples recargas/mes, aquí ya sumamos por DN.
    dn_es_m2 = rec_month_dist.groupby("DN_NORM", as_index=False)["ES_M2"].max()  # si alguna recarga del DN está en M2 => True
    anexo = anexo.merge(dn_es_m2, on="DN_NORM", how="left").rename(columns={"ES_M2": "ES_M2_MES"})
    anexo["ES_M2_MES"] = anexo["ES_M2_MES"].fillna(False)

    # Comisión Cartera (solo base +M2)
    # Interpretación: si ES_M2_MES = True y ELEGIBLE_CARTERA = True, aplicar % sobre RECARGA_TOTAL_MES
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"] & anexo["ES_M2_MES"],
        (anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"]).round(2),
        0.0
    )

    # Comisión 1ra Recarga ($15 por línea cuya primera recarga cae en el mes)
    dn_first_in_month = rec_month_dist.groupby("DN_NORM", as_index=False)["ES_PRIMERA_RECARGA_DEL_MES"].max()
    anexo = anexo.merge(dn_first_in_month, on="DN_NORM", how="left")
    anexo["ES_PRIMERA_RECARGA_DEL_MES"] = anexo["ES_PRIMERA_RECARGA_DEL_MES"].fillna(False)
    anexo["COMISION_1RA_RECARGA_$"] = np.where(anexo["ES_PRIMERA_RECARGA_DEL_MES"], 15.0, 0.0)

    # Totales para RESUMEN
    comision_cartera_total = float(anexo["COMISION_CARTERA_$"].sum())
    comision_first_total = float(anexo["COMISION_1RA_RECARGA_$"].sum())
    comision_total_mes = comision_cartera_total + comision_first_total + comision_porta_total

    recargas_totales_mes = float(rec_month_dist[col_monto_rec].sum())

    # ------- HISTORIAL DE ACTIVACIONES (solo las del mes para la hoja) -------
    hist_mes = tot_dist[(tot_dist[col_fecha_tot] >= month_start) & (tot_dist[col_fecha_tot] <= month_end)].copy()
    hist_mes_out = hist_mes[[col_fecha_tot, "DN_NORM", col_plan, col_costo]].rename(columns={
        col_fecha_tot: "FECHA",
        "DN_NORM": "DN",
        col_plan: "PLAN",
        col_costo: "COSTO PAQUETE"
    }).sort_values("FECHA")

    # ------- CARTERA MES (detalle de recargas del mes) -------
    rec_det = rec_month_dist.copy()
    rec_det["ELEGIBLE_MBB"] = rec_det[col_monto_rec] >= min_mbb
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
    # RESUMEN
    tpl_resumen = pd.read_excel(xls_tpl, sheet_name="RESUMEN", header=0)
    # Construimos un dict con claves esperadas (ajusta campos a tu plantilla exacta)
    resumen_calc = pd.DataFrame([{
        # Llaves típicas (puede que tu plantilla tenga nombres exactos distintos; como clonamos, rellenaremos vacíos)
        "Distribuidor": dist_filtro,
        "Mes": f"{month_start.strftime('%B').capitalize()} {year}",
        "Altas del mes": n_altas_mes,
        "Recargas totales del mes ($)": round(recargas_totales_mes, 2),
        "Porcentaje Cartera aplicado (MBB)": pct_mbb,
        "Comisión Cartera total ($)": round(comision_cartera_total, 2),
        "Comisión 1ra recarga ($)": round(comision_first_total, 2),
        "Comisión Portabilidad ($)": round(comision_porta_total, 2),
        "Comisión TOTAL del mes ($)": round(comision_total_mes, 2)
    }])

    resumen_out = clone_columns_like(tpl_resumen, resumen_calc)

    # ANEXO
    tpl_anexo = pd.read_excel(xls_tpl, sheet_name="ANEXO", header=0)
    # Intento de mapeo a nombres comunes de tu anexo. Si faltan columnas, se rellenan vacías.
    anexo_out_base = anexo.rename(columns={
        col_fecha_tot: "FECHA ALTA",
        col_plan: "PLAN",
        col_costo: "COSTO PAQUETE",
        col_dn_tot: "DN (raw)"
    })
    # Campos útiles adicionales
    anexo_out_base["DN"] = anexo_out_base["DN_NORM"]
    anexo_out_base["ES_M2 (tercer mes)"] = anexo_out_base["ES_M2_MES"]
    anexo_out_base["% CARTERA"] = anexo_out_base["% CARTERA APLICADA"]
    anexo_out_base["RECARGA MES ($)"] = anexo_out_base["RECARGA_TOTAL_MES"]
    anexo_out_base["COMISION MES ($)"] = anexo_out_base["COMISION_CARTERA_$"] + anexo_out_base["COMISION_1RA_RECARGA_$"]

    anexo_out = clone_columns_like(tpl_anexo, anexo_out_base)

    # HISTORIAL DE ACTIVACIONES (clon de encabezados)
    tpl_hist = pd.read_excel(xls_tpl, sheet_name="HISTORIAL DE ACTIVACIONES", header=0)
    hist_out = clone_columns_like(tpl_hist, hist_mes_out)

    # RESUMEN {MES}
    month_name_up = month_start.strftime("%B").upper()
    sheet_resumen_mes = f"RESUMEN {month_name_up} {year}"
    if sheet_resumen_mes in xls_tpl.sheet_names:
        tpl_resumen_mes = pd.read_excel(xls_tpl, sheet_name=sheet_resumen_mes, header=0)
    else:
        # Si tu plantilla trae otro nombre, tomamos las columnas del RESUMEN como fallback
        tpl_resumen_mes = tpl_resumen.copy()

    resumen_mes_out = clone_columns_like(tpl_resumen_mes, resumen_mes)

    # CARTERA {MES}
    sheet_cartera_mes = f"CARTERA {month_name_up} {year}"
    if sheet_cartera_mes in xls_tpl.sheet_names:
        tpl_cartera_mes = pd.read_excel(xls_tpl, sheet_name=sheet_cartera_mes, header=0)
    else:
        tpl_cartera_mes = pd.DataFrame(columns=["FECHA", "DN", "PLAN", "MONTO"])
    cartera_mes_out = clone_columns_like(tpl_cartera_mes, cartera_mes_out)

    # =========================
    # Exportar a Excel (XLSX)
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
        st.success("✅ Reporte generado con encabezados clonados de tu plantilla.")
        st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
