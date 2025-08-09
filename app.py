# app.py
# --------------------------------------------------------------------------------------
# Valor Telecom - Generador de Comisiones (MVP)
# Lee: (1) Base mensual (VT Reporte Comercial...) y (2) Archivo histórico/plantilla del distribuidor.
# Calcula: Cartera +M2, 1ª recarga ($15), Portabilidad ($30).
# Replica hojas y encabezados del ejemplo: RESUMEN, ANEXO, HISTORIAL DE ACTIVACIONES,
# RESUMEN {MES AÑO}, CARTERA {MES AÑO}.
# --------------------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# ---------- Config UI ----------
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y el archivo histórico del distribuidor. Genera un Excel con RESUMEN, ANEXO, HISTORIAL, RESUMEN {MES} y CARTERA {MES} con los cálculos de comisiones.")

# ---------- Parámetros de negocio (fáciles de ajustar) ----------
# Ventanas por días desde la activación:
# M (0-30), M+1 (31-60), M2 (61-90). (Ajusta si cambian los umbrales)
WINDOWS = {
    "M": (0, 30),
    "M+1": (31, 60),
    "M2": (61, 90),
}

# Mínimos de recarga por producto para cartera
MIN_MBB = 35
MIN_MIFI = 110
MIN_HBB = 99

# Comisión fija por 1ª recarga y por portabilidad
BONO_PRIMERA_RECARGA = 15.0
BONO_PORTABILIDAD = 30.0   # a partir del mes actual

# Clasificación de producto por costo de paquete (fallback si TIPO no trae distinción)
HBB_COSTOS = {99, 115, 349, 399, 439, 500}
MIFI_COSTOS = {110, 120, 160, 245, 375, 480, 620}

# ---------- Utilidades ----------
def normalize_dn_series(series: pd.Series) -> pd.Series:
    """Normaliza el número (DN) para evitar 55e+12, .0, etc."""
    s = series.astype(str).str.replace(r"\.0$", "", regex=True)
    def fix(x):
        try:
            if "e+" in x.lower():
                return str(int(float(x)))
            return x.split(".")[0]
        except:
            return x
    return s.apply(fix).str.strip()

def cartera_pct_mbb(n_altas_mes: int) -> float:
    """
    % cartera MBB según altas del distribuidor en el mes:
      - <50  -> 3%
      - 50–299 -> 5%
      - 300–499 -> 7%
      - 500–999 -> 8%
      - >=1000 -> 10%
    """
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

def classify_producto(row) -> str:
    """Clasifica MBB / MiFi / HBB usando TIPO y/o COSTO PAQUETE."""
    tipo = str(row.get("TIPO", "")).upper()
    costo = row.get("COSTO PAQUETE", np.nan)
    if "MOB" in tipo:             # cuando la base trae MOB para movilidad
        return "MBB"
    if pd.notna(costo):
        try:
            c = float(costo)
        except:
            c = None
        if c is not None:
            if c in HBB_COSTOS:
                return "HBB"
            if c in MIFI_COSTOS:
                return "MiFi"
    # por defecto
    return "MBB"

def days_between(d1: pd.Timestamp, d2: pd.Timestamp) -> int:
    """Diferencia en días (entero) entre dos fechas."""
    if pd.isna(d1) or pd.isna(d2):
        return np.nan
    return (d2.normalize() - d1.normalize()).days

def label_window(days: float) -> str:
    """Etiqueta ventana M / M+1 / M2 según días desde alta."""
    if pd.isna(days):
        return ""
    for k, (lo, hi) in WINDOWS.items():
        if lo <= days <= hi:
            return k
    return ""  # fuera de rango M1–M2

def month_bounds(year: int, month: int):
    start = pd.Timestamp(year, month, 1)
    end = start + pd.offsets.MonthEnd(1)
    return start, end

def try_read_sheet(xls, name, header=None):
    """Lee una hoja si existe; si no, regresa df vacío."""
    try:
        return pd.read_excel(xls, sheet_name=name, header=header, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

def first_existing_col(df: pd.DataFrame, candidates):
    """Devuelve el primer nombre de columna que exista de una lista de candidatos; error si ninguna."""
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(f"No encontré ninguna de las columnas {candidates} en: {list(df.columns)}")

def get_month_name_es(dt: pd.Timestamp) -> str:
    return dt.strftime("%B").capitalize()

# ---------- Cálculo principal ----------
def calc_report(
    xls_base,
    xls_hist_tpl,
    dist_name: str,
    year: int,
    month: int,
) -> BytesIO:
    """Genera el Excel final clonando estructura del archivo ejemplo y calculando comisiones."""

    month_start, month_end = month_bounds(year, month)
    mes_nombre = get_month_name_es(month_start).upper()
    periodo_titulo = f"{mes_nombre} {year}"

    # --- Leer base mensual ---
    df_tot = try_read_sheet(xls_base, "Desgloce Totales", header=1)   # encabezado inicia en fila 2
    df_rec = try_read_sheet(xls_base, "Desgloce Recarga", header=3)   # encabezado inicia en fila 4

    if df_tot.empty or df_rec.empty:
        raise ValueError("No pude leer 'Desgloce Totales' y/o 'Desgloce Recarga' de la base mensual.")

    # Normalizar columnas clave
    # FECHA (altas y recargas)
    if "FECHA" in df_tot.columns:
        df_tot["FECHA"] = pd.to_datetime(df_tot["FECHA"], errors="coerce")
    if "FECHA" in df_rec.columns:
        df_rec["FECHA"] = pd.to_datetime(df_rec["FECHA"], errors="coerce")

    # DN normalizado
    df_tot["DN_NORM"] = normalize_dn_series(df_tot["DN"])
    df_rec["DN_NORM"] = normalize_dn_series(df_rec["DN"])

    # Filtro por distribuidor (en Totales viene el distribuidor)
    col_dist = first_existing_col(df_tot, ["DISTRIBUIDOR ", "DISTRIBUIDOR"])
    mask_dist = df_tot[col_dist].astype(str).str.strip().str.lower() == dist_name.strip().lower()
    tot_dist = df_tot[mask_dist].copy()

    if tot_dist.empty:
        raise ValueError(f"No encontré líneas del distribuidor '{dist_name}' en 'Desgloce Totales'.")

    # Universo de DNs del distribuidor
    dns_dist = set(tot_dist["DN_NORM"].dropna())

    # Recargas del mes de esos DN
    rec_month_all = df_rec[(df_rec["FECHA"] >= month_start) & (df_rec["FECHA"] <= month_end)].copy()
    rec_month_dist = rec_month_all[rec_month_all["DN_NORM"].isin(dns_dist)].copy()

    # Clasificación por producto
    tot_dist["PRODUCTO"] = tot_dist.apply(classify_producto, axis=1)

    # --- Activaciones del mes (para % MBB y para HISTORIAL) ---
    altas_mes = tot_dist[(tot_dist["FECHA"] >= month_start) & (tot_dist["FECHA"] <= month_end)].copy()
    altas_mes_mbb = altas_mes[altas_mes["PRODUCTO"] == "MBB"].copy()
    n_altas_mbb_mes = altas_mes_mbb["DN_NORM"].nunique()
    pct_mbb = cartera_pct_mbb(n_altas_mbb_mes)

    # --- Fecha de alta por DN (primera fecha en la base totales para ese DN) ---
    altas_por_dn = (
        tot_dist.sort_values("FECHA")
        .groupby("DN_NORM", as_index=False)
        .agg(Fecha_Alta=("FECHA", "first"))
    )

    # --- Portabilidad: desde HISTORIAL DE ACTIVACIONES de la plantilla (distribuidor) ---
    hist_tpl = try_read_sheet(xls_hist_tpl, "HISTORIAL DE ACTIVACIONES", header=0)
    # Acomodar columnas esperadas
    # Buscamos DN (columna DN o similar), FECHA (alta) y DN PORTADO
    col_dn_hist = first_existing_col(hist_tpl, ["DN", "MSISDN", "NUMERO"])
    # FECHA alta (tolerar nombres típicos)
    col_fecha_alta_hist = first_existing_col(hist_tpl, ["FECHA (alta)", "FECHA ALTA", "FECHA_ALTA", "FECHA"])
    # DN PORTADO
    col_portado = first_existing_col(hist_tpl, ["DN PORTADO", "PORTABILIDAD", "DN_PORTADO"])

    hist_tpl[col_fecha_alta_hist] = pd.to_datetime(hist_tpl[col_fecha_alta_hist], errors="coerce")
    hist_tpl[col_dn_hist] = normalize_dn_series(hist_tpl[col_dn_hist])

    # Portabilidades activadas en el mes (DN_PORTADO no vacío)
    port_mes = hist_tpl[
        (hist_tpl[col_fecha_alta_hist] >= month_start) &
        (hist_tpl[col_fecha_alta_hist] <= month_end) &
        (hist_tpl[col_portado].astype(str).str.strip() != "")
    ].copy()
    n_port = port_mes[col_dn_hist].nunique()
    com_port = n_port * BONO_PORTABILIDAD

    # --- 1ª recarga: $15 por la primera recarga "en la vida" de la línea que ocurra en el mes ---
    # Para determinar si es la primera, necesitamos historial de recargas previas.
    # 1) Tomamos de la base mensual todas las recargas (histórico que venga)
    rec_all = df_rec.copy()
    # 2) Además, intentamos sumar historial desde las hojas "CARTERA ..." de la plantilla
    #    (por si el archivo del distribuidor trae meses anteriores).
    rec_hist_list = []
    for sn in getattr(xls_hist_tpl, "sheet_names", []):
        up = sn.upper()
        if up.startswith("CARTERA "):
            df_c = try_read_sheet(xls_hist_tpl, sn, header=0)
            if df_c.empty:
                continue
            # Detectar columnas plausibles
            try:
                col_fecha_c = first_existing_col(df_c, ["FECHA", "FECHA RECARGA", "FECHA_RECARGA"])
                col_dn_c = first_existing_col(df_c, ["DN", "MSISDN", "NUMERO"])
                col_monto_c = first_existing_col(df_c, ["MONTO", "IMPORTE"])
            except Exception:
                continue
            df_c = df_c[[col_fecha_c, col_dn_c, col_monto_c]].copy()
            df_c.columns = ["FECHA", "DN", "MONTO"]
            df_c["FECHA"] = pd.to_datetime(df_c["FECHA"], errors="coerce")
            df_c["DN_NORM"] = normalize_dn_series(df_c["DN"])
            rec_hist_list.append(df_c[["FECHA", "DN_NORM", "MONTO"]])

    rec_hist_tpl = pd.concat(rec_hist_list, ignore_index=True) if rec_hist_list else pd.DataFrame(columns=["FECHA", "DN_NORM", "MONTO"])

    # Unificamos para validar "primera recarga"
    # Nos quedamos con DN_NORM y FECHA de todas las recargas conocidas (base + plantilla)
    rec_all_norm = rec_all.copy()
    rec_all_norm["FECHA"] = pd.to_datetime(rec_all_norm["FECHA"], errors="coerce")
    rec_all_norm["DN_NORM"] = normalize_dn_series(rec_all_norm["DN"])
    rec_union = pd.concat([
        rec_all_norm[["FECHA", "DN_NORM"]],
        rec_hist_tpl[["FECHA", "DN_NORM"]]
    ], ignore_index=True).dropna()

    # Para cada DN, su primera recarga en la vida
    first_recharge = (
        rec_union.sort_values("FECHA")
        .groupby("DN_NORM", as_index=False)
        .agg(Fecha_Primera_Recarga=("FECHA", "first"))
    )

    # Ver cuáles de esas primeras recargas cayeron en el mes objetivo **y** pertenecen al universo del distribuidor
    first_rec_mes = first_recharge[
        (first_recharge["Fecha_Primera_Recarga"] >= month_start) &
        (first_recharge["Fecha_Primera_Recarga"] <= month_end) &
        (first_recharge["DN_NORM"].isin(dns_dist))
    ]
    n_first_rec = first_rec_mes["DN_NORM"].nunique()
    com_primera = n_first_rec * BONO_PRIMERA_RECARGA

    # --- Base Cartera por producto ---
    # Enlazar recargas del mes con fecha de alta para calcular días y ventana
    rec_month_dist = rec_month_dist.merge(
        altas_por_dn, on="DN_NORM", how="left"
    )

    rec_month_dist["DIAS_DESDE_ALTA"] = rec_month_dist.apply(
        lambda r: days_between(r.get("Fecha_Alta"), r.get("FECHA")), axis=1
    )
    rec_month_dist["WINDOW"] = rec_month_dist["DIAS_DESDE_ALTA"].apply(label_window)

    # Producto por DN (desde tot_dist)
    prod_por_dn = tot_dist[["DN_NORM", "PRODUCTO"]].drop_duplicates()
    rec_month_dist = rec_month_dist.merge(prod_por_dn, on="DN_NORM", how="left")

    # Montos por DN en el mes
    rec_sum_por_dn = rec_month_dist.groupby(["DN_NORM", "PRODUCTO"], as_index=False)["MONTO"].sum().rename(columns={"MONTO": "RECARGA_TOTAL_MES"})

    # Elegibilidad por producto (mínimos)
    def elegible_cartera(prod: str, rec_total: float) -> bool:
        if prod == "MBB":
            return rec_total >= MIN_MBB
        elif prod == "MiFi":
            return rec_total >= MIN_MIFI
        elif prod == "HBB":
            return rec_total >= MIN_HBB
        return False

    rec_sum_por_dn["ELEGIBLE_MINIMO"] = rec_sum_por_dn.apply(lambda r: elegible_cartera(r["PRODUCTO"], r["RECARGA_TOTAL_MES"]), axis=1)

    # --- Comisión Cartera ---
    # MBB: aplicar % del distribuidor **sólo a líneas en ventana M2** (tercer mes) durante el mes.
    # Para identificar si el DN estuvo en M2 en el mes, basta con que tenga alguna recarga marcada WINDOW=="M2".
    dnis_mbb_m2 = set(
        rec_month_dist[(rec_month_dist["PRODUCTO"] == "MBB") & (rec_month_dist["WINDOW"] == "M2")]["DN_NORM"].unique()
    )
    base_mbb = rec_sum_por_dn[
        (rec_sum_por_dn["PRODUCTO"] == "MBB") &
        (rec_sum_por_dn["DN_NORM"].isin(dnis_mbb_m2)) &
        (rec_sum_por_dn["ELEGIBLE_MINIMO"])
    ]["RECARGA_TOTAL_MES"].sum()

    com_cartera_mbb = base_mbb * pct_mbb

    # MiFi/HBB: 5% M1–M12 (en este MVP tomamos las recargas del mes si cumplen mínimo)
    base_mifi = rec_sum_por_dn[(rec_sum_por_dn["PRODUCTO"] == "MiFi") & (rec_sum_por_dn["ELEGIBLE_MINIMO"])]["RECARGA_TOTAL_MES"].sum()
    base_hbb  = rec_sum_por_dn[(rec_sum_por_dn["PRODUCTO"] == "HBB") & (rec_sum_por_dn["ELEGIBLE_MINIMO"])]["RECARGA_TOTAL_MES"].sum()
    com_cartera_mifi = base_mifi * 0.05
    com_cartera_hbb  = base_hbb  * 0.05

    com_cartera_total = round(com_cartera_mbb + com_cartera_mifi + com_cartera_hbb, 2)

    # --- ANEXO (detalle por DN con % aplicado y comisión) ---
    # Para ANEXO usamos una fila por DN, con: FECHA alta, PLAN, COSTO PAQUETE, PRODUCTO, recarga del mes, elegible, % aplicado, comisión
    # Traemos PLAN y COSTO PAQUETE desde tot_dist (el último o el primero conocido para el DN)
    plan_por_dn = (
        tot_dist.sort_values("FECHA")
        .groupby("DN_NORM", as_index=False)
        .agg(
            PLAN=("PLAN", "last"),
            COSTO=("COSTO PAQUETE", "last"),
            PRODUCTO=("PRODUCTO","last"),
            FECHA_ALTA=("FECHA","first")
        )
    )
    anexo = plan_por_dn.merge(rec_sum_por_dn[["DN_NORM","RECARGA_TOTAL_MES"]], on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    def pct_aplicado_row(prod, dn):
        if prod == "MBB":
            return pct_mbb if dn in dnis_mbb_m2 else 0.0
        if prod in ("MiFi", "HBB"):
            # aplica 5% si cumple mínimo de su producto
            rec_total = anexo.loc[anexo["DN_NORM"]==dn, "RECARGA_TOTAL_MES"].values[0]
            if (prod == "MiFi" and rec_total >= MIN_MIFI) or (prod == "HBB" and rec_total >= MIN_HBB):
                return 0.05
            return 0.0
        return 0.0

    anexo["% CARTERA APLICADA"] = anexo.apply(lambda r: pct_aplicado_row(r["PRODUCTO"], r["DN_NORM"]), axis=1)
    anexo["COMISION_CARTERA_$"] = (anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"]).round(2)

    # --- HISTORIAL DE ACTIVACIONES (del mes)
    hist_out = (
        tot_dist[(tot_dist["FECHA"] >= month_start) & (tot_dist["FECHA"] <= month_end)]
        .copy()
        .assign(DN=lambda d: d["DN_NORM"])
        .sort_values("FECHA")
    )
    # Encabezados típicos del ejemplo
    hist_out = hist_out.rename(columns={
        "FECHA": "FECHA (alta)",
        "COSTO PAQUETE": "COSTO PAQUETE"
    })
    hist_cols = ["FECHA (alta)", "DN", "PLAN", "COSTO PAQUETE"]
    hist_out = hist_out.reindex(columns=hist_cols, fill_value="")

    # --- CARTERA {MES} (detalle recargas del mes)
    rec_det = rec_month_dist.copy()
    # Traer PLAN de tot_dist si no viene en recargas
    if "PLAN" not in rec_det.columns or rec_det["PLAN"].isna().all():
        rec_det = rec_det.merge(
            tot_dist[["DN_NORM","PLAN"]].drop_duplicates(), on="DN_NORM", how="left"
        )
    rec_det_out = rec_det[["FECHA","DN_NORM","PLAN","MONTO"]].rename(columns={"DN_NORM":"DN"})
    rec_det_out = rec_det_out.sort_values(["FECHA","DN"])

    # --- RESUMEN {MES} (por producto)
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
        .agg(
            Lineas=("DN_NORM","nunique"),
            Recarga_Mes_$=("RECARGA_TOTAL_MES","sum"),
            Comision_Mes_$=("COMISION_CARTERA_$","sum"),
        )
    )
    total_row = pd.DataFrame([{
        "PRODUCTO": "TOTAL",
        "Lineas": resumen_mes["Lineas"].sum(),
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
        "Comision_Mes_$": resumen_mes["Comision_Mes_$"].sum(),
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # --- RESUMEN (portada): sumar Cartera + 1ª recarga + Portabilidad
    resumen_portada = pd.DataFrame([{
        "Distribuidor": dist_name,
        "Mes": f"{mes_nombre.capitalize()} {year}",
    }])

    # Columnas numéricas del resumen portada siguiendo el ejemplo
    # (si tu ejemplo tiene nombres distintos, los puedes ajustar aquí)
    resumen_portada["Comisión Cartera ($)"] = round(com_cartera_total, 2)
    resumen_portada["Comisión 1ra recarga ($)"] = round(com_primera, 2)
    resumen_portada["Comisión Portabilidad ($)"] = round(com_port, 2)
    resumen_portada["Total a pagar ($)"] = round(
        resumen_portada["Comisión Cartera ($)"] +
        resumen_portada["Comisión 1ra recarga ($)"] +
        resumen_portada["Comisión Portabilidad ($)"], 2
    )

    # ---------- Exportar Excel (clonando estructura y encabezados) ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # RESUMEN (portada)
        resumen_portada.to_excel(writer, sheet_name="RESUMEN", index=False)

        # ANEXO (detalle por DN)
        # Alinear encabezados comunes del ejemplo: DN, FECHA, PLAN, COSTO PAQUETE, PRODUCTO, RECARGA_TOTAL_MES, %..., COMISION...
        anexo_out = anexo.copy()
        anexo_out = anexo_out.rename(columns={
            "FECHA_ALTA": "FECHA",
            "COSTO": "COSTO PAQUETE",
        })
        cols_anexo = ["DN_NORM","FECHA","PLAN","COSTO PAQUETE","PRODUCTO","RECARGA_TOTAL_MES","% CARTERA APLICADA","COMISION_CARTERA_$"]
        anexo_out = anexo_out.reindex(columns=cols_anexo)
        anexo_out = anexo_out.rename(columns={"DN_NORM":"DN"})
        anexo_out.to_excel(writer, sheet_name="ANEXO", index=False)

        # HISTORIAL DE ACTIVACIONES
        hist_out.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)

        # RESUMEN {MES AÑO}
        sheet_resumen_mes = f"RESUMEN {periodo_titulo}"
        resumen_mes.to_excel(writer, sheet_name=sheet_resumen_mes, index=False)

        # CARTERA {MES AÑO}
        sheet_cartera_mes = f"CARTERA {periodo_titulo}"
        rec_det_out.to_excel(writer, sheet_name=sheet_cartera_mes, index=False)

    output.seek(0)
    return output

# ---------- UI ----------
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial...)", type=["xlsx"])
    st.caption("Debe contener: 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")

with col2:
    hist_file = st.file_uploader("Histórico/plantilla del distribuidor (Ejemplo)", type=["xlsx"])
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.markdown("---")
if base_file and hist_file and st.button("Generar reporte"):
    try:
        xls_base = pd.ExcelFile(base_file, engine="openpyxl")
        xls_hist_tpl = pd.ExcelFile(hist_file, engine="openpyxl")

        buf = calc_report(
            xls_base=xls_base,
            xls_hist_tpl=xls_hist_tpl,
            dist_name=dist.strip(),
            year=int(year),
            month=int(month),
        )

        fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {int(year)}.xlsx"
        st.success("✅ Reporte generado con comisiones.")
        st.download_button(
            "⬇️ Descargar Excel",
            data=buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error("Ocurrió un error al generar el reporte.")
        st.exception(e)
else:
    st.info("Sube **ambos archivos** y da clic en **Generar reporte**.")
