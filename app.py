import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import unicodedata

st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("MVP • Sube la base mensual y el historial del distribuidor. Exporta un Excel con RESUMEN, ANEXO, HISTORIAL (mes), RESUMEN MES y CARTERA MES.")

# ========== Utilidades de normalización ==========
def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        _strip_accents(str(c)).strip().upper().replace("\n", " ").replace("  ", " ")
        for c in df.columns
    ]
    return df

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = list(df.columns)
    for cand in candidates:
        uc = _strip_accents(cand).strip().upper()
        for c in cols:
            if _strip_accents(str(c)).strip().upper() == uc:
                return c
    return None

def normalize_dn(series: pd.Series) -> pd.Series:
    out = series.astype(str).str.replace(r"\.0$", "", regex=True)
    def fix(x):
        try:
            if "e+" in x.lower():
                return str(int(float(x)))
            return x.split(".")[0]
        except Exception:
            return x
    return out.apply(fix)

# ========== Clasificación y reglas ==========
def classify_row(row):
    tipo = str(row.get("TIPO","")).upper()
    costo = row.get("COSTO PAQUETE", np.nan)
    if "MOB" in tipo:
        return "MBB"
    # HBB por costos típicos
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    # MiFi por costos típicos
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    return "MBB"

def cartera_pct_mbb(n_altas_mes):
    if n_altas_mes < 50:   return 0.03
    if n_altas_mes < 300:  return 0.05
    if n_altas_mes < 500:  return 0.07
    if n_altas_mes < 1000: return 0.08
    return 0.10

# ========== Lectura robusta de hojas ==========
def read_with_possible_header(xls, sheet_name, header_try=(0,1,2,3,4)):
    for h in header_try:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=h, engine="openpyxl")
            if df is not None and df.shape[1] > 0:
                return df
        except Exception:
            continue
    # último intento con header por defecto
    return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")

def load_base(base_file):
    xls = pd.ExcelFile(base_file, engine="openpyxl")
    # nombres fijos por tu confirmación
    if "Desgloce Totales" not in xls.sheet_names or "Desgloce Recarga" not in xls.sheet_names:
        raise ValueError("El archivo base debe tener hojas 'Desgloce Totales' y 'Desgloce Recarga'.")
    df_tot = pd.read_excel(xls, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
    df_rec = pd.read_excel(xls, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")

    df_tot = norm_cols(df_tot)
    df_rec = norm_cols(df_rec)

    # Renombres estándar esperados
    # DN
    dn_tot = pick_col(df_tot, ["DN","NUMERO","LINEA","MSISDN"])
    if dn_tot and dn_tot != "DN": df_tot.rename(columns={dn_tot:"DN"}, inplace=True)
    dn_rec = pick_col(df_rec, ["DN","NUMERO","LINEA","MSISDN"])
    if dn_rec and dn_rec != "DN": df_rec.rename(columns={dn_rec:"DN"}, inplace=True)

    # FECHA (activación en totales)
    fecha_tot = pick_col(df_tot, ["FECHA","FECHA ALTA","FECHA ACTIVACION","ALTA"])
    if fecha_tot and fecha_tot != "FECHA": df_tot.rename(columns={fecha_tot:"FECHA"}, inplace=True)

    # FECHA (recarga en recargas)
    fecha_rec = pick_col(df_rec, ["FECHA","FECHA RECARGA","FECHA DE RECARGA"])
    if fecha_rec and fecha_rec != "FECHA": df_rec.rename(columns={fecha_rec:"FECHA"}, inplace=True)

    # MONTO
    monto_rec = pick_col(df_rec, ["MONTO","IMPORTE","MONTO RECARGA","CANTIDAD"])
    if monto_rec and monto_rec != "MONTO": df_rec.rename(columns={monto_rec:"MONTO"}, inplace=True)

    # PLAN / COSTO PAQUETE / FORMA DE PAGO / DISTRIBUIDOR
    plan_tot = pick_col(df_tot, ["PLAN","PLAN TARIFARIO"])
    if plan_tot and plan_tot != "PLAN": df_tot.rename(columns={plan_tot:"PLAN"}, inplace=True)
    costo_tot = pick_col(df_tot, ["COSTO PAQUETE","COSTO DEL PAQUETE","PAQUETE"])
    if costo_tot and costo_tot != "COSTO PAQUETE": df_tot.rename(columns={costo_tot:"COSTO PAQUETE"}, inplace=True)
    tipo_tot = pick_col(df_tot, ["TIPO","TIPO PRODUCTO","PRODUCTO"])
    if tipo_tot and tipo_tot != "TIPO": df_tot.rename(columns={tipo_tot:"TIPO"}, inplace=True)
    dist_tot = pick_col(df_tot, ["DISTRIBUIDOR ","DISTRIBUIDOR"])  # hay versión con espacio al final
    if dist_tot and dist_tot != "DISTRIBUIDOR ":
        df_tot.rename(columns={dist_tot:"DISTRIBUIDOR "}, inplace=True)

    forma_rec = pick_col(df_rec, ["FORMA DE PAGO","METODO DE PAGO","PAGO"])
    if forma_rec and forma_rec != "FORMA DE PAGO": df_rec.rename(columns={forma_rec:"FORMA DE PAGO"}, inplace=True)

    # Normalizaciones finales
    for df in (df_tot, df_rec):
        if "DN" in df.columns:
            df["DN_NORM"] = normalize_dn(df["DN"])
    if "FECHA" in df_tot.columns:
        df_tot["FECHA"] = pd.to_datetime(df_tot["FECHA"], errors="coerce")
    if "FECHA" in df_rec.columns:
        df_rec["FECHA"] = pd.to_datetime(df_rec["FECHA"], errors="coerce")

    return df_tot, df_rec

def load_historial(hist_file):
    """Leemos todas las hojas del historial y juntamos:
       - Activaciones (si vienen)
       - Recargas (si vienen)
       Renombramos columnas equivalentes a FECHA / MONTO / DN / PLAN
    """
    xls = pd.ExcelFile(hist_file, engine="openpyxl")
    rec_list = []
    act_list = []

    for sh in xls.sheet_names:
        try:
            df = read_with_possible_header(xls, sh, header_try=(0,1,2,3,4))
            df = norm_cols(df)

            # detectar DN
            dn_col = pick_col(df, ["DN","NUMERO","LINEA","MSISDN"])
            if not dn_col: 
                continue
            if dn_col != "DN": df.rename(columns={dn_col:"DN"}, inplace=True)

            # detectar FECHA
            fecha_col = pick_col(df, ["FECHA","FECHA RECARGA","FECHA DE RECARGA","FECHA DE PAGO","FECHA_PAGO","FECHA "])
            if fecha_col:
                if fecha_col != "FECHA": df.rename(columns={fecha_col:"FECHA"}, inplace=True)
                df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")

            # detectar MONTO (para recargas)
            monto_col = pick_col(df, ["MONTO","IMPORTE","MONTO RECARGA","CANTIDAD","PAGO"])
            if monto_col and monto_col != "MONTO":
                df.rename(columns={monto_col:"MONTO"}, inplace=True)

            # PLAN si existiera
            plan_col = pick_col(df, ["PLAN","PLAN TARIFARIO"])
            if plan_col and plan_col != "PLAN":
                df.rename(columns={plan_col:"PLAN"}, inplace=True)

            # Heurística: si trae MONTO y FECHA, lo consideramos recargas
            if "FECHA" in df.columns and "MONTO" in df.columns:
                sub = df[["FECHA","DN","MONTO"] + ([ "PLAN" ] if "PLAN" in df.columns else [])].copy()
                rec_list.append(sub)
            # Heurística de activaciones: presencia de FECHA + PLAN, o columnas típicas
            elif "FECHA" in df.columns and ("PLAN" in df.columns or "COSTO PAQUETE" in df.columns):
                keep = ["FECHA","DN"]
                if "PLAN" in df.columns: keep.append("PLAN")
                if "COSTO PAQUETE" in df.columns: keep.append("COSTO PAQUETE")
                act_list.append(df[keep].copy())

        except Exception:
            continue

    rec_hist = pd.concat(rec_list, ignore_index=True) if rec_list else pd.DataFrame(columns=["FECHA","DN","MONTO","PLAN"])
    act_hist = pd.concat(act_list, ignore_index=True) if act_list else pd.DataFrame(columns=["FECHA","DN","PLAN","COSTO PAQUETE"])

    # DN normalizado
    for d in (rec_hist, act_hist):
        if "DN" in d.columns:
            d["DN_NORM"] = normalize_dn(d["DN"])

    return act_hist, rec_hist

# ========== Cálculo del reporte ==========
def calc_report(df_tot, df_rec, dist_name, year, month, act_hist=None, rec_hist=None):
    month_start = pd.Timestamp(year, month, 1)
    month_end   = pd.Timestamp(year, month, 1) + pd.offsets.MonthEnd(1)

    # Universo distribuidor con base mensual
    df_tot = df_tot.copy()
    df_rec = df_rec.copy()
    mask_dist = df_tot["DISTRIBUIDOR "].astype(str).str.strip().str.lower() == dist_name.lower()
    tot_dist = df_tot[mask_dist].copy()

    # Normalización de fechas
    if "FECHA" in df_rec.columns:
        df_rec["FECHA"] = pd.to_datetime(df_rec["FECHA"], errors="coerce")
    if "FECHA" in tot_dist.columns:
        tot_dist["FECHA"] = pd.to_datetime(tot_dist["FECHA"], errors="coerce")

    dns_dist = set(tot_dist["DN_NORM"].dropna())

    # Activaciones del mes (base)
    altas_mes_base = tot_dist[(tot_dist["FECHA"]>=month_start) & (tot_dist["FECHA"]<=month_end)].copy()

    # Recargas del mes (base) solo de esas DN
    rec_month_base = df_rec[(df_rec["FECHA"]>=month_start) & (df_rec["FECHA"]<=month_end)].copy()
    rec_month_base = rec_month_base[rec_month_base["DN_NORM"].isin(dns_dist)].copy()

    # Historial: combinar si viene info válida
    if act_hist is not None and not act_hist.empty:
        act_hist = act_hist.copy()
        if "FECHA" in act_hist.columns:
            act_hist["FECHA"] = pd.to_datetime(act_hist["FECHA"], errors="coerce")
        act_hist = act_hist[act_hist["DN_NORM"].isin(dns_dist)]
        altas_mes_hist = act_hist[(act_hist["FECHA"]>=month_start) & (act_hist["FECHA"]<=month_end)].copy()
    else:
        altas_mes_hist = pd.DataFrame(columns=altas_mes_base.columns)

    if rec_hist is not None and not rec_hist.empty:
        rec_hist = rec_hist.copy()
        if "FECHA" in rec_hist.columns:
            rec_hist["FECHA"] = pd.to_datetime(rec_hist["FECHA"], errors="coerce")
        rec_hist = rec_hist[rec_hist["DN_NORM"].isin(dns_dist)]
        rec_month_hist = rec_hist[(rec_hist["FECHA"]>=month_start) & (rec_hist["FECHA"]<=month_end)].copy()
    else:
        rec_month_hist = pd.DataFrame(columns=rec_month_base.columns if not rec_month_base.empty else ["FECHA","DN","MONTO","DN_NORM"])

    # Unimos (preferencia: sumar ambas fuentes)
    altas_mes = pd.concat([altas_mes_base, altas_mes_hist], ignore_index=True) if not altas_mes_hist.empty else altas_mes_base
    rec_month = pd.concat([rec_month_base, rec_month_hist], ignore_index=True) if not rec_month_hist.empty else rec_month_base

    # Clasificación y mínimos
    tot_dist["PRODUCTO"] = tot_dist.apply(classify_row, axis=1)
    n_altas = altas_mes["DN_NORM"].nunique()
    pct_mbb = cartera_pct_mbb(n_altas)
    min_mbb, min_mifi, min_hbb = 35, 110, 99

    # Suma recargas por DN en el mes
    if rec_month.empty:
        rec_by_dn = pd.DataFrame({"DN_NORM": list(dns_dist), "RECARGA_TOTAL_MES": 0.0})
    else:
        rec_by_dn = (
            rec_month.groupby("DN_NORM", as_index=False)["MONTO"]
            .sum()
            .rename(columns={"MONTO":"RECARGA_TOTAL_MES"})
        )

    # ANEXO
    keep_cols = [c for c in ["DN","DN_NORM","FECHA","PLAN","COSTO PAQUETE","PRODUCTO"] if c in tot_dist.columns]
    anexo = tot_dist[keep_cols].merge(rec_by_dn, on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # Elegibilidad
    def elegible(row):
        if row["PRODUCTO"] == "MBB":
            return row["RECARGA_TOTAL_MES"] >= min_mbb
        elif row["PRODUCTO"] == "MiFi":
            return row["RECARGA_TOTAL_MES"] >= min_mifi
        elif row["PRODUCTO"] == "HBB":
            return row["RECARGA_TOTAL_MES"] >= min_hbb
        return False
    anexo["ELEGIBLE_CARTERA"] = anexo.apply(elegible, axis=1)

    # % aplicado
    def pct_aplicado(row):
        if row["PRODUCTO"] == "MBB":
            return pct_mbb
        elif row["PRODUCTO"] in ("MiFi","HBB"):
            return 0.05  # base M1–12
        return 0.0
    anexo["% CARTERA APLICADA"] = anexo.apply(pct_aplicado, axis=1)
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"],
        anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"],
        0.0
    ).round(2)

    # RESUMEN
    resumen = pd.DataFrame([{
        "Distribuidor": dist_name,
        "Mes": f'{month_start.strftime("%B").capitalize()} {year}',
        "Altas del mes": int(n_altas),
        "Recargas totales del mes ($)": round(rec_month["MONTO"].sum() if "MONTO" in rec_month.columns else 0.0, 2),
        "Porcentaje Cartera aplicado (MBB)": pct_mbb,
        "Comisión Cartera total ($)": round(anexo["COMISION_CARTERA_$"].sum(), 2)
    }])

    # RESUMEN MES
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
        .agg(Lineas=("DN_NORM","nunique"),
             Recarga_Mes_$=("RECARGA_TOTAL_MES","sum"),
             Comision_Mes_$=("COMISION_CARTERA_$","sum"))
    )
    total_row = pd.DataFrame([{
        "PRODUCTO": "TOTAL",
        "Lineas": resumen_mes["Lineas"].sum(),
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
        "Comision_Mes_$": resumen_mes["Comision_Mes_$"].sum()
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # HISTORIAL ACTIVACIONES (solo mes)
    if not altas_mes.empty:
        hist_cols = ["FECHA","DN_NORM"] + [c for c in ["PLAN","COSTO PAQUETE"] if c in altas_mes.columns]
        hist = altas_mes[hist_cols].rename(columns={"DN_NORM":"DN"}).sort_values("FECHA")
    else:
        hist = pd.DataFrame(columns=["FECHA","DN","PLAN","COSTO PAQUETE"])

    # CARTERA MES (detalle recargas)
    if not rec_month.empty:
        rec_det = rec_month.copy()
        rec_det["ELEGIBLE_MBB"] = rec_det["MONTO"] >= min_mbb
        # Traemos PLAN si se conoce por totales
        plan_map = tot_dist[["DN_NORM","PLAN"]].dropna().drop_duplicates()
        rec_det = rec_det.merge(plan_map, on="DN_NORM", how="left")
        rec_det = rec_det[["FECHA","DN_NORM","PLAN","MONTO"] + (["FORMA DE PAGO"] if "FORMA DE PAGO" in rec_det.columns else []) + ["ELEGIBLE_MBB"]]
        rec_det = rec_det.rename(columns={"DN_NORM":"DN"}).sort_values("FECHA")
    else:
        rec_det = pd.DataFrame(columns=["FECHA","DN","PLAN","MONTO","ELEGIBLE_MBB"])

    # Exportar a Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        anexo.to_excel(writer, sheet_name="ANEXO", index=False)
        hist.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)
        resumen_mes.to_excel(writer, sheet_name=f'RESUMEN {month_start.strftime("%B").upper()} {year}', index=False)
        rec_det.to_excel(writer, sheet_name=f'CARTERA {month_start.strftime("%B").upper()} {year}', index=False)
    output.seek(0)
    return output

# ========== UI ==========
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("📥 Base mensual (VT Reporte Comercial…)", type=["xlsx"])
    st.caption("Debe tener 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")
    hist_file = st.file_uploader("📥 Historial del distribuidor (opcional, recomendado)", type=["xlsx"])
with col2:
    st.write("Parámetros")
    dist = st.text_input("Distribuidor (igual que en la base)", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

# Info de columnas detectadas (ayuda para debug)
def show_detected(name, df):
    st.markdown(f"**{name}** columnas: " + ", ".join(df.columns[:20]))

if base_file:
    try:
        df_tot, df_rec = load_base(base_file)
        with st.expander("Columnas detectadas (Base)"):
            show_detected("Desgloce Totales", df_tot)
            show_detected("Desgloce Recarga", df_rec)
    except Exception as e:
        st.error(f"Error leyendo Base: {e}")

act_hist = None
rec_hist = None
if hist_file:
    try:
        ah, rh = load_historial(hist_file)
        act_hist, rec_hist = ah, rh
        with st.expander("Columnas detectadas (Historial)"):
            if not act_hist.empty: show_detected("Activaciones (hist)", act_hist)
            if not rec_hist.empty: show_detected("Recargas (hist)", rec_hist)
            if act_hist.empty and rec_hist.empty:
                st.info("No se detectaron hojas de historial con columnas DN/FECHA/(MONTO). Se usará solo la base mensual.")
    except Exception as e:
        st.warning(f"No se pudo interpretar el historial: {e}")

if base_file and st.button("Generar reporte"):
    try:
        buf = calc_report(
            df_tot=df_tot,
            df_rec=df_rec,
            dist_name=dist,
            year=int(year),
            month=int(month),
            act_hist=act_hist,
            rec_hist=rec_hist
        )
        fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {year}.xlsx"
        st.success("✅ Reporte generado.")
        st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
