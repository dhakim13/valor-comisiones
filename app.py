import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# -------------------- Config --------------------
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y (opcional) el archivo histórico/plantilla del distribuidor. Genera RESUMEN, ANEXO, HISTORIAL, RESUMEN {MES} y CARTERA {MES} con cálculos de comisiones.")

# -------------------- Constantes negocio --------------------
MIN_REC_MBB = 35
MIN_REC_MIFI = 110
MIN_REC_HBB  = 99

PRIMERA_RECARGA_BONO = 15.0
PORTABILIDAD_BONO    = 30.0   # a partir de este mes

# -------------------- Utilidades --------------------
def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def try_read_xlsx(xls, sheet_name, header_candidates=(0,1,2,3,4,5,6,7,8,9)):
    for h in header_candidates:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=h, engine="openpyxl")
            df = normalize_headers(df)
            if any(not str(c).startswith("Unnamed") for c in df.columns):
                return df
        except Exception:
            continue
    df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
    return normalize_headers(df)

def first_existing_col(df: pd.DataFrame, candidates):
    cols = list(df.columns)
    low = {c.strip().lower(): c for c in cols}
    for c in candidates:
        if c in df.columns:
            return c
        key = str(c).strip().lower()
        if key in low:
            return low[key]
    raise KeyError(f"No encontré ninguna de las columnas {candidates} en: {list(df.columns)}")

def find_column_contains(df: pd.DataFrame, must_contain: list[str]):
    """
    Busca una columna que contenga TODOS los fragmentos de 'must_contain' (insensible a may/min y acentos simples).
    p.ej. must_contain=["fecha","activ"] => "FECHA DE ACTIVACION"
    """
    def norm(s):
        t = str(s).lower()
        repl = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u"}
        for a,b in repl.items(): t = t.replace(a,b)
        return t
    cols = list(df.columns)
    cols_norm = [(c, norm(c)) for c in cols]
    must = [norm(x) for x in must_contain]
    for original, cn in cols_norm:
        if all(m in cn for m in must):
            return original
    return None

def normalize_dn_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    def fix(x):
        try:
            if "e+" in x.lower():
                return str(int(float(x)))
            return x.split(".")[0]
        except:
            return x
    return s.apply(fix)

def month_name_es(year:int, month:int) -> str:
    return pd.Timestamp(year, month, 1).strftime("%B").capitalize()

def cartera_pct_mbb(n_altas_mes:int) -> float:
    if n_altas_mes < 50:   return 0.03
    if n_altas_mes < 300:  return 0.05
    if n_altas_mes < 500:  return 0.07
    if n_altas_mes < 1000: return 0.08
    return 0.10

def window_index_30d(fecha_alta: pd.Timestamp, fecha_evento: pd.Timestamp) -> int:
    if pd.isna(fecha_alta) or pd.isna(fecha_evento):
        return -1
    days = (fecha_evento - fecha_alta).days + 1
    if days <= 0:
        return -1
    return (days - 1) // 30   # 0=M (1–30), 1=M+1 (31–60), 2=M+2 (61–90)

def classify_producto(tipo_val, costo_val):
    tipo = str(tipo_val).upper()
    try:
        costo = float(costo_val)
    except:
        costo = np.nan
    if "MOB" in tipo or "MBB" in tipo or "MOVIL" in tipo:
        return "MBB"
    if costo in [99,115,349,399,439,500]:
        return "HBB"
    if costo in [110,120,160,245,375,480,620]:
        return "MiFi"
    return "MBB"

# -------------------- Cálculo --------------------
def calc_report(xls_base, xls_tpl, year:int, month:int, dist_filtro:str):
    mes_inicio = pd.Timestamp(year, month, 1)
    mes_fin    = mes_inicio + pd.offsets.MonthEnd(1)

    # ---- BASE mensual ----
    tot = try_read_xlsx(xls_base, "Desgloce Totales", header_candidates=(1,2,0))
    rec = try_read_xlsx(xls_base, "Desgloce Recarga", header_candidates=(3,4,2,1,0))
    tot = normalize_headers(tot)
    rec = normalize_headers(rec)

    col_fecha_tot = first_existing_col(tot, ["FECHA", "FECHA ALTA", "ALTA", "F. ALTA"])
    col_dn_tot    = first_existing_col(tot, ["DN", "NUMERO", "MSISDN"])
    # PLAN puede venir como PLAN o PAQUETE en algunas bases; intentamos PLAN primero
    col_plan_tot  = col if (col:=("PLAN" if "PLAN" in tot.columns else ("PAQUETE" if "PAQUETE" in tot.columns else None))) else first_existing_col(tot, ["PLAN","PAQUETE"])
    col_costo     = first_existing_col(tot, ["COSTO PAQUETE", "COSTO"])
    col_tipo_tot  = first_existing_col(tot, ["TIPO"])
    # DISTRIBUIDOR con (o sin) espacio al final
    col_dist      = "DISTRIBUIDOR " if "DISTRIBUIDOR " in tot.columns else first_existing_col(tot, ["DISTRIBUIDOR"])

    col_fecha_rec = first_existing_col(rec, ["FECHA"])
    col_dn_rec    = first_existing_col(rec, ["DN", "NUMERO", "MSISDN"])
    col_plan_rec  = col_plan_tot if "PLAN" in rec.columns else None
    col_monto_rec = first_existing_col(rec, ["MONTO", "IMPORTE", "MONTO RECARGA"])
    col_fpago_rec = first_existing_col(rec, ["FORMA DE PAGO", "METODO PAGO", "FORMA PAGO"])

    tot["FECHA_ALTA"] = pd.to_datetime(tot[col_fecha_tot], errors="coerce")
    rec["FECHA_REC"]  = pd.to_datetime(rec[col_fecha_rec], errors="coerce")
    tot["DN_NORM"]    = normalize_dn_series(tot[col_dn_tot])
    rec["DN_NORM"]    = normalize_dn_series(rec[col_dn_rec])
    tot["PRODUCTO"]   = tot.apply(lambda r: classify_producto(r.get(col_tipo_tot, ""), r.get(col_costo, np.nan)), axis=1)

    mask_dist = tot[col_dist].astype(str).str.strip().str.lower() == dist_filtro.strip().lower()
    tot_dist  = tot[mask_dist].copy()
    dns_dist  = set(tot_dist["DN_NORM"].dropna())

    altas_mes = tot_dist[(tot_dist["FECHA_ALTA"]>=mes_inicio) & (tot_dist["FECHA_ALTA"]<=mes_fin)].copy()
    n_altas_mbb = int(altas_mes[altas_mes["PRODUCTO"]=="MBB"]["DN_NORM"].nunique())
    pct_mbb = cartera_pct_mbb(n_altas_mbb)

    rec_mes = rec[(rec["FECHA_REC"]>=mes_inicio) & (rec["FECHA_REC"]<=mes_fin)].copy()
    rec_mes = rec_mes[rec_mes["DN_NORM"].isin(dns_dist)].copy()

    # ---- HISTORIAL / PLANTILLA ----
    hist_altas = pd.DataFrame(columns=["DN_NORM","FECHA_ALTA","DN PORTADO"])
    rec_hist_all = []

    if xls_tpl is not None:
        # HISTORIAL DE ACTIVACIONES (con “FECHA DE ACTIVACION”)
        if "HISTORIAL DE ACTIVACIONES" in xls_tpl.sheet_names:
            ha = try_read_xlsx(xls_tpl, "HISTORIAL DE ACTIVACIONES", header_candidates=(0,1,2,3,4,5,6,7,8))
            ha = normalize_headers(ha)

            col_dn_hist = first_existing_col(ha, ["DN", "MSISDN", "NUMERO"])
            # Buscar una columna que contenga "fecha" y "activ" (ej. "FECHA DE ACTIVACION")
            col_fa_hist = find_column_contains(ha, ["fecha","activ"])
            if col_fa_hist is None:
                # fallback genérico (por si acaso)
                col_fa_hist = first_existing_col(ha, ["FECHA DE ACTIVACION", "FECHA ACTIVACION", "FECHA (alta)", "FECHA", "ALTA", "F. ALTA"])

            # Portabilidad (suele ser "DN PORTADO")
            col_porta = None
            for c in ha.columns:
                if c.strip().upper().startswith("DN PORTADO"):
                    col_porta = c
                    break
            if col_porta is None:
                col_porta = "DN PORTADO"
                if col_porta not in ha.columns:
                    ha[col_porta] = ""

            ha["DN_NORM"]    = normalize_dn_series(ha[col_dn_hist])
            ha["FECHA_ALTA"] = pd.to_datetime(ha[col_fa_hist], errors="coerce")
            ha = ha[["DN_NORM","FECHA_ALTA", col_porta]].rename(columns={col_porta:"DN PORTADO"})
            hist_altas = ha.copy()

        # Recargas históricas: todas las hojas cuyo nombre empiece con "CARTERA"
        for s in xls_tpl.sheet_names:
            if str(s).strip().upper().startswith("CARTERA"):
                rr = try_read_xlsx(xls_tpl, s, header_candidates=(0,1,2,3,4,5,6))
                rr = normalize_headers(rr)
                try:
                    c_f = first_existing_col(rr, ["FECHA"])
                    c_d = first_existing_col(rr, ["DN","MSISDN","NUMERO"])
                    c_m = first_existing_col(rr, ["MONTO","IMPORTE","MONTO RECARGA"])
                    rr["FECHA_REC"] = pd.to_datetime(rr[c_f], errors="coerce")
                    rr["DN_NORM"]   = normalize_dn_series(rr[c_d])
                    rr["MONTO"]     = pd.to_numeric(rr[c_m], errors="coerce").fillna(0.0)
                    rec_hist_all.append(rr[["FECHA_REC","DN_NORM","MONTO"]])
                except Exception:
                    continue

    # Altas por DN (si no hay historial, tomamos la mínima de la base)
    if hist_altas.empty:
        altas_by_dn = tot_dist[["DN_NORM","FECHA_ALTA"]].dropna(subset=["DN_NORM"]).sort_values("FECHA_ALTA")
        altas_by_dn = altas_by_dn.groupby("DN_NORM", as_index=False)["FECHA_ALTA"].min()
        altas_by_dn["DN PORTADO"] = ""
    else:
        altas_by_dn = hist_altas.dropna(subset=["DN_NORM"]).sort_values("FECHA_ALTA")
        altas_by_dn = altas_by_dn.groupby(["DN_NORM","DN PORTADO"], as_index=False)["FECHA_ALTA"].min()

    # Recargas históricas consolidadas + todas las recargas de base (para detectar 1a recarga real)
    rec_all = []
    if rec_hist_all:
        rec_all.append(pd.concat(rec_hist_all, ignore_index=True))
    rec_base_all = rec.copy()
    rec_base_all = rec_base_all[["FECHA_REC","DN_NORM",col_monto_rec]].rename(columns={col_monto_rec:"MONTO"})
    rec_base_all["MONTO"] = pd.to_numeric(rec_base_all["MONTO"], errors="coerce").fillna(0.0)
    rec_all.append(rec_base_all)

    rec_hist_total = pd.concat(rec_all, ignore_index=True)
    rec_hist_total = rec_hist_total.dropna(subset=["DN_NORM"])

    first_topup = (rec_hist_total[rec_hist_total["MONTO"]>0]
                   .groupby("DN_NORM", as_index=False)["FECHA_REC"].min()
                   .rename(columns={"FECHA_REC":"FECHA_PRIMERA_REC"}))
    first_topup["FIRST_IN_MONTH"] = (first_topup["FECHA_PRIMERA_REC"]>=mes_inicio) & (first_topup["FECHA_PRIMERA_REC"]<=mes_fin)

    # Enriquecer recargas del mes
    prod_by_dn = tot_dist[["DN_NORM","PRODUCTO",col_costo]].drop_duplicates("DN_NORM")
    rec_mes_enr = rec_mes.merge(altas_by_dn[["DN_NORM","FECHA_ALTA","DN PORTADO"]], on="DN_NORM", how="left")
    rec_mes_enr = rec_mes_enr.merge(prod_by_dn, on="DN_NORM", how="left")
    rec_mes_enr["WIN30"] = rec_mes_enr.apply(lambda r: window_index_30d(r["FECHA_ALTA"], r["FECHA_REC"]), axis=1)

    def cumple_min(row):
        p = row.get("PRODUCTO","MBB")
        m = pd.to_numeric(row.get(col_monto_rec), errors="coerce")
        if pd.isna(m): m=0.0
        if p == "MBB":  return m >= MIN_REC_MBB
        if p == "MiFi": return m >= MIN_REC_MIFI
        if p == "HBB":  return m >= MIN_REC_HBB
        return False
    rec_mes_enr["ELEGIBLE_MIN"] = rec_mes_enr.apply(cumple_min, axis=1)

    # Base +M2 (MBB, ventana==2)
    base_m2_mbb = rec_mes_enr.loc[(rec_mes_enr["PRODUCTO"]=="MBB") & (rec_mes_enr["WIN30"]==2) & (rec_mes_enr["ELEGIBLE_MIN"]), col_monto_rec].astype(float).sum()
    # Base MiFi/HBB (5%) ventanas 0..11
    base_mifi = rec_mes_enr.loc[(rec_mes_enr["PRODUCTO"]=="MiFi") & (rec_mes_enr["WIN30"].between(0,11)) & (rec_mes_enr["ELEGIBLE_MIN"]), col_monto_rec].astype(float).sum()
    base_hbb  = rec_mes_enr.loc[(rec_mes_enr["PRODUCTO"]=="HBB")  & (rec_mes_enr["WIN30"].between(0,11)) & (rec_mes_enr["ELEGIBLE_MIN"]), col_monto_rec].astype(float).sum()

    comi_cartera = (cartera_pct_mbb(n_altas_mbb) * base_m2_mbb) + 0.05 * (base_mifi + base_hbb)

    # 1ª recarga
    first_in_month_dns = set(first_topup.loc[first_topup["FIRST_IN_MONTH"], "DN_NORM"])
    n_first = len(first_in_month_dns & dns_dist)
    comi_first = PRIMERA_RECARGA_BONO * n_first

    # Portabilidad (altas del mes con DN PORTADO != vacío)
    if not hist_altas.empty:
        altas_mes_hist = hist_altas[(hist_altas["FECHA_ALTA"]>=mes_inicio) & (hist_altas["FECHA_ALTA"]<=mes_fin)]
        altas_mes_hist = altas_mes_hist[altas_mes_hist["DN_NORM"].isin(dns_dist)]
        n_porta = int(altas_mes_hist["DN PORTADO"].astype(str).str.strip().replace({"nan":""}).ne("").sum())
    else:
        if "DN PORTADO" in altas_mes.columns:
            n_porta = int(altas_mes["DN PORTADO"].astype(str).str.strip().ne("").sum())
        else:
            n_porta = 0
    comi_porta = PORTABILIDAD_BONO * n_porta

    total_comisiones = round(comi_cartera + comi_first + comi_porta, 2)

    # -------- ANEXO --------
    rec_by_dn = rec_mes.groupby("DN_NORM", as_index=False)[col_monto_rec].sum().rename(columns={col_monto_rec:"RECARGA_TOTAL_MES"})
    anexo = tot_dist[["DN_NORM", col_dn_tot, col_plan_tot, col_costo, "PRODUCTO", "FECHA_ALTA"]].drop_duplicates("DN_NORM")
    anexo = anexo.merge(rec_by_dn, on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    def eleg_cartera_total(row):
        if row["PRODUCTO"] == "MBB":  return row["RECARGA_TOTAL_MES"] >= MIN_REC_MBB
        if row["PRODUCTO"] == "MiFi": return row["RECARGA_TOTAL_MES"] >= MIN_REC_MIFI
        if row["PRODUCTO"] == "HBB":  return row["RECARGA_TOTAL_MES"] >= MIN_REC_HBB
        return False
    anexo["ELEGIBLE_CARTERA"] = anexo.apply(eleg_cartera_total, axis=1)

    def pct_row(row):
        if row["PRODUCTO"] == "MBB":      return cartera_pct_mbb(n_altas_mbb)
        if row["PRODUCTO"] in ("MiFi","HBB"): return 0.05
        return 0.0
    anexo["PCT_APLICADO"] = anexo.apply(pct_row, axis=1)

    def comi_line(row):
        if not row["ELEGIBLE_CARTERA"]:
            return 0.0
        return round(row["RECARGA_TOTAL_MES"] * row["PCT_APLICADO"], 2)
    anexo["COMISION_CARTERA_LINEA"] = anexo.apply(comi_line, axis=1)

    # -------- RESUMEN --------
    resumen = pd.DataFrame([{
        "Distribuidor": dist_filtro,
        "Mes": f"{month_name_es(year,month)} {year}",
        "Altas MBB del mes": n_altas_mbb,
        "Base +M2 (MBB) $": round(base_m2_mbb,2),
        "Base MiFi $ (5%)": round(base_mifi,2),
        "Base HBB $ (5%)": round(base_hbb,2),
        "Pct Cartera MBB": cartera_pct_mbb(n_altas_mbb),
        "Comisión Cartera $": round(comi_cartera,2),
        "Líneas 1ª recarga": n_first,
        "Comisión 1ª recarga $": round(comi_first,2),
        "Líneas portadas (altas mes)": n_porta,
        "Comisión Portabilidad $": round(comi_porta,2),
        "TOTAL COMISIONES $": total_comisiones
    }])

    # -------- RESUMEN {MES} --------
    resumen_mes = (anexo.groupby("PRODUCTO", as_index=False)
                   .agg({"DN_NORM":"nunique",
                         "RECARGA_TOTAL_MES":"sum",
                         "COMISION_CARTERA_LINEA":"sum"})
                   .rename(columns={"DN_NORM":"Lineas",
                                    "RECARGA_TOTAL_MES":"Recarga_Mes_$",
                                    "COMISION_CARTERA_LINEA":"Comision_Mes_$"}))
    total_row = pd.DataFrame([{
        "PRODUCTO":"TOTAL",
        "Lineas": resumen_mes["Lineas"].sum(),
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
        "Comision_Mes_$": resumen_mes["Comision_Mes_$"].sum()
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # -------- HISTORIAL DE ACTIVACIONES (salida) --------
    if hist_altas.empty:
        hist_out = (tot_dist[["FECHA_ALTA","DN_NORM", col_plan_tot, col_costo]]
                    .rename(columns={"FECHA_ALTA":"FECHA (alta)","DN_NORM":"DN",
                                     col_plan_tot:"PLAN", col_costo:"COSTO PAQUETE"})
                    .sort_values("FECHA (alta)"))
        hist_out.insert(2, "DN PORTADO", "")
    else:
        hist_out = hist_altas.rename(columns={"FECHA_ALTA":"FECHA (alta)","DN_NORM":"DN"}).sort_values("FECHA (alta)")

    # -------- CARTERA {MES} --------
    cartera_mes_out = (rec_mes_enr[["FECHA_REC","DN_NORM", col_monto_rec, col_fpago_rec]]
                       .rename(columns={"FECHA_REC":"FECHA","DN_NORM":"DN",
                                        col_monto_rec:"MONTO", col_fpago_rec:"FORMA DE PAGO"})
                       .sort_values("FECHA"))

    # -------- Exportar --------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)

        anexo_out = anexo.rename(columns={
            "DN_NORM":"DN",
            col_dn_tot:"DN_original",
            col_plan_tot:"PLAN",
            col_costo:"COSTO PAQUETE",
            "FECHA_ALTA":"FECHA (alta)",
            "PCT_APLICADO":"% CARTERA APLICADA",
            "COMISION_CARTERA_LINEA":"COMISION_CARTERA_$"
        })
        cols_order = ["DN","DN_original","FECHA (alta)","PLAN","COSTO PAQUETE","PRODUCTO",
                      "RECARGA_TOTAL_MES","ELEGIBLE_CARTERA","% CARTERA APLICADA","COMISION_CARTERA_$"]
        anexo_out = anexo_out.reindex(columns=[c for c in cols_order if c in anexo_out.columns])
        anexo_out.to_excel(writer, sheet_name="ANEXO", index=False)

        hist_cols = ["FECHA (alta)","DN","DN PORTADO"]
        for extra in ["PLAN","COSTO PAQUETE"]:
            if extra in hist_out.columns and extra not in hist_cols:
                hist_cols.append(extra)
        hist_out.reindex(columns=hist_cols).to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)

        nom_res_mes = f"RESUMEN {month_name_es(year,month).upper()} {year}"
        resumen_mes.to_excel(writer, sheet_name=nom_res_mes, index=False)

        nom_cart_mes = f"CARTERA {month_name_es(year,month).upper()} {year}"
        cartera_mes_out.to_excel(writer, sheet_name=nom_cart_mes, index=False)

    output.seek(0)
    return output

# -------------------- UI --------------------
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial…)", type=["xlsx"])
    st.caption("Debe contener: 'Desgloce Totales' (header aprox. fila 2) y 'Desgloce Recarga' (header aprox. fila 4).")
with col2:
    tpl_file = st.file_uploader("Archivo histórico/plantilla del distribuidor", type=["xlsx"])
    dist  = st.text_input("Distribuidor", value="ActivateCel")
    year  = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.markdown("---")
if base_file and st.button("Generar reporte"):
    try:
        xls_base = pd.ExcelFile(base_file, engine="openpyxl")
        xls_tpl  = pd.ExcelFile(tpl_file, engine="openpyxl") if tpl_file is not None else None
        buf = calc_report(xls_base, xls_tpl, int(year), int(month), dist_filtro=dist.strip())
        fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {pd.Timestamp(int(year), int(month), 1).strftime('%B').upper()} {int(year)}.xlsx"
        st.success("✅ Reporte generado con cálculos de comisiones.")
        st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
