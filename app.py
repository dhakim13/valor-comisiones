import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re

# =========================
# Configuración de página
# =========================
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga el Excel con HISTORIAL DE ACTIVACIONES y las hojas de CARTERA por mes. Elige el mes y genera el archivo con RESUMEN, ANEXO y detalle.")

# =========================
# Helpers
# =========================
def normalize_dn(s: pd.Series) -> pd.Series:
    out = s.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    def fix(x: str) -> str:
        try:
            xl = x.lower()
            if 'e+' in xl or 'e-' in xl:
                return str(int(float(x)))
            return x.split('.')[0]
        except Exception:
            return x
    return out.apply(fix)

def classify_producto(tipo: str, costo) -> str:
    t = str(tipo or "").upper()
    if "MOB" in t:
        return "MBB"
    try:
        c = float(costo)
    except Exception:
        c = np.nan
    # Heurística por costo (ajústala si cambia)
    if c in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    if c in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    return "MBB"

def cartera_pct_mbb(n_altas_mes: int) -> float:
    if n_altas_mes < 50:   return 0.03
    if n_altas_mes < 300:  return 0.05
    if n_altas_mes < 500:  return 0.07
    if n_altas_mes < 1000: return 0.08
    return 0.10

def month_bounds(year: int, month: int):
    start = pd.Timestamp(year, month, 1)
    end = start + pd.offsets.MonthEnd(1)
    return start, end

def spanish_month_name(ts: pd.Timestamp) -> str:
    meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
             "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
    return meses[ts.month-1]

# =========================
# Núcleo de cálculo
# =========================
def calc_report(xls: pd.ExcelFile, year: int, month: int, dist_filtro: str = "") -> BytesIO:
    month_start, month_end = month_bounds(year, month)
    mes_mayus = spanish_month_name(month_start)

    # ---------- Leer HISTORIAL DE ACTIVACIONES ----------
    # Permitimos variantes de nombre (con / sin tildes / mayúsculas)
    hist_name = None
    for s in xls.sheet_names:
        if re.sub(r"\s+", "", s).upper() in {
            "HISTORIALDEACTIVACIONES",
            "HISTORIALDEACTIVACION",
            "HISTORIALACTIVACIONES"
        }:
            hist_name = s
            break
    if hist_name is None:
        raise ValueError("No se encontró la hoja 'HISTORIAL DE ACTIVACIONES' en el archivo.")

    hist = pd.read_excel(xls, sheet_name=hist_name)
    # Normalizar columnas esperadas
    cols = {c.upper().strip(): c for c in hist.columns}
    def pick(*cands):
        for c in cands:
            if c in cols:
                return cols[c]
        return None

    col_fecha_a = pick("FECHA","FECHA DE ALTA","ALTA")
    col_dn      = pick("DN","NUMERO","LÍNEA","LINEA")
    col_plan    = pick("PLAN",)
    col_costo   = pick("COSTO PAQUETE","COSTO","PRECIO")
    col_tipo    = pick("TIPO",)
    col_porta   = pick("DN PORTADO","PORTADO","PORTABILIDAD")

    for c,nom in [(col_fecha_a,"FECHA (alta)"),(col_dn,"DN")]:
        if c is None:
            raise ValueError(f"En HISTORIAL DE ACTIVACIONES falta la columna '{nom}'.")

    hist = hist.rename(columns={
        col_fecha_a:"FECHA_ALTA",
        col_dn:"DN",
        **({col_plan:"PLAN"} if col_plan else {}),
        **({col_costo:"COSTO PAQUETE"} if col_costo else {}),
        **({col_tipo:"TIPO"} if col_tipo else {}),
        **({col_porta:"DN PORTADO"} if col_porta else {})
    })

    hist["FECHA_ALTA"] = pd.to_datetime(hist["FECHA_ALTA"], errors="coerce")
    hist["DN"] = normalize_dn(hist["DN"])
    if "DN PORTADO" in hist.columns:
        hist["ES_PORTADO"] = hist["DN PORTADO"].astype(str).str.strip().ne("")
    else:
        hist["ES_PORTADO"] = False

    # Clasificación de producto
    if "PRODUCTO" not in hist.columns:
        hist["PRODUCTO"] = hist.apply(lambda r: classify_producto(r.get("TIPO",""), r.get("COSTO PAQUETE", np.nan)), axis=1)

    # Filtro de distribuidor (opcional): si hay columna DISTRIBUIDOR
    if dist_filtro:
        col_dist = None
        for c in hist.columns:
            if c.strip().upper().startswith("DISTRIB"):
                col_dist = c
                break
        if col_dist:
            hist = hist[hist[col_dist].astype(str).str.strip().str.lower()==dist_filtro.lower()].copy()

    # Altas del mes (para % MBB y para portabilidad del mes)
    altas_mes = hist[(hist["FECHA_ALTA"]>=month_start) & (hist["FECHA_ALTA"]<=month_end)].copy()
    n_altas_mes = altas_mes["DN"].nunique()
    pct_mbb = cartera_pct_mbb(n_altas_mes)

    # Portabilidad del mes (pago fijo por alta portada en el mes)
    PAGO_PORTA = 30.0
    n_porta_mes = altas_mes.loc[altas_mes["ES_PORTADO"], "DN"].nunique()
    com_porta = n_porta_mes * PAGO_PORTA

    # ---------- Consolidar TODAS las hojas de CARTERA (histórico de recargas) ----------
    rec_list = []
    for s in xls.sheet_names:
        if re.sub(r"\s+", "", s).upper().startswith("CARTERA"):
            df = pd.read_excel(xls, sheet_name=s)
            rec_list.append(df)

    if not rec_list:
        raise ValueError("No se encontraron hojas de 'CARTERA ...' con el historial de recargas.")

    rec = pd.concat(rec_list, ignore_index=True)
    # Normalizar columnas en recargas
    cols_r = {c.upper().strip(): c for c in rec.columns}
    col_fecha_r = next((cols_r[c] for c in ["FECHA","FECHA RECARGA","FECHA_RECA","FECHA DE RECARGA"] if c in cols_r), None)
    col_dn_r    = next((cols_r[c] for c in ["DN","NUMERO","LÍNEA","LINEA"] if c in cols_r), None)
    col_monto   = next((cols_r[c] for c in ["MONTO","IMPORTE","CANTIDAD"] if c in cols_r), None)
    col_plan_r  = next((cols_r[c] for c in ["PLAN"] if c in cols_r), None)
    col_fpago   = next((cols_r[c] for c in ["FORMA DE PAGO","PAGO","FPAGO"] if c in cols_r), None)

    for c,nom in [(col_fecha_r,"FECHA (recarga)"),(col_dn_r,"DN"),(col_monto,"MONTO")]:
        if c is None:
            raise ValueError(f"En hojas CARTERA falta la columna '{nom}'.")

    rec = rec.rename(columns={
        col_fecha_r:"FECHA",
        col_dn_r:"DN",
        col_monto:"MONTO",
        **({col_plan_r:"PLAN"} if col_plan_r else {}),
        **({col_fpago:"FORMA DE PAGO"} if col_fpago else {})
    })
    rec["FECHA"] = pd.to_datetime(rec["FECHA"], errors="coerce")
    rec["DN"] = normalize_dn(rec["DN"])
    rec["MONTO"] = pd.to_numeric(rec["MONTO"], errors="coerce").fillna(0.0)

    # ---------- Enriquecer recargas con fecha de ALTA y PRODUCTO ----------
    rec = rec.merge(hist[["DN","FECHA_ALTA","PRODUCTO"]], on="DN", how="left")
    rec["EDAD_MESES"] = ((rec["FECHA"] - rec["FECHA_ALTA"]) / np.timedelta64(1, "M")).fillna(-999)
    # +M2 = edad >= 2 meses
    rec["+M2"] = rec["EDAD_MESES"] >= 2 - 1e-9  # pequeño epsilon por redondeo

    # ---------- Base del mes seleccionado ----------
    rec_mes = rec[(rec["FECHA"]>=month_start) & (rec["FECHA"]<=month_end)].copy()

    # ---------- 1ª recarga ($15) ----------
    PAGO_PRIMERA = 15.0
    # fecha de 1ª recarga histórica por DN
    primera_fecha = rec.groupby("DN", as_index=False)["FECHA"].min().rename(columns={"FECHA":"FECHA_1RA"})
    rec_mes = rec_mes.merge(primera_fecha, on="DN", how="left")
    es_primera_en_mes = (rec_mes["FECHA"].dt.date == rec_mes["FECHA_1RA"].dt.date)
    n_primera = rec_mes.loc[es_primera_en_mes, "DN"].nunique()
    com_primera = n_primera * PAGO_PRIMERA

    # ---------- Cartera (+M2) y mínimos por producto ----------
    MIN_MBB = 35.0
    MIN_MIFI = 110.0
    MIN_HBB = 99.0

    # Suma de recargas del mes por línea
    by_line = rec_mes.groupby(["DN","PRODUCTO","+M2"], as_index=False)["MONTO"].sum().rename(columns={"MONTO":"RECARGA_TOTAL_MES"})

    # Elegibilidad por mínimos
    def cumple_min(row):
        if row["PRODUCTO"] == "MBB": return row["RECARGA_TOTAL_MES"] >= MIN_MBB
        if row["PRODUCTO"] == "MiFi": return row["RECARGA_TOTAL_MES"] >= MIN_MIFI
        if row["PRODUCTO"] == "HBB": return row["RECARGA_TOTAL_MES"] >= MIN_HBB
        return False

    by_line["ELEGIBLE_MIN"] = by_line.apply(cumple_min, axis=1)

    # Base +M2: solo líneas con +M2, elegibles por mínimo
    base_m2 = by_line[(by_line["+M2"]) & (by_line["ELEGIBLE_MIN"])].copy()

    # % por producto
    base_m2["PCT"] = np.where(base_m2["PRODUCTO"]=="MBB", pct_mbb, 0.05)
    base_m2["COMISION_CARTERA_$"] = (base_m2["RECARGA_TOTAL_MES"] * base_m2["PCT"]).round(2)

    # Totales cartera por producto
    cartera_prod = base_m2.groupby("PRODUCTO", as_index=False).agg({
        "DN":"nunique",
        "RECARGA_TOTAL_MES":"sum",
        "COMISION_CARTERA_$":"sum"
    }).rename(columns={"DN":"Lineas"})

    # Fila total
    total_row = pd.DataFrame([{
        "PRODUCTO":"TOTAL",
        "Lineas": int(cartera_prod["Lineas"].sum()) if not cartera_prod.empty else 0,
        "RECARGA_TOTAL_MES": float(cartera_prod["RECARGA_TOTAL_MES"].sum()) if not cartera_prod.empty else 0.0,
        "COMISION_CARTERA_$": float(cartera_prod["COMISION_CARTERA_$"].sum()) if not cartera_prod.empty else 0.0
    }])
    cartera_resumen = pd.concat([cartera_prod, total_row], ignore_index=True)

    # ---------- RESUMEN GENERAL ----------
    resumen = pd.DataFrame([{
        "Mes": f"{mes_mayus} {year}",
        "Altas del mes": int(n_altas_mes),
        "% Cartera MBB aplicado": pct_mbb,
        "Cartera +M2 ($)": float(cartera_resumen.loc[cartera_resumen["PRODUCTO"]=="TOTAL","COMISION_CARTERA_$"].values[0] if "TOTAL" in cartera_resumen["PRODUCTO"].values else cartera_resumen["COMISION_CARTERA_$"].sum()),
        "1ra Recarga ($)": float(com_primera),
        "Portabilidad ($)": float(com_porta),
        "Total Comisión Mes ($)": float(
            (cartera_resumen["COMISION_CARTERA_$"].sum() if "TOTAL" not in cartera_resumen["PRODUCTO"].values else cartera_resumen.loc[cartera_resumen["PRODUCTO"]=="TOTAL","COMISION_CARTERA_$"].values[0])
            + com_primera + com_porta
        )
    }])

    # ---------- ANEXO (detalle por línea) ----------
    # Unimos by_line con info de alta y si fue primera recarga en el mes
    altas_lookup = hist[["DN","FECHA_ALTA","ES_PORTADO","PRODUCTO","PLAN","COSTO PAQUETE"]]
    anexo = by_line.merge(altas_lookup, on="DN", how="left")
    # marcar si DN tuvo 1ª recarga en el mes
    first_in_month = rec_mes.loc[es_primera_en_mes, ["DN"]].drop_duplicates().assign(PRIMERA_RECARGA_MES=True)
    anexo = anexo.merge(first_in_month, on="DN", how="left")
    anexo["PRIMERA_RECARGA_MES"] = anexo["PRIMERA_RECARGA_MES"].fillna(False)
    # marcar si activó portabilidad en el mes
    porta_mes_dns = altas_mes.loc[altas_mes["ES_PORTADO"], ["DN"]].drop_duplicates().assign(PORTADO_EN_MES=True)
    anexo = anexo.merge(porta_mes_dns, on="DN", how="left")
    anexo["PORTADO_EN_MES"] = anexo["PORTADO_EN_MES"].fillna(False)

    anexo["%_APLICADO"] = np.where(anexo["PRODUCTO"]=="MBB", pct_mbb, 0.05)
    anexo["COMISION_CARTERA_$"] = np.where(
        (anexo["+M2"]) & (anexo["ELEGIBLE_MIN"]),
        (anexo["RECARGA_TOTAL_MES"] * anexo["%_APLICADO"]).round(2),
        0.0
    )
    anexo["COMISION_1RA_$"] = np.where(anexo["PRIMERA_RECARGA_MES"], PAGO_PRIMERA, 0.0)
    anexo["COMISION_PORTA_$"] = np.where(anexo["PORTADO_EN_MES"], PAGO_PORTA, 0.0)
    anexo["COMISION_TOTAL_$"] = (anexo["COMISION_CARTERA_$"] + anexo["COMISION_1RA_$"] + anexo["COMISION_PORTA_$"]).round(2)

    # ---------- HISTORIAL DEL MES (activaciones del mes seleccionado) ----------
    historial_mes = altas_mes[["FECHA_ALTA","DN","PRODUCTO","PLAN","COSTO PAQUETE","ES_PORTADO"]].rename(
        columns={"FECHA_ALTA":"FECHA","ES_PORTADO":"PORTADO"}
    ).sort_values("FECHA")

    # ---------- RECARGAS DETALLE DEL MES ----------
    rec_det_mes = rec_mes[["FECHA","DN","PLAN","MONTO","FORMA DE PAGO","PRODUCTO","+M2"]].sort_values("FECHA")
    rec_det_mes = rec_det_mes.rename(columns={"+M2":"ES_MAS_M2"})

    # ---------- Exportar a Excel en memoria ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        cartera_resumen.rename(columns={
            "RECARGA_TOTAL_MES":"Recarga_Mes_$",
            "COMISION_CARTERA_$":"Comision_Mes_$"
        }).to_excel(writer, sheet_name=f"RESUMEN {mes_mayus} {year}", index=False)

        anexo_cols = [
            "DN","PRODUCTO","PLAN","COSTO PAQUETE","FECHA_ALTA",
            "RECARGA_TOTAL_MES","ELEGIBLE_MIN","+M2","%_APLICADO",
            "PRIMERA_RECARGA_MES","PORTADO_EN_MES",
            "COMISION_CARTERA_$","COMISION_1RA_$","COMISION_PORTA_$","COMISION_TOTAL_$"
        ]
        # mantener solo columnas existentes
        anexo[[c for c in anexo_cols if c in anexo.columns]].to_excel(writer, sheet_name="ANEXO", index=False)

        historial_mes.to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)
        rec_det_mes.to_excel(writer, sheet_name=f"CARTERA {mes_mayus} {year}", index=False)

    output.seek(0)
    return output

# =========================
# UI
# =========================
with st.container():
    base_file = st.file_uploader("📄 Excel con HISTORIAL DE ACTIVACIONES y hojas de CARTERA…", type=["xlsx","xlsm"])
    colA, colB, colC = st.columns([1,1,1])
    with colA:
        year = st.number_input("Año", min_value=2023, max_value=2100, value=datetime.now().year, step=1)
    with colB:
        month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=datetime.now().month, step=1)
    with colC:
        dist = st.text_input("Distribuidor (opcional para filtrar altas)", value="")

    if base_file and st.button("Generar reporte"):
        try:
            xls = pd.ExcelFile(base_file, engine="openpyxl")
            buf = calc_report(xls, int(year), int(month), dist_filtro=dist.strip())
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper() or 'GENERAL'} {spanish_month_name(pd.Timestamp(int(year),int(month),1))} {int(year)}.xlsx"
            st.success("✅ Reporte generado.")
            st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error("Ocurrió un error al generar el reporte.")
            st.exception(e)
