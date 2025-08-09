import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")

st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y la PLANTILLA del distribuidor. Calcula y escribe comisiones en las pestañas originales de la plantilla.")

# ----------------- Helpers -----------------
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

def month_span(year:int, month:int):
    start = pd.Timestamp(year, month, 1)
    end   = start + pd.offsets.MonthEnd(1)
    return start, end

def days_since_activation(act_date: pd.Timestamp, ref_date: pd.Timestamp):
    if pd.isna(act_date) or pd.isna(ref_date):
        return np.nan
    return (ref_date.normalize() - act_date.normalize()).days + 1

def cartera_pct_mbb(n_altas_mes:int) -> float:
    # Tiers (sustituyen el 5%; no son adicionales)
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

def classify_producto(tipo:str, costo):
    t = str(tipo or "").upper()
    if "MOB" in t:
        return "MBB"
    # Map por costo de paquete (ajustable)
    # HBB
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    # MiFi
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    # Por defecto asumimos MBB
    return "MBB"

def es_primera_recarga_en_mes(recargas_dn: pd.DataFrame, fecha_rec: pd.Timestamp) -> bool:
    # recargas_dn: todas las recargas de ese DN (historia completa)
    if recargas_dn.empty:
        return False
    first_date = recargas_dn["FECHA"].min()
    return pd.notna(first_date) and (first_date.normalize() == fecha_rec.normalize())

def write_df_over_table(ws, df: pd.DataFrame, header_row:int=1):
    """
    Reemplaza tabla manteniendo ENCABEZADOS en header_row.
    - Borra todo debajo del header y vuelve a escribir contenido.
    - Si la hoja está vacía, escribe encabezados + datos.
    """
    # Detectar si hay encabezados existentes en header_row:
    max_col = df.shape[1]
    # Limpia todo a partir de header_row+1
    ws.delete_rows(idx=header_row+1, amount=max(0, ws.max_row - header_row))
    # Si no hay encabezados en la hoja, escribe nuestros headers
    existing_headers = [c.value for c in ws[header_row]]
    if not any(existing_headers):
        # Hoja sin encabezado: escribimos encabezados de df
        rows = dataframe_to_rows(df, index=False, header=True)
    else:
        # Hoja con encabezado: escribe SOLO datos (sin header)
        rows = dataframe_to_rows(df, index=False, header=False)
    start_row = header_row + (0 if not any(existing_headers) else 1)
    r = start_row
    for row in rows:
        # rows puede incluir filas vacías si header=True; filtramos vacías
        if all([x is None for x in row]):
            continue
        ws.append(row)
        r += 1

def spanish_month_upper(dt: pd.Timestamp) -> str:
    meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
    return meses[dt.month-1]

# ----------------- Core calc -----------------
def calc_report(df_tot: pd.DataFrame,
                df_rec: pd.DataFrame,
                df_rec_hist: pd.DataFrame,
                dist_name: str,
                year: int,
                month: int):
    # Fechas
    month_start, month_end = month_span(year, month)
    month_label_up = f"{spanish_month_upper(month_start)} {year}"

    # Normalización
    dft = df_tot.copy()
    dfr = df_rec.copy()
    dfrh = df_rec_hist.copy() if df_rec_hist is not None else pd.DataFrame(columns=dfr.columns)

    # Cast fechas
    for df in (dft, dfr, dfrh):
        if not df.empty and "FECHA" in df.columns:
            df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")

    # DN normalizado
    for df in (dft, dfr, dfrh):
        if not df.empty and "DN" in df.columns:
            df["DN_NORM"] = normalize_dn(df["DN"])

    # Columnas mínimas esperadas
    # dft: FECHA (activación), PLAN, COSTO PAQUETE, DISTRIBUIDOR , TIPO, DN
    # dfr/dfrh: FECHA (recarga), PLAN, MONTO, FORMA DE PAGO, DN
    # Filtro distribuidor
    if "DISTRIBUIDOR " in dft.columns:
        mask_dist = dft["DISTRIBUIDOR "].astype(str).str.strip().str.lower() == dist_name.lower()
    else:
        # fallback por si columna viene sin espacio final
        mask_dist = dft.get("DISTRIBUIDOR","").astype(str).str.strip().str.lower() == dist_name.lower()

    dft_dist = dft[mask_dist].copy()
    dft_dist["PRODUCTO"] = dft_dist.apply(lambda r: classify_producto(r.get("TIPO",""), r.get("COSTO PAQUETE", np.nan)), axis=1)
    dns_dist = set(dft_dist["DN_NORM"].dropna())

    # Altas del mes (por fecha de activación dentro del mes)
    altas_mes = dft_dist[(dft_dist["FECHA"]>=month_start) & (dft_dist["FECHA"]<=month_end)].copy()
    n_altas_mes = altas_mes["DN_NORM"].nunique()

    # Recargas del mes (base mensual + histórico cruzado por DNs del distribuidor)
    rec_month_base = dfr[(dfr["FECHA"]>=month_start) & (dfr["FECHA"]<=month_end)].copy()
    rec_month_hist = dfrh[(dfrh["FECHA"]>=month_start) & (dfrh["FECHA"]<=month_end)].copy() if not dfrh.empty else pd.DataFrame(columns=rec_month_base.columns)
    rec_month = pd.concat([rec_month_base, rec_month_hist], ignore_index=True)
    rec_month_dist = rec_month[rec_month["DN_NORM"].isin(dns_dist)].copy()

    # % cartera MBB para este distribuidor según altas del mes
    pct_mbb = cartera_pct_mbb(n_altas_mes)

    # Mínimos por producto
    min_mbb = 35
    min_mifi = 110
    min_hbb  = 99

    # ----------------- ANEXO -----------------
    # Suma recarga por DN dentro del mes
    rec_by_dn = rec_month_dist.groupby("DN_NORM", as_index=False)["MONTO"].sum().rename(columns={"MONTO":"RECARGA_TOTAL_MES"})
    anexo = dft_dist[["DN","DN_NORM","FECHA","PLAN","COSTO PAQUETE","PRODUCTO"]].merge(rec_by_dn, on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # Elegibilidad cartera y % aplicado
    def elegible_row(row):
        if row["PRODUCTO"] == "MBB":
            return row["RECARGA_TOTAL_MES"] >= min_mbb
        elif row["PRODUCTO"] == "MiFi":
            return row["RECARGA_TOTAL_MES"] >= min_mifi
        elif row["PRODUCTO"] == "HBB":
            return row["RECARGA_TOTAL_MES"] >= min_hbb
        return False

    def pct_aplicado_row(row):
        if row["PRODUCTO"] == "MBB":
            return pct_mbb
        elif row["PRODUCTO"] in ("MiFi","HBB"):
            return 0.05
        return 0.0

    anexo["ELEGIBLE_CARTERA"] = anexo.apply(elegible_row, axis=1)
    anexo["% CARTERA APLICADA"] = anexo.apply(pct_aplicado_row, axis=1)
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"],
        anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"],
        0.0
    ).round(2)

    # ----------------- Bono MiFi $50 (10GB + 1a recarga día 31–60) -----------------
    # Detectamos potenciales líneas de plan 10GB por Costo Paquete 120
    # y verificamos si su PRIMERA recarga (global) cae dentro de 31–60 días desde activación
    # y además la primera recarga cae dentro del mes calculado (para pagarla en ese mes).
    # Tomamos historia de recargas completa (dfr + dfrh).
    rec_hist_all = pd.concat([dfr, dfrh], ignore_index=True)
    if not rec_hist_all.empty:
        rec_hist_all = rec_hist_all[rec_hist_all["DN_NORM"].isin(dns_dist)].copy()

    # Primera recarga global por DN
    first_rec = (
        rec_hist_all.sort_values("FECHA")
        .groupby("DN_NORM", as_index=False)
        .agg(FirstRecargaFecha=("FECHA","first"),
             FirstRecargaMonto=("MONTO","first"))
    )

    # Merge al anexo
    anexo = anexo.merge(first_rec, on="DN_NORM", how="left")

    # Días desde activación a 1ª recarga
    anexo["DIAS_A_PRIMERA_REC"] = (
        anexo.apply(lambda r: days_since_activation(r["FECHA"], r["FirstRecargaFecha"]), axis=1)
    )

    # Condiciones de bono:
    # - Producto MiFi
    # - Costo Paquete == 120 (10GB)
    # - 1ª recarga entre 31 y 60 días inclusive
    # - Y esa primera recarga ocurre dentro del mes en cálculo (para pagar ese mes)
    cond_bono = (
        (anexo["PRODUCTO"]=="MiFi") &
        (anexo["COSTO PAQUETE"]==120) &
        (anexo["DIAS_A_PRIMERA_REC"].between(31,60, inclusive="both")) &
        (anexo["FirstRecargaFecha"]>=month_start) &
        (anexo["FirstRecargaFecha"]<=month_end)
    )
    anexo["BONO_MIFI_50_$"] = np.where(cond_bono, 50.0, 0.0)

    # Comisión total línea (cartera + bono)
    anexo["COMISION_TOTAL_$"] = (anexo["COMISION_CARTERA_$"] + anexo["BONO_MIFI_50_$"]).round(2)

    # ----------------- RESUMEN -----------------
    resumen = pd.DataFrame([{
        "Distribuidor": dist_name,
        "Mes": f'{month_start.strftime("%B").capitalize()} {year}',
        "Altas del mes": int(n_altas_mes),
        "Recargas totales del mes ($)": round(rec_month_dist["MONTO"].sum(),2),
        "Porcentaje Cartera aplicado (MBB)": pct_mbb,
        "Comisión Cartera total ($)": round(anexo["COMISION_CARTERA_$"].sum(),2),
        "Bonos MiFi ($)": round(anexo["BONO_MIFI_50_$"].sum(),2),
        "Comisión Total ($)": round(anexo["COMISION_TOTAL_$"].sum(),2)
    }])

    # ----------------- RESUMEN MES (por producto) -----------------
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
            .agg(Lineas=("DN_NORM","nunique"),
                 Recarga_Mes_$=("RECARGA_TOTAL_MES","sum"),
                 Comision_Cartera_$=("COMISION_CARTERA_$","sum"),
                 Bono_MiFi_$=("BONO_MIFI_50_$","sum"),
                 Comision_Total_$=("COMISION_TOTAL_$","sum"))
            .sort_values("PRODUCTO")
    )
    tot_row = pd.DataFrame([{
        "PRODUCTO":"TOTAL",
        "Lineas": resumen_mes["Lineas"].sum(),
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
        "Comision_Cartera_$": resumen_mes["Comision_Cartera_$"].sum(),
        "Bono_MiFi_$": resumen_mes["Bono_MiFi_$"].sum(),
        "Comision_Total_$": resumen_mes["Comision_Total_$"].sum()
    }])
    resumen_mes = pd.concat([resumen_mes, tot_row], ignore_index=True)

    # ----------------- HISTORIAL DE ACTIVACIONES (solo mes) -----------------
    hist_acts = altas_mes[["FECHA","DN_NORM","PLAN","COSTO PAQUETE"]].rename(columns={"DN_NORM":"DN"}).sort_values("FECHA")

    # ----------------- CARTERA MES (detalle recargas del mes) -----------------
    rec_det = rec_month_dist.copy()
    rec_det["DN"] = rec_det["DN_NORM"]
    rec_det["ELEGIBLE_MBB"] = rec_det["MONTO"] >= min_mbb
    rec_det = rec_det[["FECHA","DN","PLAN","MONTO","FORMA DE PAGO","ELEGIBLE_MBB"]].sort_values("FECHA")

    # Ajustes finales de redondeo
    for col in ["RECARGA_TOTAL_MES","COMISION_CARTERA_$","BONO_MIFI_50_$","COMISION_TOTAL_$",
                "Recarga_Mes_$","Comision_Cartera_$","Bono_MiFi_$","Comision_Total_$",
                "MONTO"]:
        if col in anexo.columns:
            anexo[col] = anexo[col].round(2)
        if col in resumen_mes.columns:
            resumen_mes[col] = resumen_mes[col].round(2)
        if col in rec_det.columns:
            rec_det[col] = rec_det[col].round(2)

    return {
        "RESUMEN": resumen,
        "ANEXO": anexo[[
            "DN","DN_NORM","FECHA","PLAN","COSTO PAQUETE","PRODUCTO",
            "RECARGA_TOTAL_MES","ELEGIBLE_CARTERA","% CARTERA APLICADA",
            "COMISION_CARTERA_$","BONO_MIFI_50_$","COMISION_TOTAL_$",
            "FirstRecargaFecha","DIAS_A_PRIMERA_REC"
        ]],
        "HISTORIAL DE ACTIVACIONES": hist_acts,
        f"RESUMEN {month_label_up}": resumen_mes,
        f"CARTERA {month_label_up}": rec_det
    }

def render_into_template(template_file, dataframes: dict, header_row_by_sheet: dict=None) -> BytesIO:
    """
    Abre la PLANTILLA con openpyxl y reemplaza el contenido de las hojas indicadas,
    respetando encabezados. Si header_row_by_sheet no se da, asume encabezado en fila 1.
    """
    wb = load_workbook(template_file)
    for sheet_name, df in dataframes.items():
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            # si no existe, la creamos (pero idealmente ya existen en la plantilla)
            ws = wb.create_sheet(title=sheet_name)
        header_row = 1
        if header_row_by_sheet and sheet_name in header_row_by_sheet:
            header_row = header_row_by_sheet[sheet_name]
        write_df_over_table(ws, df, header_row=header_row)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ----------------- UI -----------------
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial…)", type=["xlsx"])
    st.caption("Debe contener: 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")
with col2:
    template_file = st.file_uploader("PLANTILLA del distribuidor (Excel con todas las pestañas)", type=["xlsx"])
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.markdown("---")

hist_file = st.file_uploader("Opcional: **Histórico de recargas** del distribuidor (si existe en otro Excel). Si viene dentro de la base, puedes dejar vacío.", type=["xlsx"], help="Si no lo cargas, la app sólo usará 'Desgloce Recarga' de la base para buscar la primera recarga.")

if base_file and template_file and st.button("Generar reporte en PLANTILLA"):
    try:
        # Leer base
        xls = pd.ExcelFile(base_file, engine="openpyxl")
        if "Desgloce Totales" not in xls.sheet_names or "Desgloce Recarga" not in xls.sheet_names:
            st.error("El archivo base debe contener las hojas 'Desgloce Totales' y 'Desgloce Recarga'.")
        else:
            df_tot = pd.read_excel(base_file, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
            df_rec = pd.read_excel(base_file, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")

            # Histórico (si se sube aparte). Si no, usamos sólo df_rec para primera recarga.
            if hist_file:
                # Buscamos una hoja que tenga columnas compatibles; si no sabemos el nombre, tomamos la primera
                xh = pd.ExcelFile(hist_file, engine="openpyxl")
                # intenta detectar por columnas
                df_hist = None
                for sh in xh.sheet_names:
                    tmp = pd.read_excel(hist_file, sheet_name=sh, engine="openpyxl")
                    cols = {c.upper().strip() for c in tmp.columns.astype(str)}
                    if {"FECHA","DN","MONTO"}.issubset(cols):
                        df_hist = tmp
                        break
                if df_hist is None:
                    # fallback: primera hoja
                    df_hist = pd.read_excel(hist_file, sheet_name=0, engine="openpyxl")
            else:
                df_hist = pd.DataFrame(columns=df_rec.columns)

            result = calc_report(
                df_tot=df_tot,
                df_rec=df_rec,
                df_rec_hist=df_hist,
                dist_name=dist,
                year=int(year),
                month=int(month)
            )

            # Si tu plantilla tiene encabezados en otra fila, mapea aquí:
            header_rows = {
                # "RESUMEN": 1,
                # "ANEXO": 1,
                # "HISTORIAL DE ACTIVACIONES": 1,
                # f"RESUMEN {spanish_month_upper(pd.Timestamp(int(year),int(month),1))} {int(year)}": 1,
                # f"CARTERA {spanish_month_upper(pd.Timestamp(int(year),int(month),1))} {int(year)}": 1
            }

            output = render_into_template(template_file, result, header_row_by_sheet=header_rows)
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {spanish_month_upper(pd.Timestamp(int(year),int(month),1))} {int(year)}.xlsx"
            st.success("✅ Reporte generado en la PLANTILLA.")
            st.download_button("⬇️ Descargar Excel", data=output, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
else:
    st.info("Sube **Base mensual** y **PLANTILLA** del distribuidor, ajusta parámetros y presiona **Generar**.")
