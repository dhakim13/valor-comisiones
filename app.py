import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# =========================
# Configuración de página
# =========================
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual (VT Reporte Comercial…) y, opcionalmente, el archivo del distribuidor (historial). Exporta un Excel con RESUMEN, ANEXO, HISTORIAL, RESUMEN MES y CARTERA MES con formato.")

# =========================
# Helpers
# =========================
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
    """
    Clasificación:
    - Si TIPO contiene 'MOB' => MBB
    - Por costo de paquete inferimos HBB/MiFi (ajustable).
    """
    tipo = str(row.get("TIPO", "")).upper()
    costo = row.get("COSTO PAQUETE", np.nan)
    if "MOB" in tipo:
        return "MBB"
    # HBB (ejemplos de precios del tablóide)
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    # MiFi (ejemplos)
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    # Default al principal
    return "MBB"

def cartera_pct_mbb(n_altas_mes: int) -> float:
    """
    MBB por volumen mensual de altas:
      <50 -> 3%
      50–299 -> 5%
      300–499 -> 7%
      500–999 -> 8%
      1000+ -> 10%
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

# =========================
# Estilos XLSXWriter
# =========================
def best_width(series, minw=10, maxw=40):
    try:
        lens = series.astype(str).str.len()
        return int(min(max(lens.max() + 2, minw), maxw))
    except Exception:
        return minw

def apply_common_formats(writer):
    wb = writer.book
    fmt_title = wb.add_format({
        "bold": True, "font_size": 14, "align": "left", "valign": "vcenter"
    })
    fmt_subtitle = wb.add_format({
        "bold": True, "font_size": 10, "font_color": "#666666"
    })
    fmt_header = wb.add_format({
        "bold": True, "bg_color": "#F2F2F2", "border": 1,
        "align": "center", "valign": "vcenter"
    })
    fmt_cell = wb.add_format({"border": 1})
    fmt_money = wb.add_format({"num_format": "$#,##0.00", "border": 1})
    fmt_percent = wb.add_format({"num_format": "0.00%", "border": 1})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    fmt_total = wb.add_format({"bold": True, "border": 1, "bg_color": "#FFF2CC"})
    fmt_note = wb.add_format({"font_color": "#404040", "italic": True})
    return {
        "title": fmt_title, "subtitle": fmt_subtitle, "header": fmt_header,
        "cell": fmt_cell, "money": fmt_money, "percent": fmt_percent,
        "date": fmt_date, "total": fmt_total, "note": fmt_note
    }

def write_df_sheet(
    writer, df, sheet_name,
    money_cols=None, percent_cols=None, date_cols=None,
    freeze_panes=(1,0), add_filter=True, add_total_row=False, total_label="TOTAL"
):
    """Escribe un DataFrame con encabezado estilizado, bordes, anchos y formatos."""
    if money_cols is None: money_cols = []
    if percent_cols is None: percent_cols = []
    if date_cols is None: date_cols = []

    formats = apply_common_formats(writer)
    df_to_write = df.copy()

    # Escribir datos crudos sin índices
    df_to_write.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
    ws = writer.sheets[sheet_name]
    wb = writer.book

    # Encabezados estilizados
    for col_idx, col_name in enumerate(df_to_write.columns):
        ws.write(0, col_idx, col_name, formats["header"])

    # Bordes/formato por columnas + ancho
    for col_idx, col_name in enumerate(df_to_write.columns):
        col_data = df_to_write[col_name]
        width = best_width(col_data)
        ws.set_column(col_idx, col_idx, width)

        # Detecta rango de celdas con datos
        nrows = len(df_to_write)
        if nrows > 0:
            rng = (1, col_idx, nrows, col_idx)  # (row_first, col_first, row_last, col_last)
            col_format = formats["cell"]
            if col_name in money_cols:
                col_format = formats["money"]
            elif col_name in percent_cols:
                col_format = formats["percent"]
            elif col_name in date_cols:
                col_format = formats["date"]
            ws.conditional_format(rng[0], rng[1], rng[2], rng[3], {"type": "no_errors", "format": col_format})

    # Filtros
    if add_filter and len(df_to_write.columns) > 0:
        ws.autofilter(0, 0, len(df_to_write), len(df_to_write.columns)-1)

    # Panes congelados
    if freeze_panes:
        ws.freeze_panes(*freeze_panes)

    # Fila de totales
    if add_total_row and len(df_to_write) > 0:
        last_row = len(df_to_write) + 1  # por startrow=1
        ws.write(last_row, 0, total_label, formats["total"])
        for col_idx, col_name in enumerate(df_to_write.columns[1:], start=1):
            # suma si es numérica
            if pd.api.types.is_numeric_dtype(df_to_write[col_name]):
                col_letter = xlsx_col_letter(col_idx)
                ws.write_formula(last_row, col_idx, f"=SUBTOTAL(9,{col_letter}2:{col_letter}{last_row})", formats["total"])

def xlsx_col_letter(idx):
    # 0 -> A, 1 -> B, ...
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx-1, 26)
        letters = chr(65+rem) + letters
    return letters

# =========================
# Núcleo de cálculo
# =========================
def calc_report(
    df_tot: pd.DataFrame,
    df_rec: pd.DataFrame,
    dist_name: str,
    year: int,
    month: int,
    hist_activ: pd.DataFrame | None = None,
    hist_rec: pd.DataFrame | None = None,
) -> BytesIO:

    month_start = pd.Timestamp(year, month, 1)
    month_end   = month_start + pd.offsets.MonthEnd(1)

    # --------- Normalización y llaves ---------
    df_tot = df_tot.copy()
    df_rec = df_rec.copy()

    # Fechas
    if "FECHA" in df_tot.columns:
        df_tot["FECHA"] = pd.to_datetime(df_tot["FECHA"], errors="coerce")
    if "FECHA" in df_rec.columns:
        df_rec["FECHA"] = pd.to_datetime(df_rec["FECHA"], errors="coerce")

    # DN normalizado
    if "DN" not in df_tot.columns:
        raise ValueError("En 'Desgloce Totales' falta la columna 'DN'.")
    if "DN" not in df_rec.columns:
        raise ValueError("En 'Desgloce Recarga' falta la columna 'DN'.")

    df_tot["DN_NORM"] = normalize_dn(df_tot["DN"])
    df_rec["DN_NORM"] = normalize_dn(df_rec["DN"])

    # --------- Filtro por distribuidor ---------
    col_dist = "DISTRIBUIDOR " if "DISTRIBUIDOR " in df_tot.columns else "DISTRIBUIDOR"
    if col_dist not in df_tot.columns:
        raise ValueError("No encuentro la columna de distribuidor en 'Desgloce Totales'.")

    mask_dist = df_tot[col_dist].astype(str).str.strip().str.lower() == dist_name.strip().lower()
    tot_dist = df_tot[mask_dist].copy()
    dns_dist = set(tot_dist["DN_NORM"].dropna()) if not tot_dist.empty else set()

    # --------- Activaciones y Recargas del mes ---------
    altas_mes = (
        tot_dist[(tot_dist.get("FECHA") >= month_start) & (tot_dist.get("FECHA") <= month_end)].copy()
        if "FECHA" in tot_dist.columns else tot_dist.iloc[0:0].copy()
    )

    rec_month_base = (
        df_rec[(df_rec.get("FECHA") >= month_start) & (df_rec.get("FECHA") <= month_end)].copy()
        if "FECHA" in df_rec.columns else df_rec.iloc[0:0].copy()
    )
    rec_month_dist = rec_month_base[rec_month_base["DN_NORM"].isin(dns_dist)].copy()

    # --------- Clasificación de producto ---------
    if not tot_dist.empty:
        tot_dist["PRODUCTO"] = tot_dist.apply(classify_row, axis=1)
    else:
        tot_dist["PRODUCTO"] = []

    # --------- Reglas de negocio ---------
    n_altas = altas_mes["DN_NORM"].nunique() if not altas_mes.empty else 0
    pct_mbb = cartera_pct_mbb(n_altas)
    min_mbb = 35
    min_mifi = 110
    min_hbb = 99

    # Suma de recargas del mes por DN
    rec_by_dn = (
        rec_month_dist.groupby("DN_NORM", as_index=False)["MONTO"]
        .sum()
        .rename(columns={"MONTO": "RECARGA_TOTAL_MES"})
    )

    # --------- ANEXO (detalle por línea del universo del distribuidor) ---------
    cols_keep = [c for c in ["DN", "DN_NORM", "FECHA", "PLAN", "COSTO PAQUETE", "PRODUCTO"] if c in tot_dist.columns]
    if "DN_NORM" not in cols_keep:
        cols_keep.insert(1, "DN_NORM")
    anexo = tot_dist[cols_keep].merge(rec_by_dn, on="DN_NORM", how="left")
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # Elegibilidad cartera por producto
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
        elif row["PRODUCTO"] in ("MiFi", "HBB"):
            return 0.05
        return 0.0

    anexo["% CARTERA APLICADA"] = anexo.apply(pct_aplicado, axis=1)
    anexo["COMISION_CARTERA_$"] = np.where(
        anexo["ELEGIBLE_CARTERA"],
        anexo["RECARGA_TOTAL_MES"] * anexo["% CARTERA APLICADA"],
        0.0
    ).round(2)

    # --------- RESUMEN (portada) ---------
    rec_total_mes = float(rec_month_dist["MONTO"].sum()) if "MONTO" in rec_month_dist.columns else 0.0
    com_total = float(anexo["COMISION_CARTERA_$"].sum()) if not anexo.empty else 0.0

    resumen = pd.DataFrame([{
        "Distribuidor": dist_name,
        "Mes": f"{month_start.strftime('%B').capitalize()} {year}",
        "Altas del mes": int(n_altas),
        "Recargas totales del mes ($)": round(rec_total_mes, 2),
        "Porcentaje Cartera aplicado (MBB)": pct_mbb,
        "Comisión Cartera total ($)": round(com_total, 2)
    }])

    # --------- RESUMEN MES (agregado por producto + total)
    if not anexo.empty:
        resumen_mes = (
            anexo.groupby("PRODUCTO", as_index=False)
            .agg(**{
                "Lineas": ("DN_NORM", "nunique"),
                "Recarga_Mes_$": ("RECARGA_TOTAL_MES", "sum"),
                "Comision_Cartera_$": ("COMISION_CARTERA_$", "sum"),
            })
        )
    else:
        resumen_mes = pd.DataFrame(columns=["PRODUCTO", "Lineas", "Recarga_Mes_$", "Comision_Cartera_$"])

    total_row = pd.DataFrame([{
        "PRODUCTO": "TOTAL",
        "Lineas": resumen_mes["Lineas"].sum() if not resumen_mes.empty else 0,
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum() if not resumen_mes.empty else 0.0,
        "Comision_Cartera_$": resumen_mes["Comision_Cartera_$"].sum() if not resumen_mes.empty else 0.0
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # --------- HISTORIAL DE ACTIVACIONES (del mes) ---------
    if hist_activ is not None and not hist_activ.empty:
        hist = hist_activ.copy()
        if "FECHA" in hist.columns:
            hist["FECHA"] = pd.to_datetime(hist["FECHA"], errors="coerce")
            hist = hist[(hist["FECHA"] >= month_start) & (hist["FECHA"] <= month_end)].copy()
        if "FECHA" in hist.columns:
            hist = hist.sort_values("FECHA")
        # Normaliza nombre de DN si viniera como 'DN_NORM'
        if "DN_NORM" in hist.columns and "DN" not in hist.columns:
            hist = hist.rename(columns={"DN_NORM": "DN"})
    else:
        base_cols = []
        if "FECHA" in altas_mes.columns: base_cols.append("FECHA")
        if "DN_NORM" in altas_mes.columns: base_cols.append("DN_NORM")
        if "PLAN" in altas_mes.columns: base_cols.append("PLAN")
        if "COSTO PAQUETE" in altas_mes.columns: base_cols.append("COSTO PAQUETE")
        hist = altas_mes[base_cols].rename(columns={"DN_NORM": "DN"}) if base_cols else pd.DataFrame(columns=["FECHA","DN","PLAN","COSTO PAQUETE"])
        if "FECHA" in hist.columns:
            hist = hist.sort_values("FECHA")

    # --------- CARTERA {MES} (detalle recargas del mes)
    rec_det = rec_month_dist.copy()
    if not rec_det.empty:
        rec_det["ELEGIBLE_MBB"] = rec_det["MONTO"] >= min_mbb if "MONTO" in rec_det.columns else False
        keep = [c for c in ["FECHA", "DN_NORM", "PLAN", "MONTO", "FORMA DE PAGO"] if c in rec_det.columns]
        if "DN_NORM" in keep:
            rec_det = rec_det[keep].rename(columns={"DN_NORM": "DN"})
        else:
            rec_det = rec_det[keep]
        if "FECHA" in rec_det.columns:
            rec_det = rec_det.sort_values("FECHA")
    else:
        rec_det = pd.DataFrame(columns=["FECHA", "DN", "PLAN", "MONTO", "FORMA DE PAGO", "ELEGIBLE_MBB"])

    # =========================
    # Exportar a Excel (memoria) con estilos
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        fmts = apply_common_formats(writer)

        # 1) RESUMEN (portada tipo cuadro)
        sheet_name_resumen = "RESUMEN"
        resumen.to_excel(writer, sheet_name=sheet_name_resumen, index=False, startrow=5)
        ws = writer.sheets[sheet_name_resumen]

        # Título
        title = f"COMISIONES {dist_name.upper()} — {month_start.strftime('%B').upper()} {year}"
        ws.write(0, 0, title, fmts["title"])
        ws.write(2, 0, "Este reporte incluye: ANEXO (detalle líneas), HISTORIAL DE ACTIVACIONES, RESUMEN y CARTERA del mes.", fmts["subtitle"])

        # Dar formato a columnas
        # Moneda/porcentaje por nombre de columna
        if "Recargas totales del mes ($)" in resumen.columns:
            col_idx = resumen.columns.get_loc("Recargas totales del mes ($)")
            ws.set_column(col_idx, col_idx, 20, fmts["money"])
        if "Comisión Cartera total ($)" in resumen.columns:
            col_idx = resumen.columns.get_loc("Comisión Cartera total ($)")
            ws.set_column(col_idx, col_idx, 20, fmts["money"])
        if "Porcentaje Cartera aplicado (MBB)" in resumen.columns:
            col_idx = resumen.columns.get_loc("Porcentaje Cartera aplicado (MBB)")
            ws.set_column(col_idx, col_idx, 15, fmts["percent"])

        # Encabezado tipo header
        for col_idx, col_name in enumerate(resumen.columns):
            ws.write(5, col_idx, col_name, fmts["header"])
            ws.set_column(col_idx, col_idx, max(15, best_width(resumen[col_name])))

        # 2) ANEXO
        money_cols_anexo = ["RECARGA_TOTAL_MES", "COMISION_CARTERA_$"]
        percent_cols_anexo = ["% CARTERA APLICADA"]
        date_cols_anexo = ["FECHA"] if "FECHA" in anexo.columns else []
        write_df_sheet(
            writer, anexo, "ANEXO",
            money_cols=money_cols_anexo,
            percent_cols=percent_cols_anexo,
            date_cols=date_cols_anexo,
            freeze_panes=(1,0),
            add_filter=True,
            add_total_row=True,
            total_label="TOTAL"
        )

        # 3) HISTORIAL DE ACTIVACIONES
        date_cols_hist = ["FECHA"] if "FECHA" in hist.columns else []
        write_df_sheet(
            writer, hist, "HISTORIAL DE ACTIVACIONES",
            date_cols=date_cols_hist,
            add_filter=True,
            freeze_panes=(1,0),
            add_total_row=False
        )

        # 4) RESUMEN {MES}
        sheet_name_resumen_mes = f"RESUMEN {month_start.strftime('%B').upper()} {year}"
        money_cols_resumen_mes = ["Recarga_Mes_$", "Comision_Cartera_$"]
        write_df_sheet(
            writer, resumen_mes, sheet_name_resumen_mes,
            money_cols=money_cols_resumen_mes,
            add_filter=True,
            freeze_panes=(1,0),
            add_total_row=False
        )

        # 5) CARTERA {MES}
        sheet_name_cartera_mes = f"CARTERA {month_start.strftime('%B').upper()} {year}"
        date_cols_cartera = ["FECHA"] if "FECHA" in rec_det.columns else []
        money_cols_cartera = ["MONTO"] if "MONTO" in rec_det.columns else []
        write_df_sheet(
            writer, rec_det, sheet_name_cartera_mes,
            date_cols=date_cols_cartera,
            money_cols=money_cols_cartera,
            add_filter=True,
            freeze_panes=(1,0),
            add_total_row=True,
            total_label="TOTAL"
        )

    output.seek(0)
    return output

# =========================
# UI
# =========================
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial…)", type=["xlsx"], key="base")
    st.caption("Debe incluir: 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")
    hist_file = st.file_uploader("Archivo del distribuidor (historial) - opcional", type=["xlsx"], key="hist")
    st.caption("Si lo subes, usaré sus hojas para HISTORIAL; si no, lo construyo desde la base.")

with col2:
    st.write("Parámetros")
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

if base_file and st.button("Generar reporte"):
    try:
        # Leer base (con encabezados en filas específicas)
        xls_base = pd.ExcelFile(base_file, engine="openpyxl")
        if ("Desgloce Totales" not in xls_base.sheet_names) or ("Desgloce Recarga" not in xls_base.sheet_names):
            st.error("El archivo base debe contener las hojas 'Desgloce Totales' y 'Desgloce Recarga'.")
        else:
            df_tot = pd.read_excel(base_file, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
            df_rec = pd.read_excel(base_file, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")

            # Historial (opcional)
            hist_activ = None
            hist_rec = None
            if hist_file is not None:
                try:
                    xls_hist = pd.ExcelFile(hist_file, engine="openpyxl")
                    if "HISTORIAL DE ACTIVACIONES" in xls_hist.sheet_names:
                        hist_activ = pd.read_excel(hist_file, sheet_name="HISTORIAL DE ACTIVACIONES", engine="openpyxl")
                    # Si existiera hoja de recargas histórica
                    for cand in ["DETALLE RECARGAS", "CARTERA", f"CARTERA {datetime(int(year), int(month), 1).strftime('%B').upper()} {int(year)}"]:
                        if cand in xls_hist.sheet_names:
                            hist_rec = pd.read_excel(hist_file, sheet_name=cand, engine="openpyxl")
                            break
                except Exception:
                    hist_activ, hist_rec = None, None

            buf = calc_report(
                df_tot=df_tot,
                df_rec=df_rec,
                dist_name=dist,
                year=int(year),
                month=int(month),
                hist_activ=hist_activ,
                hist_rec=hist_rec
            )
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {int(year)}.xlsx"
            st.success("✅ Reporte generado.")
            st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.exception(e)
