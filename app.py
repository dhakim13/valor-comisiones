import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta

# ===================== CONFIG =====================
st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Carga la base mensual y la plantilla del distribuidor. Genera un Excel con RESUMEN, ANEXO, HISTORIAL DE ACTIVACIONES, RESUMEN {MES} y CARTERA {MES} reemplazando tablas y manteniendo encabezados.")

# ===================== HELPERS =====================
def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Regresa copia con columnas normalizadas (lower, strip, sin dobles espacios)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def first_existing_col(df: pd.DataFrame, candidates):
    """Devuelve el primer nombre de columna que exista en df (comparación exacta)."""
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(f"No encontré ninguna de las columnas {candidates} en: {list(df.columns)}")

def normalize_dn_series(s: pd.Series) -> pd.Series:
    """Normaliza DN a string numérica sin exponentes/decimales."""
    s = s.astype(str).str.strip()
    def fix(x):
        try:
            xl = x.lower()
            if 'e+' in xl or 'e-' in xl:
                return str(int(float(x)))
            if '.' in x:
                return x.split('.')[0]
            return ''.join(ch for ch in x if ch.isdigit())
        except:
            return x
    return s.apply(fix)

def month_span(year: int, month: int):
    start = pd.Timestamp(year, month, 1)
    end = start + pd.offsets.MonthEnd(1)
    return start, end

def cartera_pct_mbb(n_altas_mes: int) -> float:
    """Escalones MBB confirmados por el cliente."""
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

def classify_producto(tipo_val: str, costo_paquete):
    """Clasificación rápida si se requiere (ANEXO)."""
    tipo = str(tipo_val or "").upper()
    try:
        costo = float(costo_paquete)
    except:
        costo = np.nan

    if "MOB" in tipo:
        return "MBB"
    if costo in [99, 115, 349, 399, 439, 500]:
        return "HBB"
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return "MiFi"
    return "MBB"  # default

def ensure_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def find_sheet_like(xls: pd.ExcelFile, starts_with: str, default_name: str):
    """Busca hoja que comience con starts_with; si no existe, usa default_name."""
    for sn in xls.sheet_names:
        if sn.strip().upper().startswith(starts_with.upper()):
            return sn
    # Si no está en plantilla, devolvemos el nombre sugerido
    return default_name

# ===================== CÁLCULO =====================
def calc_report(
    xls_base: pd.ExcelFile,
    xls_tpl: pd.ExcelFile,
    year: int,
    month: int,
    dist_filtro: str
):
    # --- Rangos de mes
    month_start, month_end = month_span(year, month)
    mes_label_upper = month_start.strftime("%B").upper()
    mes_label_cap = month_start.strftime("%B").capitalize()

    # ----------- LECTURA: BASE (VT Reporte Comercial) -----------
    # Hoja totales
    if "Desgloce Totales" not in xls_base.sheet_names:
        raise ValueError("En la base mensual falta la hoja 'Desgloce Totales'.")

    df_tot = pd.read_excel(xls_base, sheet_name="Desgloce Totales", header=1, engine="openpyxl")
    df_tot = norm_cols(df_tot)

    # Columnas esperadas (con variantes conocidas)
    col_dist_candidates = ["DISTRIBUIDOR", "DISTRIBUIDOR "]
    col_fecha_alta_candidates = ["FECHA", "FECHA (ALTA)", "FECHA ALTA"]
    col_dn_candidates = ["DN", "MSISDN", "NUMERO"]
    col_plan_candidates = ["PLAN"]
    col_costo_candidates = ["COSTO PAQUETE", "PAQUETE", "COSTO"]

    col_dist_tot = first_existing_col(df_tot, [c for c in col_dist_candidates if c in df_tot.columns] or ["DISTRIBUIDOR "])
    col_fecha_alta = first_existing_col(df_tot, col_fecha_alta_candidates)
    col_dn_tot = first_existing_col(df_tot, col_dn_candidates)
    col_plan_tot = first_existing_col(df_tot, col_plan_candidates)
    col_costo_tot = first_existing_col(df_tot, col_costo_candidates)

    # Filtro distribuidor
    df_tot[col_dist_tot] = df_tot[col_dist_tot].astype(str).str.strip()
    mask_dist = df_tot[col_dist_tot].str.lower() == dist_filtro.lower().strip()
    tot_dist = df_tot[mask_dist].copy()
    if tot_dist.empty:
        raise ValueError(f"No hay registros en 'Desgloce Totales' para el distribuidor '{dist_filtro}'.")

    # Normalizaciones
    tot_dist["FECHA_ALTA"] = ensure_datetime(tot_dist[col_fecha_alta])
    tot_dist["DN_NORM"] = normalize_dn_series(tot_dist[col_dn_tot])
    tot_dist["PLAN_SRC"] = tot_dist[col_plan_tot].astype(str).str.strip()
    tot_dist["COSTO_SRC"] = tot_dist[col_costo_tot]

    # Intento de "TIPO" para ANEXO (si existe)
    col_tipo = "TIPO" if "TIPO" in tot_dist.columns else None
    tot_dist["PRODUCTO"] = tot_dist.apply(
        lambda r: classify_producto(r[col_tipo] if col_tipo else "", r["COSTO_SRC"]), axis=1
    )

    # Altas del mes (para métricas y %MBB)
    altas_mes = tot_dist[(tot_dist["FECHA_ALTA"] >= month_start) & (tot_dist["FECHA_ALTA"] <= month_end)].copy()
    n_altas_mes = int(altas_mes["DN_NORM"].nunique())

    # ----------- LECTURA: RECARGAS BASE -----------
    if "Desgloce Recarga" not in xls_base.sheet_names:
        raise ValueError("En la base mensual falta la hoja 'Desgloce Recarga'.")

    df_rec = pd.read_excel(xls_base, sheet_name="Desgloce Recarga", header=3, engine="openpyxl")
    df_rec = norm_cols(df_rec)

    # Columnas recarga (con variantes)
    col_fecha_rec_candidates = ["FECHA", "FECHA RECARGA"]
    col_dn_rec_candidates = ["DN", "MSISDN", "NUMERO"]
    col_monto_rec_candidates = ["MONTO", "IMPORTE", "MONTO RECARGA"]
    col_plan_rec_candidates = ["PLAN", "PAQUETE"]
    col_forma_pago_candidates = ["FORMA DE PAGO", "FORMA PAGO", "MEDIO DE PAGO"]

    col_fecha_rec = first_existing_col(df_rec, col_fecha_rec_candidates)
    col_dn_rec = first_existing_col(df_rec, col_dn_rec_candidates)
    col_monto_rec = first_existing_col(df_rec, col_monto_rec_candidates)
    col_plan_rec = None
    for c in col_plan_rec_candidates:
        if c in df_rec.columns:
            col_plan_rec = c
            break
    col_forma_pago = None
    for c in col_forma_pago_candidates:
        if c in df_rec.columns:
            col_forma_pago = c
            break

    df_rec["FECHA_REC"] = ensure_datetime(df_rec[col_fecha_rec])
    df_rec["DN_NORM"] = normalize_dn_series(df_rec[col_dn_rec])
    df_rec["MONTO"] = pd.to_numeric(df_rec[col_monto_rec], errors="coerce").fillna(0.0)

    # Universo de líneas del distribuidor
    dns_dist = set(tot_dist["DN_NORM"].dropna())
    rec_base_univ = df_rec[df_rec["DN_NORM"].isin(dns_dist)].copy()

    # Recargas del mes (base)
    rec_month_base = rec_base_univ[(df_rec["FECHA_REC"] >= month_start) & (df_rec["FECHA_REC"] <= month_end)].copy()

    # ----------- LECTURA: PLANTILLA (históricos) -----------
    # Usaremos HISTORIAL DE ACTIVACIONES para detectar portabilidad y estructura de columnas
    hist_sheet_name = "HISTORIAL DE ACTIVACIONES"
    if hist_sheet_name in xls_tpl.sheet_names:
        hist_tpl = pd.read_excel(xls_tpl, sheet_name=hist_sheet_name, header=0, engine="openpyxl")
        hist_tpl = norm_cols(hist_tpl)
    else:
        # Si no existe en plantilla, creamos estructura vacía con nombres estándar
        hist_tpl = pd.DataFrame(columns=["FECHA (alta)", "DN", "PLAN", "COSTO PAQUETE", "DN PORTADO"])

    # Detectar columnas en plantilla para clonar encabezados
    hist_cols = list(hist_tpl.columns)
    col_hist_fecha_alta = None
    for c in hist_cols:
        if c.strip().upper().startswith("FECHA"):
            col_hist_fecha_alta = c
            break
    if not col_hist_fecha_alta:
        col_hist_fecha_alta = "FECHA (alta)"

    col_hist_dn = None
    for cand in ["DN", "MSISDN", "NUMERO"]:
        if cand in hist_cols:
            col_hist_dn = cand
            break
    if not col_hist_dn:
        col_hist_dn = "DN"

    col_hist_plan = "PLAN" if "PLAN" in hist_cols else "PLAN"
    col_hist_costo = "COSTO PAQUETE" if "COSTO PAQUETE" in hist_cols else "COSTO PAQUETE"
    col_hist_portado = None
    for c in hist_cols:
        if "PORTAD" in c.upper():
            col_hist_portado = c
            break
    if not col_hist_portado:
        col_hist_portado = "DN PORTADO"  # columna nueva si no estaba

    # Unimos datos para HISTORIAL DE ACTIVACIONES (solo altas del mes)
    hist_out = (
        altas_mes[["FECHA_ALTA", "DN_NORM", "PLAN_SRC", "COSTO_SRC"]]
        .rename(columns={
            "FECHA_ALTA": col_hist_fecha_alta,
            "DN_NORM": col_hist_dn,
            "PLAN_SRC": col_hist_plan,
            "COSTO_SRC": col_hist_costo
        })
        .copy()
    )
    # Portabilidad: viene de plantilla original si se tuviera (no hay mapeo 1:1), por ahora en blanco
    hist_out[col_hist_portado] = ""  # se llenará si detectamos portados desde plantilla subida por el usuario

    # Si la plantilla traía valores de portado históricos y DN coinciden, intentamos arrastrarlos
    if not hist_tpl.empty and col_hist_portado in hist_tpl.columns and col_hist_dn in hist_tpl.columns:
        dn_portados = (
            hist_tpl[[col_hist_dn, col_hist_portado]]
            .dropna(subset=[col_hist_dn])
            .copy()
        )
        dn_portados[col_hist_dn] = normalize_dn_series(dn_portados[col_hist_dn])
        hist_out = hist_out.merge(
            dn_portados,
            on=col_hist_dn,
            how="left",
            suffixes=("", "_tpl")
        )
        # si viene algo en *_tpl preferimos ese valor
        hist_out[col_hist_portado] = hist_out[col_hist_portado].where(
            hist_out[col_hist_portado].astype(str).str.strip() != "",
            hist_out.get(f"{col_hist_portado}_tpl", "")
        )
        # limpiar columna auxiliar
        drop_aux = [c for c in hist_out.columns if c.endswith("_tpl")]
        if drop_aux:
            hist_out.drop(columns=drop_aux, inplace=True)

    hist_out = hist_out.sort_values(by=[col_hist_fecha_alta, col_hist_dn]).reset_index(drop=True)

    # ----------- CARTERA MES (detalle recargas) -----------
    # Estructura de columnas según plantilla (si existe)
    cartera_sheet_tpl = find_sheet_like(xls_tpl, "CARTERA ", f"CARTERA {mes_label_upper} {year}")
    if cartera_sheet_tpl in xls_tpl.sheet_names:
        cartera_tpl = pd.read_excel(xls_tpl, sheet_name=cartera_sheet_tpl, header=0, engine="openpyxl")
        cartera_tpl = norm_cols(cartera_tpl)
    else:
        cartera_tpl = pd.DataFrame(columns=["FECHA", "DN", "PLAN", "MONTO", "FORMA DE PAGO", "ELEGIBLE"])

    cartera_cols = list(cartera_tpl.columns)
    # Definimos nombres de salida tomando los headers visibles de la plantilla si existen
    cartera_col_fecha = None
    for c in cartera_cols:
        if c.strip().upper().startswith("FECHA"):
            cartera_col_fecha = c
            break
    if not cartera_col_fecha:
        cartera_col_fecha = "FECHA"

    cartera_col_dn = "DN" if "DN" in cartera_cols else "DN"
    cartera_col_plan = "PLAN" if "PLAN" in cartera_cols else (col_plan_rec or "PLAN")
    cartera_col_monto = None
    for c in cartera_cols:
        if "MONTO" in c.upper() or "IMPORTE" in c.upper():
            cartera_col_monto = c
            break
    if not cartera_col_monto:
        cartera_col_monto = "MONTO"

    cartera_col_forma = None
    for c in cartera_cols:
        if "PAGO" in c.upper():
            cartera_col_forma = c
            break
    if not cartera_col_forma:
        cartera_col_forma = "FORMA DE PAGO"

    cartera_col_elig = None
    for c in cartera_cols:
        if "ELEG" in c.upper():
            cartera_col_elig = c
            break
    if not cartera_col_elig:
        cartera_col_elig = "ELEGIBLE"

    # Construimos CARTERA {MES}
    rec_det = rec_month_base.copy()
    # Elegibilidad MBB por umbral de $35 (solamente para marcar; la comisión de cartera se calcula aparte con ventana M2)
    rec_det["ELEGIBLE_MBB"] = rec_det["MONTO"] >= 35

    cartera_mes_out = pd.DataFrame({
        cartera_col_fecha: ensure_datetime(rec_det["FECHA_REC"]),
        cartera_col_dn: rec_det["DN_NORM"],
        cartera_col_plan: (rec_det[col_plan_rec].astype(str).str.strip() if col_plan_rec else rec_det.get("PLAN", "").astype(str)),
        cartera_col_monto: rec_det["MONTO"],
        cartera_col_forma: (rec_det[col_forma_pago].astype(str).str.strip() if col_forma_pago else ""),
        cartera_col_elig: rec_det["ELEGIBLE_MBB"]
    }).sort_values(by=[cartera_col_fecha, cartera_col_dn]).reset_index(drop=True)

    # ----------- ANEXO (por línea, recarga total del mes y % aplicado informativo) -----------
    rec_by_dn = (
        rec_month_base.groupby("DN_NORM", as_index=False)["MONTO"].sum()
        .rename(columns={"MONTO": "RECARGA_TOTAL_MES"})
    )

    anexo = (
        tot_dist[[col_dn_tot, "DN_NORM", "FECHA_ALTA", "PLAN_SRC", "COSTO_SRC", "PRODUCTO"]]
        .merge(rec_by_dn, on="DN_NORM", how="left")
        .rename(columns={
            col_dn_tot: "DN",
            "FECHA_ALTA": "FECHA",
            "PLAN_SRC": "PLAN",
            "COSTO_SRC": "COSTO PAQUETE"
        })
        .copy()
    )
    anexo["RECARGA_TOTAL_MES"] = anexo["RECARGA_TOTAL_MES"].fillna(0.0)

    # % informativo (la cartera real se calcula sólo con recargas en ventana M2)
    pct_mbb_mes = cartera_pct_mbb(n_altas_mes)
    def pct_info(row):
        if row["PRODUCTO"] == "MBB":
            return pct_mbb_mes
        # si quieres, podríamos poner 0.05 para MiFi/HBB, pero el cliente pidió fórmula cartera con base M2 (en la práctica MBB).
        return 0.0
    anexo["% CARTERA (INFO)"] = anexo.apply(pct_info, axis=1)

    # ----------- CÁLCULOS DE COMISIONES -----------
    # 1) Cartera (M2): recargas en el MES cuya antigüedad respecto a FECHA_ALTA esté en [61, 90] días
    #    (M=1-30, M1=31-60, M2=61-90). Se aplica solo a MBB y a líneas del distribuidor.
    #    Usamos FECHA_ALTA de tot_dist para cada DN.
    altas_map = tot_dist[["DN_NORM", "FECHA_ALTA", "PRODUCTO"]].dropna().copy()
    rec_month_join = rec_month_base.merge(altas_map, on="DN_NORM", how="left")
    rec_month_join["DIAS_ALTA"] = (rec_month_join["FECHA_REC"] - rec_month_join["FECHA_ALTA"]).dt.days

    mask_m2 = (rec_month_join["DIAS_ALTA"] >= 61) & (rec_month_join["DIAS_ALTA"] <= 90)
    mask_mbb = (rec_month_join["PRODUCTO"] == "MBB")
    base_cartera_m2 = rec_month_join[mask_m2 & mask_mbb]["MONTO"].sum()

    comision_cartera = round(base_cartera_m2 * pct_mbb_mes, 2)

    # 2) Bono primera recarga ($15 por línea cuya primera recarga en la vida ocurra dentro del mes)
    #    Para identificar la "primera" necesitamos históricos. Tomamos:
    #    - todas las recargas del archivo plantilla que vengan en hojas "CARTERA *"
    #    - más la base del mes actual
    rec_hist_list = []
    for sh in xls_tpl.sheet_names:
        if sh.strip().upper().startswith("CARTERA"):
            tmp = pd.read_excel(xls_tpl, sheet_name=sh, header=0, engine="openpyxl")
            if tmp is not None and not tmp.empty:
                tmp = norm_cols(tmp)
                # Intento de mapear columnas mínimas
                try:
                    c_fecha = first_existing_col(tmp, [c for c in tmp.columns if str(c).strip().upper().startswith("FECHA")] or ["FECHA"])
                    c_dn = first_existing_col(tmp, ["DN", "MSISDN", "NUMERO"])
                    c_monto = None
                    for c in tmp.columns:
                        if "MONTO" in str(c).upper() or "IMPORTE" in str(c).upper():
                            c_monto = c
                            break
                    if c_dn and c_fecha:
                        cur = pd.DataFrame({
                            "FECHA_REC": ensure_datetime(tmp[c_fecha]),
                            "DN_NORM": normalize_dn_series(tmp[c_dn]),
                            "MONTO": pd.to_numeric(tmp[c_monto], errors="coerce") if c_monto else 0.0
                        })
                        rec_hist_list.append(cur)
                except Exception:
                    # si una hoja no cuadra, la omitimos
                    pass

    rec_hist_total = pd.concat(rec_hist_list + [rec_base_univ[["FECHA_REC", "DN_NORM", "MONTO"]]], ignore_index=True) if rec_hist_list else rec_base_univ[["FECHA_REC", "DN_NORM", "MONTO"]].copy()
    if not rec_hist_total.empty:
        first_rec = rec_hist_total.sort_values(["DN_NORM", "FECHA_REC"]).dropna(subset=["DN_NORM", "FECHA_REC"]).groupby("DN_NORM", as_index=False).agg({"FECHA_REC": "min"})
        # Líneas cuya primera recarga cayó en el mes
        first_in_month = first_rec[(first_rec["FECHA_REC"] >= month_start) & (first_rec["FECHA_REC"] <= month_end)]
        bono_primera_recarga = 15 * int(first_in_month["DN_NORM"].nunique())
    else:
        bono_primera_recarga = 0

    # 3) Portabilidad ($30 por alta portada en el mes)
    #    Se detecta desde HISTORIAL: columna tipo "DN PORTADO" no vacía.
    #    Si la plantilla traía esa columna, la usamos; si no, la de hist_out está vacía (=> 0).
    hist_mes_for_port = hist_out[(ensure_datetime(hist_out[col_hist_fecha_alta]) >= month_start) & (ensure_datetime(hist_out[col_hist_fecha_alta]) <= month_end)].copy()
    if col_hist_portado in hist_mes_for_port.columns:
        portadas_mes = hist_mes_for_port[col_hist_portado].astype(str).str.strip()
        n_portadas = int((portadas_mes != "").sum())
    else:
        n_portadas = 0
    comision_portabilidad = 30 * n_portadas

    # ----------- RESUMEN (estructura clonada de plantilla si existe) -----------
    # Intentamos leer RESUMEN de la plantilla para respetar encabezados
    resumen_tpl_name = "RESUMEN" if "RESUMEN" in xls_tpl.sheet_names else "RESUMEN"
    if resumen_tpl_name in xls_tpl.sheet_names:
        resumen_tpl = pd.read_excel(xls_tpl, sheet_name=resumen_tpl_name, header=0, engine="openpyxl")
        resumen_tpl = norm_cols(resumen_tpl)
        resumen_cols = list(resumen_tpl.columns)
    else:
        resumen_cols = ["Concepto", "Valor"]

    total_cartera_mes = round(rec_month_base["MONTO"].sum(), 2)

    resumen_rows = [
        ("Distribuidor", dist_filtro),
        ("Mes", f"{mes_label_cap} {year}"),
        ("Altas del mes", n_altas_mes),
        ("% Cartera MBB (por altas del mes)", pct_mbb_mes),
        ("Base Cartera (M2, $)", round(base_cartera_m2, 2)),
        ("Comisión Cartera ($)", comision_cartera),
        ("# Líneas con 1ra recarga en el mes", int(bono_primera_recarga / 15) if bono_primera_recarga else 0),
        ("Comisión 1ra recarga ($)", bono_primera_recarga),
        ("# Altas portadas del mes", n_portadas),
        ("Comisión Portabilidad ($)", comision_portabilidad),
        ("Recargas totales del mes ($)", total_cartera_mes),
        ("Comisión Total del mes ($)", comision_cartera + bono_primera_recarga + comision_portabilidad)
    ]
    # Construimos DataFrame RESUMEN con dos columnas
    resumen_df = pd.DataFrame(resumen_rows, columns=[resumen_cols[0] if len(resumen_cols) > 0 else "Concepto",
                                                     resumen_cols[1] if len(resumen_cols) > 1 else "Valor"])

    # ----------- RESUMEN {MES} (agregado por PRODUCTO) -----------
    # Usamos ANEXO como base (lineas del distribuidor + sumas de recarga del mes).
    resumen_mes = (
        anexo.groupby("PRODUCTO", as_index=False)
             .agg({
                 "DN_NORM": "nunique",
                 "RECARGA_TOTAL_MES": "sum"
             })
             .rename(columns={
                 "DN_NORM": "Lineas",
                 "RECARGA_TOTAL_MES": "Recarga_Mes_$"
             })
    )
    # Comisión por producto (sólo MBB con regla de cartera M2 del mes). Para mostrar algo, prorrateamos por peso de recarga.
    total_recargas_anexo = resumen_mes["Recarga_Mes_$"].sum()
    if total_recargas_anexo > 0:
        resumen_mes["Comision_Mes_$"] = resumen_mes["Recarga_Mes_$"] / total_recargas_anexo * (comision_cartera + bono_primera_recarga + comision_portabilidad)
    else:
        resumen_mes["Comision_Mes_$"] = 0.0
    resumen_mes = resumen_mes.round({"Recarga_Mes_$": 2, "Comision_Mes_$": 2})

    total_row = pd.DataFrame([{
        "PRODUCTO": "TOTAL",
        "Lineas": resumen_mes["Lineas"].sum(),
        "Recarga_Mes_$": resumen_mes["Recarga_Mes_$"].sum(),
        "Comision_Mes_$": resumen_mes["Comision_Mes_$"].sum()
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # ----------- NOMBRE DE HOJAS RESUMEN/CARTERA SEGÚN PLANTILLA -----------
    resumen_mes_sheet = find_sheet_like(xls_tpl, "RESUMEN ", f"RESUMEN {mes_label_upper} {year}")
    cartera_mes_sheet = find_sheet_like(xls_tpl, "CARTERA ", f"CARTERA {mes_label_upper} {year}")

    # ----------- EXPORTAR: copiar nombres de hojas y reemplazar datos -----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # RESUMEN
        resumen_df.to_excel(writer, sheet_name="RESUMEN", index=False)

        # ANEXO
        anexo_out = anexo.rename(columns={
            "FECHA": "FECHA",
            "PLAN": "PLAN",
            "COSTO PAQUETE": "COSTO PAQUETE",
            "RECARGA_TOTAL_MES": "RECARGA_TOTAL_MES",
            "PRODUCTO": "PRODUCTO",
            "DN": "DN"
        })
        anexo_out.to_excel(writer, sheet_name="ANEXO", index=False)

        # HISTORIAL DE ACTIVACIONES (encabezados clonados)
        # Aseguramos el orden de columnas de salida respetando encabezados de la plantilla
        desired_hist_cols = []
        for c in [col_hist_fecha_alta, col_hist_dn, col_hist_plan, col_hist_costo, col_hist_portado]:
            if c not in desired_hist_cols:
                desired_hist_cols.append(c)
        # Si la plantilla trae más columnas, las agregamos al final en blanco
        for c in hist_cols:
            if c not in desired_hist_cols:
                desired_hist_cols.append(c)
                if c not in hist_out.columns:
                    hist_out[c] = ""
        hist_out[desired_hist_cols].to_excel(writer, sheet_name="HISTORIAL DE ACTIVACIONES", index=False)

        # RESUMEN {MES}
        resumen_mes.to_excel(writer, sheet_name=resumen_mes_sheet, index=False)

        # CARTERA {MES} (encabezados clonados)
        desired_cart_cols = [cartera_col_fecha, cartera_col_dn, cartera_col_plan, cartera_col_monto, cartera_col_forma, cartera_col_elig]
        for c in cartera_cols:
            if c not in desired_cart_cols:
                desired_cart_cols.append(c)
                if c not in cartera_mes_out.columns:
                    cartera_mes_out[c] = ""
        cartera_mes_out[desired_cart_cols].to_excel(writer, sheet_name=cartera_mes_sheet, index=False)

    output.seek(0)
    fname = f"COMISION VALOR DISTRIBUIDOR {dist_filtro.upper()} {mes_label_upper} {year}.xlsx"
    return output, fname

# ===================== UI =====================
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("Base mensual (VT Reporte Comercial…)", type=["xlsx"])
    st.caption("Debe traer: 'Desgloce Totales' (header en fila 2) y 'Desgloce Recarga' (header en fila 4).")
with col2:
    tpl_file = st.file_uploader("Plantilla del distribuidor (archivo ejemplo)", type=["xlsx"])
    dist = st.text_input("Distribuidor", value="ActivateCel")
    year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
    month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

st.markdown("---")
if base_file and tpl_file and st.button("Generar reporte"):
    try:
        xls_base = pd.ExcelFile(base_file, engine="openpyxl")
        xls_tpl = pd.ExcelFile(tpl_file, engine="openpyxl")

        buf, fname = calc_report(
            xls_base=xls_base,
            xls_tpl=xls_tpl,
            year=int(year),
            month=int(month),
            dist_filtro=dist.strip()
        )
        st.success("✅ Reporte generado.")
        st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
else:
    st.info("Sube **ambos** archivos, confirma distribuidor/mes y da clic en **Generar reporte**.")

