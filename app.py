import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Valor Telecom - Comisiones", page_icon="📊", layout="wide")
st.title("📊 Generador de Comisiones | Valor Telecom")
st.caption("Sube la base mensual y el HISTORIAL del distribuidor. Calcula comisiones por esquema con ventanas M, M+1, M+2, M3–12 y exporta Excel final.")

# =========================
# Helpers
# =========================
def normalize_dn(series):
    out = series.astype(str).str.replace(r'\.0$', '', regex=True)
    def fix(x):
        try:
            if 'e+' in x.lower():
                return str(int(float(x)))
            return x.split('.')[0]
        except:
            return x
    return out.apply(fix)

def classify_producto(row):
    # Fallback por costo/identificadores cuando no hay ESQUEMA
    tipo = str(row.get('TIPO', '')).upper()
    costo = row.get('COSTO PAQUETE', np.nan)
    if 'MOB' in tipo:
        return 'MBB'
    if costo in [99, 115, 349, 399, 439, 500]:
        return 'HBB'
    if costo in [110, 120, 160, 245, 375, 480, 620]:
        return 'MiFi'
    return 'MBB'

def cartera_pct_mbb(n_altas_mes):
    # Ajuste por volumen del mes (no adicional): <50 = 3%, 50–299 = 5%, 300–499 = 7%, 500–999 = 8%, >=1000 = 10%
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

def ventana_m(rec_date, act_date):
    # Ventanas por días desde la activación (1–30=M, 31–60=M+1, 61–90=M+2, >=91=M3–12)
    if pd.isna(rec_date) or pd.isna(act_date):
        return None
    days = (rec_date - act_date).days + 1
    if days <= 0:
        return None
    if days <= 30:
        return "M"
    if days <= 60:
        return "M+1"
    if days <= 90:
        return "M+2"
    return "M3-12"

# =========================
# Reglas por esquema
# =========================
MIN_MBB  = 35
MIN_MIFI = 110
MIN_MIFI_10GB = 120
MIN_HBB  = 99

def tasa_por_esquema(esquema:str, ventana:str):
    # Devuelve porcentaje aplicable según esquema y ventana
    if not esquema:
        return None
    e = esquema.strip().lower()
    v = (ventana or "").upper()

    # MiFi con equipo (esquema 1): 10%/15%/10%/5%
    if e in ["mifi equipo", "mifi con equipo", "mifi (equipo)"]:
        return {"M":0.10, "M+1":0.15, "M+2":0.10, "M3-12":0.05}.get(v, 0.0)

    # MiFi Solo SIM: 5% M1–12
    if e in ["mifi solo sim", "mifi sim", "mifi (sim)"]:
        return 0.05

    # MiFi 10GB + 1a recarga: 5% M1–12  (el bono de $50 se trata aparte)
    if e in ["mifi 10gb + 1a recarga", "mifi 10gb", "mifi 10gb primera recarga"]:
        return 0.05

    # HBB con equipo: 5% M1–12
    if e in ["hbb equipo", "hbb con equipo", "hbb (equipo)"]:
        return 0.05

    # HBB Solo SIM: 5% M1–12
    if e in ["hbb solo sim", "hbb sim", "hbb (sim)"]:
        return 0.05

    # MBB: la tasa se define por volumen (se asigna fuera)
    if e == "mbb":
        return None

    # Si no coincide, None
    return None

def minimo_por_esquema(esquema:str):
    if not esquema:
        return None
    e = esquema.strip().lower()
    if e in ["mifi equipo", "mifi con equipo", "mifi (equipo)", "mifi solo sim", "mifi sim", "mifi (sim)"]:
        return MIN_MIFI
    if e in ["mifi 10gb + 1a recarga", "mifi 10gb", "mifi 10gb primera recarga"]:
        return MIN_MIFI_10GB
    if e in ["hbb equipo", "hbb con equipo", "hbb (equipo)", "hbb solo sim", "hbb sim", "hbb (sim)"]:
        return MIN_HBB
    if e == "mbb":
        return MIN_MBB
    return None

# =========================
# Cálculo de reporte
# =========================
def calc_report(df_tot, df_rec, df_hist_act=None, df_hist_rec=None, dist_name="ActivateCel", year=2025, month=6):
    month_start = pd.Timestamp(year, month, 1)
    month_end   = pd.Timestamp(year, month, 1) + pd.offsets.MonthEnd(1)

    # Normalización
    def prep(df):
        d = df.copy()
        if 'FECHA' in d.columns:
            d['FECHA'] = pd.to_datetime(d['FECHA'], errors='coerce')
        if 'DN' in d.columns:
            d['DN_NORM'] = normalize_dn(d['DN'])
        return d

    df_tot = prep(df_tot)
    df_rec = prep(df_rec)

    # Filtro distribuidor en totales
    mask_dist = df_tot['DISTRIBUIDOR '].astype(str).str.strip().str.lower() == dist_name.lower()
    tot_dist = df_tot[mask_dist].copy()

    # PRODUCTO base
    if 'PRODUCTO' not in tot_dist.columns:
        tot_dist['PRODUCTO'] = tot_dist.apply(classify_producto, axis=1)

    # Incorporar HISTORIAL (si viene)
    hist_act = prep(df_hist_act) if df_hist_act is not None else pd.DataFrame(columns=['DN','FECHA'])
    hist_rec = prep(df_hist_rec) if df_hist_rec is not None else pd.DataFrame(columns=['DN','FECHA','MONTO'])

    # Fecha de activación real por DN (mínima fecha entre historial y totales del dist)
    act_src = pd.concat([
        tot_dist[['DN_NORM','FECHA']],
        hist_act[['DN_NORM','FECHA']] if 'DN_NORM' in hist_act.columns else pd.DataFrame(columns=['DN_NORM','FECHA'])
    ], ignore_index=True).dropna(subset=['DN_NORM','FECHA'])

    act_dates = act_src.groupby('DN_NORM', as_index=False)['FECHA'].min().rename(columns={'FECHA':'FECHA_ACTIVACION'})

    # Universo DN del distribuidor
    dns_dist = set(tot_dist['DN_NORM'].dropna())

    # Recargas del mes (de DF base + historial si trae)
    rec_month_base = df_rec[(df_rec['FECHA']>=month_start) & (df_rec['FECHA']<=month_end)].copy()
    rec_month_hist = hist_rec[(hist_rec['FECHA']>=month_start) & (hist_rec['FECHA']<=month_end)].copy() if not hist_rec.empty else pd.DataFrame(columns=rec_month_base.columns)

    rec_month = pd.concat([rec_month_base, rec_month_hist], ignore_index=True)
    rec_month = rec_month[rec_month['DN_NORM'].isin(dns_dist)].copy()

    # Enriquecer con datos de totales (PLAN, COSTO, PRODUCTO, y ESQUEMA si existe en historial de activaciones)
    base_cols = [c for c in ['DN_NORM','PLAN','COSTO PAQUETE','PRODUCTO'] if c in tot_dist.columns]
    rec_month = rec_month.merge(tot_dist[base_cols].drop_duplicates('DN_NORM'), on='DN_NORM', how='left')

    # Traer ESQUEMA si el historial de activaciones lo trae
    if 'ESQUEMA' in hist_act.columns:
        esquemas = hist_act[['DN_NORM','ESQUEMA']].dropna().drop_duplicates('DN_NORM')
        rec_month = rec_month.merge(esquemas, on='DN_NORM', how='left')
    else:
        # Fallback: usar PRODUCTO como esquema base cuando no se especifica
        rec_month['ESQUEMA'] = rec_month.get('PRODUCTO', 'MBB')

    # Agregar fecha de activación
    rec_month = rec_month.merge(act_dates, on='DN_NORM', how='left')

    # Ventana M
    rec_month['VENTANA'] = rec_month.apply(lambda r: ventana_m(r['FECHA'], r['FECHA_ACTIVACION']), axis=1)

    # Altas del MES (para MBB % volumen -> contar DN con FECHA en mes_start..end en tot_dist)
    altas_mes = tot_dist[(tot_dist['FECHA']>=month_start) & (tot_dist['FECHA']<=month_end)].copy()
    n_altas_mes = altas_mes['DN_NORM'].nunique()
    pct_mbb = cartera_pct_mbb(n_altas_mes)

    # Cálculo de comisión por transacción de recarga
    def comision_row(row):
        monto = row.get('MONTO', 0.0) or 0.0
        esquema = str(row.get('ESQUEMA', '') or '').strip()
        producto = str(row.get('PRODUCTO', '') or '').strip()
        ventana = row.get('VENTANA', None)

        # Mínimo por esquema
        minimo = minimo_por_esquema(esquema)
        if minimo is None:
            # Si no hay esquema, usar mínimo por PRODUCTO
            minimo = MIN_MBB if producto.upper() == 'MBB' else (MIN_HBB if producto.upper() == 'HBB' else MIN_MIFI)

        if monto < minimo:
            return 0.0

        # MBB: tasa por volumen
        if esquema.lower() == 'mbb' or producto.upper() == 'MBB':
            return round(monto * pct_mbb, 2)

        # MiFi/HBB: tasa por esquema/ventana
        tasa = tasa_por_esquema(esquema, ventana)
        if tasa is None:
            # fallback 5% si no se pudo identificar el esquema concreto
            tasa = 0.05
        return round(monto * tasa, 2)

    rec_month['COMISION_TX'] = rec_month.apply(comision_row, axis=1)

    # Bono MiFi 10GB + 1a recarga (50 pesos si la primera recarga cae entre día 31–60)
    def bono_mifi_10gb(df):
        # detectar primera recarga por DN
        df_sorted = df.sort_values(['DN_NORM','FECHA'])
        first_rec = df_sorted.groupby('DN_NORM', as_index=False).first()
        cond_esquema = df_sorted['ESQUEMA'].str.strip().str.lower().isin(['mifi 10gb + 1a recarga','mifi 10gb','mifi 10gb primera recarga'])
        # calcular días de la primera recarga
        fr = df_sorted[cond_esquema].sort_values(['DN_NORM','FECHA']).groupby('DN_NORM', as_index=False).first()
        if fr.empty:
            return pd.DataFrame(columns=['DN_NORM','BONO_50'])
        fr['DIAS'] = (fr['FECHA'] - fr['FECHA_ACTIVACION']).dt.days + 1
        fr['BONO_50'] = np.where(fr['DIAS'].between(31,60, inclusive='both'), 50.0, 0.0)
        return fr[['DN_NORM','BONO_50']]

    bonos = bono_mifi_10gb(rec_month)
    if not bonos.empty:
        # bono por DN (no por transacción); lo sumamos al final por DN
        pass

    # ---- ANEXO (por línea en el mes) ----
    # Suma de recargas y comisiones por DN dentro del mes
    dn_agg = rec_month.groupby('DN_NORM', as_index=False).agg({
        'MONTO': 'sum',
        'COMISION_TX': 'sum'
    }).rename(columns={'MONTO':'RECARGA_TOTAL_MES','COMISION_TX':'COMISION_TOTAL_MES'})

    # Traer plan/costo/producto/esquema representativo por DN
    dn_info = rec_month.sort_values('FECHA').groupby('DN_NORM', as_index=False).last()[['DN_NORM','PLAN','COSTO PAQUETE','PRODUCTO','ESQUEMA']]
    anexo = dn_info.merge(dn_agg, on='DN_NORM', how='left').fillna({'RECARGA_TOTAL_MES':0.0,'COMISION_TOTAL_MES':0.0})

    # Añadir bono 50 (MiFi 10GB) por DN si corresponde
    if not bonos.empty:
        anexo = anexo.merge(bonos, on='DN_NORM', how='left')
        anexo['BONO_50'] = anexo['BONO_50'].fillna(0.0)
        anexo['COMISION_TOTAL_MES'] = (anexo['COMISION_TOTAL_MES'] + anexo['BONO_50']).round(2)
    else:
        anexo['BONO_50'] = 0.0

    # RESUMEN general
    resumen = pd.DataFrame([{
        'Distribuidor': dist_name,
        'Mes': f'{month_start.strftime("%B").capitalize()} {year}',
        'Altas del mes': int(n_altas_mes),
        'Recargas totales del mes ($)': round(rec_month['MONTO'].sum(), 2),
        'Porcentaje Cartera aplicado (MBB)': pct_mbb,
        'Comisión total del mes ($)': round(anexo['COMISION_TOTAL_MES'].sum(), 2)
    }])

    # RESUMEN MES (por ESQUEMA si existe, si no por PRODUCTO)
    agrupador = 'ESQUEMA' if 'ESQUEMA' in anexo.columns and anexo['ESQUEMA'].notna().any() else 'PRODUCTO'
    resumen_mes = (
        anexo.groupby(agrupador, as_index=False)
        .agg({
            'DN_NORM':'nunique',
            'RECARGA_TOTAL_MES':'sum',
            'COMISION_TOTAL_MES':'sum'
        })
        .rename(columns={
            'DN_NORM':'Lineas',
            'RECARGA_TOTAL_MES':'Recarga_Mes_$',
            'COMISION_TOTAL_MES':'Comision_Mes_$'
        })
    )
    total_row = pd.DataFrame([{
        agrupador: 'TOTAL',
        'Lineas': resumen_mes['Lineas'].sum(),
        'Recarga_Mes_$': resumen_mes['Recarga_Mes_$'].sum(),
        'Comision_Mes_$': resumen_mes['Comision_Mes_$'].sum()
    }])
    resumen_mes = pd.concat([resumen_mes, total_row], ignore_index=True)

    # HISTORIAL DE ACTIVACIONES (del distribuidor) -> todas las altas conocidas (historial + totales)
    hist_altas_dist = pd.concat([
        tot_dist[['FECHA','DN','DN_NORM','PLAN','COSTO PAQUETE']],
        hist_act[['FECHA','DN','DN_NORM','PLAN','COSTO PAQUETE']] if not hist_act.empty else pd.DataFrame(columns=['FECHA','DN','DN_NORM','PLAN','COSTO PAQUETE'])
    ], ignore_index=True).dropna(subset=['DN_NORM','FECHA']).drop_duplicates(subset=['DN_NORM','FECHA']).sort_values('FECHA')
    hist_altas_dist = hist_altas_dist.rename(columns={'DN_NORM':'DN_STD'})

    # CARTERA MES (detalle de recargas del mes)
    cartera_mes = rec_month[['FECHA','DN_NORM','PLAN','MONTO','FORMA DE PAGO','VENTANA','ESQUEMA','COMISION_TX']].rename(columns={'DN_NORM':'DN'}).sort_values('FECHA')

    # Export a Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
        anexo.to_excel(writer, sheet_name='ANEXO', index=False)
        hist_altas_dist.to_excel(writer, sheet_name='HISTORIAL DE ACTIVACIONES', index=False)
        resumen_mes.to_excel(writer, sheet_name=f'RESUMEN {month_start.strftime("%B").upper()} {year}', index=False)
        cartera_mes.to_excel(writer, sheet_name=f'CARTERA {month_start.strftime("%B").upper()} {year}', index=False)
    output.seek(0)
    return output

# =========================
# UI
# =========================
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader("1) Base mensual (VT Reporte Comercial…)", type=["xlsx"])
    st.caption("Debe tener 'Desgloce Totales' (header fila 2) y 'Desgloce Recarga' (header fila 4).")
with col2:
    hist_file = st.file_uploader("2) HISTORIAL del distribuidor (Excel ejemplo)", type=["xlsx"])
    st.caption("Usado para fecha de activación real, ESQUEMA y recargas históricas (si las trae).")

st.write("Parámetros")
dist = st.text_input("Distribuidor", value="ActivateCel")
year = st.number_input("Año", min_value=2023, max_value=2100, value=2025, step=1)
month = st.number_input("Mes (1–12)", min_value=1, max_value=12, value=6, step=1)

if st.button("Generar reporte"):
    try:
        if not base_file:
            st.error("Falta cargar la base mensual.")
        else:
            # Leer base mensual
            xls = pd.ExcelFile(base_file, engine="openpyxl")
            if 'Desgloce Totales' not in xls.sheet_names or 'Desgloce Recarga' not in xls.sheet_names:
                st.error("El archivo base debe contener 'Desgloce Totales' y 'Desgloce Recarga'.")
            else:
                df_tot = pd.read_excel(base_file, sheet_name='Desgloce Totales', header=1, engine="openpyxl")
                df_rec = pd.read_excel(base_file, sheet_name='Desgloce Recarga', header=3, engine="openpyxl")

                # Intentar leer historial si viene
                df_hist_act, df_hist_rec = None, None
                if hist_file:
                    xh = pd.ExcelFile(hist_file, engine="openpyxl")
                    # Heurística: buscar hojas que contengan 'HISTORIAL' y 'CARTERA' / 'RECARGA'
                    act_sheet = next((s for s in xh.sheet_names if 'HISTORIAL' in s.upper()), None)
                    rec_sheet = next((s for s in xh.sheet_names if 'CARTERA' in s.upper() or 'RECARGA' in s.upper()), None)

                    if act_sheet:
                        df_hist_act = pd.read_excel(hist_file, sheet_name=act_sheet, engine="openpyxl")
                    if rec_sheet:
                        df_hist_rec = pd.read_excel(hist_file, sheet_name=rec_sheet, engine="openpyxl")

                buf = calc_report(
                    df_tot=df_tot,
                    df_rec=df_rec,
                    df_hist_act=df_hist_act,
                    df_hist_rec=df_hist_rec,
                    dist_name=dist,
                    year=int(year),
                    month=int(month)
                )

                fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper()} {datetime(int(year), int(month), 1).strftime('%B').upper()} {year}.xlsx"
                st.success("Reporte generado.")
                st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
