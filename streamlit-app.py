# Activar el entorno .\.venv\Scripts\Activate.ps1
# Ejecutar en navegador streamlit run streamlit-app.py
import requests
import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter
from supabase import create_client, Client
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# Configuracion de la pagina
st.set_page_config(page_title="Mi Portafolio", layout="wide")

# Buscar las claves en secrets.
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
banxico_token = st.secrets["BANXICO_TOKEN"]
app_password = st.secrets["APP_PASSWORD"]
supabase: Client = create_client(url, key)

# Variables globales
BROKERS = ["Fintual", "Banco Plata", "ARQ"]

# ==========
# Funciones.
# ==========

# Funcion para verificar el password.
def verificar_password():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if not st.session_state["autenticado"]:
        st.title("🔒 Acceder al Portafolio")
        input_pass = st.text_input("Ingresa la contraseña para acceder:", type="password")
        if st.button("Ingresar"):
            if input_pass == app_password:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta. Acceso denegado.")
        return False
    return True

# Detener la ejecución si no está autenticado
if not verificar_password():
    st.stop()

# Consulta a Banxico para obtenr tasa de cambio del DOF.
@st.cache_data(ttl=86400, show_spinner="Obteniendo tasa de cambio del DOF...")
def obtener_tc_dof(fecha_str):
    if not banxico_token:
        raise RuntimeError("Falta BANXICO_TOKEN en secrets.")
        
    headers = {"Bmx-Token": banxico_token}
    fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d")
    fecha_inicio = (fecha_dt - timedelta(days=7)).strftime("%Y-%m-%d")
    fecha_fin = (fecha_dt - timedelta(days=1)).strftime("%Y-%m-%d")
    url_bmx = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/SF43718/datos/{fecha_inicio}/{fecha_fin}"

    response = requests.get(url_bmx, headers=headers, timeout=10)
    if response.status_code != 200:
        raise RuntimeError(f"Error HTTP Banxico ({response.status_code}) para fecha {fecha_str}")
    
    datos = response.json()['bmx']['series'][0]['datos']

    # Filtrar los dias festivos e inhabiles que contienen N/N.
    validos = [d for d in datos if d.get('dato') != 'N/N']
    if not validos:
        raise RuntimeError(f"Sin dato FIX valido para fecha {fecha_str}")
    return float(validos[-1]['dato'])

def obtener_tc_dof_seguro(fecha_str):
    try:
        return obtener_tc_dof(fecha_str)
    except Exception as e:
        st.warning(f"No se obtuvo TC DOF para {fecha_str}: {e}")
        return None

# Cargar la tabla INPC desde Supabase a cache.
@st.cache_data(ttl=86400)
def obtener_tabla_inpc():
    res = supabase.table("inpc").select("*").execute()
    tabla = {}
    for item in (res.data or []):
        clave = f"{item['anio']}-{str(item['mes']).zfill(2)}"
        tabla[clave] = float(item["valor"])
    return tabla

# Busca el INPC en la caché cargada desde Supabase.
def obtener_inpc(anio, mes):
    tabla = obtener_tabla_inpc()
    clave = f"{str(anio)}-{str(mes).zfill(2)}"
    if clave not in tabla:
        raise RuntimeError(f"No hay INPC en la BD para {clave}")
    return tabla[clave]

# Calcular el factor INPC
def calcular_factor_inpc_sat(fecha_compra, fecha_venta):
    # Convertir la fecha a objetos datetime.
    dt_compra = pd.to_datetime(fecha_compra)
    dt_venta = pd.to_datetime(fecha_venta)
    # Valiar si la compra y venta ocurrieron en el mismo mes/año.
    if dt_compra.year == dt_venta.year and dt_compra.month == dt_venta.month:
        return 1.00000
    # Calcular el Factor de Actualizacion.
    dt_venta_anterior = dt_venta - pd.DateOffset(months=1) # Esto resta 1 mes a la fecha de venta para cumplir con Art. 129 LISR.
    inpc_venta = obtener_inpc(dt_venta_anterior.year, dt_venta_anterior.month) # Ya con el mes anterior a la venta
    inpc_compra = obtener_inpc(dt_compra.year, dt_compra.month)
    
    factor = inpc_venta / inpc_compra
    
    return max(1.0, round(factor, 5)) # Regresa el valor 1.0 si la division es menor a 1 y redondea a 5 decimales.

TAM_PAGINA = 1000

def fetch_all_rows(tabla, order_col="fecha", desc=False, tiebreak_col="id"):
    filas = []
    inicio = 0
    while True:
        query = (
            supabase.table(tabla)
            .select("*")
            .order(order_col, desc=desc)
        )
        if tiebreak_col:
            query = query.order(tiebreak_col, desc=desc)
        res = query.range(inicio, inicio + TAM_PAGINA - 1).execute()
        lote = res.data or []
        filas.extend(lote)
        if len(lote) < TAM_PAGINA:
            break
        inicio += TAM_PAGINA
    return filas

def construir_fecha_hora(fecha, hora):
    """Combina fecha + hora del reporte Alpaca en horario US Eastern (ET/EDT/EST)."""
    dt = datetime.combine(fecha, hora, tzinfo=ZoneInfo("America/New_York"))
    return dt.isoformat()

st.title("📊 Sistema de Gestión de Acciones e Impuestos")

mensaje_flash = st.session_state.pop("mensaje_flash", None)
if mensaje_flash:
    st.toast(mensaje_flash, icon=":material/check_circle:", duration=5)

tab_tablas, tab_dividendos, tab_reporte = st.tabs(["📜 Registros", "💰 Reporte de Dividendos", "📈 Reporte Fiscal (FIFO)"])

# ================================================
# PESTAÑA 1: TABLAS DE REGISTROS COMPRAS Y VENTAS.
# ================================================

with tab_tablas:
    # Tabla Historial de compras
    st.subheader("Historial de Compras Registradas")
    res_compras = fetch_all_rows("compras", order_col="fecha_hora", desc=True)

    cols_registros_c = ["fecha", "broker", "ticker", "cantidad", "precio_usd", "comision_usd"]
    df_temp_c = pd.DataFrame(columns=cols_registros_c)
    df_filtrado_c = pd.DataFrame(columns=cols_registros_c)

    # Crear dataframe temporal con las columnas de registros.
    if res_compras:
        df_temp_c = pd.DataFrame(res_compras)[cols_registros_c]
    else:
        st.info("No hay compras registradas en la base de datos.")

    if not df_temp_c.empty:
        # Se crea columna con Fecha_dt_c objeto de Pandas. De esta columna se crea nueva columna Anio_c.
        df_temp_c["Fecha_dt_c"] = pd.to_datetime(df_temp_c["fecha"])
        df_temp_c["Anio_c"] = df_temp_c["Fecha_dt_c"].dt.year
        with st.expander("🛒 Filtros de compras"):
            col_ano_c, col_broker_c, col_ticker_c = st.columns(3)
            with col_ano_c:
                anios_opt_c = sorted(df_temp_c["Anio_c"].unique(), reverse=True)
                sel_anos_c = st.multiselect("📅 Año", options=anios_opt_c, default=anios_opt_c, key="fc_anos")
            with col_broker_c:
                brokers_opt_c = sorted(df_temp_c["broker"].unique())
                sel_brokers_c = st.multiselect("🏦 Broker", options=brokers_opt_c, default=brokers_opt_c, key="fc_brokers")
            with col_ticker_c:
                tickers_opt_c = sorted(df_temp_c["ticker"].unique())
                sel_tickers_c = st.multiselect("🏷️ Ticker", options=tickers_opt_c, default=tickers_opt_c, key="fc_tickers")

            # Se obtiene el dataframe aplicando filtros y eliminado columnas de fechas y anio creadas para filtro.
            df_filtrado_c = df_temp_c[(df_temp_c["Anio_c"].isin(sel_anos_c)) & (df_temp_c["broker"].isin(sel_brokers_c)) & (df_temp_c["ticker"].isin(sel_tickers_c))].drop(columns=["Fecha_dt_c", "Anio_c"], errors="ignore")

        # Mostrar resumen de numero de registros y tabla
        if not df_filtrado_c.empty:
            st.caption(f"Mostrando {len(df_filtrado_c)} de {len(df_temp_c)} registros")
            st.dataframe(df_filtrado_c, width="stretch", column_config={"cantidad": st.column_config.NumberColumn(format="%.9f"), "precio_usd": st.column_config.NumberColumn(format="%.4f"), "comision_usd": st.column_config.NumberColumn(format="%.2f")})
        else:
            st.info("No hay registros para los filtros seleccionados.")

    st.divider()

    # Tabla Historial de ventas.
    st.subheader("Historial de Ventas Registradas")
    res_ventas = fetch_all_rows("ventas", order_col="fecha_hora", desc=True)

    cols_registros_v = ["fecha", "broker", "ticker", "cantidad", "precio_usd", "comision_usd"]
    df_temp_v = pd.DataFrame(columns=cols_registros_v)
    df_filtrado_v = pd.DataFrame(columns=cols_registros_v)

    # Crear dataframe temporal con las columnas de registros.
    if res_ventas:
        df_temp_v = pd.DataFrame(res_ventas)[cols_registros_v]
    else:
        st.info("No hay ventas registradas en la base de datos.")

    if not df_temp_v.empty:
        # Se crea columna con Fecha_dt_v objeto de Pandas. De esta columna se crea nueva columna Anio_v.
        df_temp_v["Fecha_dt_v"] = pd.to_datetime(df_temp_v["fecha"])
        df_temp_v["Anio_v"] = df_temp_v["Fecha_dt_v"].dt.year
        with st.expander("🏷️ Filtros de ventas"):
            col_ano_v, col_broker_v, col_ticker_v = st.columns(3)
            with col_ano_v:
                anios_opt_v = sorted(df_temp_v["Anio_v"].unique(), reverse=True)
                sel_anos_v = st.multiselect("📅 Año", options=anios_opt_v, default=anios_opt_v, key="fv_anos")
            with col_broker_v:
                brokers_opt_v = sorted(df_temp_v["broker"].unique())
                sel_brokers_v = st.multiselect("🏦 Broker", options=brokers_opt_v, default=brokers_opt_v, key="fv_brokers")
            with col_ticker_v:
                tickers_opt_v = sorted(df_temp_v["ticker"].unique())
                sel_tickers_v = st.multiselect("🏷️ Ticker", options=tickers_opt_v, default=tickers_opt_v, key="fv_tickers")
           
            # Se obtiene el dataframe aplicando filtros y eliminado columnas de fechas y anio creadas para filtro.
            df_filtrado_v = df_temp_v[(df_temp_v["Anio_v"].isin(sel_anos_v)) & (df_temp_v["broker"].isin(sel_brokers_v)) & (df_temp_v["ticker"].isin(sel_tickers_v))].drop(columns=["Fecha_dt_v", "Anio_v"], errors="ignore")
            
        # Mostrar resumen de numero de registros y tabla
        if not df_filtrado_v.empty:
            st.caption(f"Mostrando {len(df_filtrado_v)} de {len(df_temp_v)} registros")
            st.dataframe(df_filtrado_v, width="stretch", column_config={"cantidad": st.column_config.NumberColumn(format="%.9f"), "precio_usd": st.column_config.NumberColumn(format="%.4f"), "comision_usd": st.column_config.NumberColumn(format="%.2f")})
        else:
            st.info("No hay registros para los filtros seleccionados.")

# =================================
# PESTAÑA 2: REPORTE DE DIVIDENDOS.
# =================================
with tab_dividendos:
    st.subheader("Balance y Registros de Dividendos")

    q_divs = fetch_all_rows("dividendos", order_col="fecha", desc=True)

    if q_divs:
        cols = ["id", "fecha", "broker", "ticker", "monto_bruto_usd", "retencion_irs_usd"]
        df_divs = pd.DataFrame(q_divs)[cols]
        if df_divs.empty:
            st.info("No hay registros disponibles.")

        # Crear copia del dataframe con la tabla de Dividendos
        df_divs_temp = df_divs.copy()

        # Crear columna fecha_dt_divs objeto de Pandas y crear columna Anio_divs.
        if "fecha" in df_divs_temp.columns:
            df_divs_temp["Fecha_dt_divs"] = pd.to_datetime(df_divs_temp["fecha"])
            df_divs_temp["Anio_divs"] = df_divs_temp["Fecha_dt_divs"].dt.year
        else:
            df_divs_temp["Anio_divs"] = "Sin fecha"

        with st.expander("💰 Filtros Dividendos"):
            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1:
                anios_disponibles = ["Todos"] + sorted(list(df_divs_temp["Anio_divs"].unique()), reverse=True)
                filtro_anio_divs = st.selectbox("Ejercicio Fiscal (Año)", options=anios_disponibles, key="f_div_anio")
            with col_d2:
                brokers_disponibles = ["Todos"] + list(df_divs_temp["broker"].unique())
                filtro_brokers_divs = st.selectbox("Filtrar por Broker", options=brokers_disponibles, key="f_div_broker")
            with col_d3:
                tickers_disponibles = list(df_divs_temp["ticker"].unique())
                filtro_tickers_divs = st.multiselect("Filtrar por Ticker", options=tickers_disponibles, default=tickers_disponibles, key="f_div_ticker")

        # Aplicacion de Filtros
        if filtro_anio_divs != "Todos":
            df_divs_temp = df_divs_temp[df_divs_temp["Anio_divs"] == filtro_anio_divs]
        if filtro_brokers_divs != "Todos":
            df_divs_temp = df_divs_temp[df_divs_temp["broker"] == filtro_brokers_divs]
        if filtro_tickers_divs:
            df_divs_temp = df_divs_temp[df_divs_temp["ticker"].isin(filtro_tickers_divs)]

        if not df_divs_temp.empty:
            # Calculo de valores fiscales con Tasa de Cambio del DOF.
            filas_calculadas = []
            for _, row in df_divs_temp.iterrows():
                fecha_pago = row["fecha"]
                monto_bruto_usd = float(row["monto_bruto_usd"])
                ret_irs_usd = float(row.get("retencion_irs_usd", 0.0))

                # Obtener el Tipo de Cambio DOF Oficial para las fechas de pago
                tc_dof = obtener_tc_dof_seguro(fecha_pago)
                if tc_dof is None:
                    continue

                # Realizar conversiones a MXN
                monto_bruto_mxn = monto_bruto_usd * tc_dof
                ret_irs_mxn = ret_irs_usd * tc_dof
                monto_neto_mxn = monto_bruto_mxn - ret_irs_mxn
                isr_10_mxn = monto_bruto_mxn * 0.10

                filas_calculadas.append({
                    "id": row["id"],
                    "Fecha": fecha_pago,
                    "Ticker": row["ticker"],
                    "Broker": row["broker"],
                    "Monto Bruto (USD)": monto_bruto_usd,
                    "Retencion IRS (USD)": ret_irs_usd,
                    "TC DOF": tc_dof,
                    "Ingreso Bruto (MXN)": monto_bruto_mxn,
                    "IRS retenido extranjero (MXN)": ret_irs_mxn,
                    "Ingreso Neto (MXN)": monto_neto_mxn,
                    "ISR 10% México (MXN)": isr_10_mxn
                })

            df_res_div = pd.DataFrame(filas_calculadas)

            if df_res_div.empty:
                st.warning("No se pudo calcular ningun dividendo (Falto TC DOF).")
            else:
                # Metricas DE RESUMEN (VALORES DIRECTOS PARA EL SAT)
                total_bruto_usd = df_res_div["Monto Bruto (USD)"].sum()
                total_irs_usd = df_res_div["Retencion IRS (USD)"].sum()
                total_bruto_mxn = df_res_div["Ingreso Bruto (MXN)"].sum()
                total_irs_mxn = df_res_div["IRS retenido extranjero (MXN)"].sum()
                total_neto_mxn = df_res_div["Ingreso Neto (MXN)"].sum()
                total_isr10_mxn = df_res_div["ISR 10% México (MXN)"].sum()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Ingreso Bruto Acumulable", f"${total_bruto_mxn:,.2f} MXN", f"${total_bruto_usd:,.2f} USD", delta_color="off")
                m2.metric("Impuesto Retenido EE.UU. (Acreditable)", f"${total_irs_mxn:,.2f} MXN", f"${total_irs_usd:,.2f} USD", delta_color="off")
                m3.metric("Ingreso Neto Percibido", f"${total_neto_mxn:,.2f} MXN")
                m4.metric("ISR 10% México (Estimado)", f"${total_isr10_mxn:,.2f} MXN")

                st.divider()

                # TABLA DESGLOSADA DE DIVIDENDOS
                st.dataframe(
                    df_res_div.drop(columns=["id"]),
                    width="stretch",
                    column_config={
                        "Monto Bruto (USD)": st.column_config.NumberColumn(format="$%.2f"),
                        "Retencion IRS (USD)": st.column_config.NumberColumn(format="$%.2f"),
                        "TC DOF": st.column_config.NumberColumn(format="$%.4f"),
                        "Ingreso Bruto (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                        "IRS retenido extranjero (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                        "Ingreso Neto (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                        "ISR 10% México (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                    }
                )
        else:
            st.info("No hay dividendos registrados para los filtros seleccionados.")
    else:
        st.info("Aún no tienes registros de dividendos guardados.")

# =================================
# Pestaña 3: Reporte Fiscal (FIFO).
# =================================
with tab_reporte:
    st.subheader("Resumen Fiscal y Algoritmo PEPS (FIFO)")
    if "df_fifo" not in st.session_state:
        st.session_state["df_fifo"] = None
    ventas_tab = fetch_all_rows("ventas", order_col="fecha_hora", desc=False)
    # Filtros por Broker y por Año de Venta.
    with st.expander("Filtros para reporte"):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            broker_filtro = st.selectbox("Filtrar por Broker", ["Todos"] + BROKERS, key="f_reporte_broker")
        with col_f2:
            df_ventas_anio = pd.DataFrame(ventas_tab or [])
            
            if df_ventas_anio.empty or "fecha" not in df_ventas_anio.columns:
                anios_disponibles = ["Todos"]
            else:
                df_ventas_anio["anio"] = df_ventas_anio["fecha"].astype(str).str[:4]
                anios_disponibles = ["Todos"] + sorted(df_ventas_anio["anio"].unique(), reverse=True)
                        
            anio_filtro = st.selectbox("Ejercicio fiscal (Año de venta)", anios_disponibles, key="f_reporte_anio")

    calcular = st.button("Calcular Impuestos (FIFO)")

    if calcular:
        # Obtener Compras y ventas de Supabase (Todas las historicas).
        q_compras = fetch_all_rows("compras", order_col="fecha_hora", desc=False)
        q_ventas = ventas_tab

        if not q_compras or not q_ventas:
            st.session_state["df_fifo"] = None
            st.warning("Necesitas tener al menos una compra y una venta registrada para ejecutar el cálculo.")
        else:
            df_compras = pd.DataFrame(q_compras).sort_values(["fecha_hora","id"]).reset_index(drop=True)
            df_ventas = pd.DataFrame(q_ventas).sort_values(["fecha_hora","id"]).reset_index(drop=True)

            # Convertir las columnas numerocas
            for col in ["cantidad", "precio_usd", "comision_usd"]:
                df_compras[col] = pd.to_numeric(df_compras[col])
                df_ventas[col] = pd.to_numeric(df_ventas[col])

            resultados_fifo = []
            #Extraer combinaciones unicas de broker y ticker (broker, ticker) que estan en ventas.
            grupos = df_ventas[["broker", "ticker"]].drop_duplicates().to_numpy()

            for b, t in grupos:
                # Filtrar compras y ventas pertenecientes al mismo broker.
                compras_t = df_compras[(df_compras["broker"] == b) & (df_compras["ticker"] == t)].copy().to_dict('records')
                ventas_t = df_ventas[(df_ventas["broker"] == b) & (df_ventas["ticker"] == t)].copy().to_dict('records')

                lotes = []
                for c in compras_t:
                    cant_c = float(c["cantidad"])
                    lotes.append({
                        "id": c["id"],
                        "fecha_compra": c["fecha"],
                        "broker": c["broker"],
                        "ticker": c["ticker"],
                        "cantidad_inicial": cant_c,
                        "cantidad_disponible": cant_c,
                        "precio_compra_usd": float(c["precio_usd"]),
                        "comision_compra_usd": float(c.get("comision_usd", 0.0))
                    })

                for v in ventas_t:
                    cant_por_vender = float(v["cantidad"])
                    cant_vta_total = float(v["cantidad"])
                    precio_vta = float(v["precio_usd"])
                    comision_vta_total = float(v.get("comision_usd", 0.0))
                    fecha_vta = v["fecha"]
                    broker_vta = v["broker"]
                    ticker_vta = v["ticker"]

                    venta_incompleta = False

                    while cant_por_vender > 0 and len(lotes) > 0:
                        lote_actual = lotes[0]
                        cant_matcheada = min(cant_por_vender, lote_actual["cantidad_disponible"])

                        # Prorrateo proporcional de comisiones por accion matcheada
                        comision_compra_prop = (lote_actual["comision_compra_usd"] / lote_actual["cantidad_inicial"]) * cant_matcheada
                        comision_venta_prop = (comision_vta_total / cant_vta_total) * cant_matcheada

                        # Costo e Ingreso netos en USD (incluyendo las comisiones)
                        costo_compra_usd = (cant_matcheada * lote_actual["precio_compra_usd"]) + comision_compra_prop
                        ingreso_venta_usd = (cant_matcheada * precio_vta) - comision_venta_prop
                        ganancia_usd = ingreso_venta_usd - costo_compra_usd

                        # Conversion a MXN con tipo de cambio DOF en la fecha correspondiente
                        tc_compra = obtener_tc_dof_seguro(lote_actual["fecha_compra"])
                        tc_venta = obtener_tc_dof_seguro(fecha_vta)
                        if tc_compra is None or tc_venta is None:
                            st.error(f"Sin TC DOF para {ticker_vta} (compra: {lote_actual['fecha_compra']}, venta: {fecha_vta}).")
                            venta_incompleta = True
                            break

                        costo_compra_mxn_orig = costo_compra_usd * tc_compra
                        ingreso_venta_mxn = ingreso_venta_usd * tc_venta

                        # Factor de actualizacion INPC Art. 129 LISR
                        try:
                            factor_act = calcular_factor_inpc_sat(lote_actual["fecha_compra"], fecha_vta)
                        except Exception as e:
                            st.error(f"Sin INPC para {ticker_vta} (compra: {lote_actual['fecha_compra']} / venta: {fecha_vta}): {e}")
                            venta_incompleta = True
                            break

                        costo_compra_mxn_ajustado = costo_compra_mxn_orig * factor_act
                        ganancia_mxn = ingreso_venta_mxn - costo_compra_mxn_ajustado
                        
                        resultados_fifo.append({
                            "Ticker": ticker_vta,
                            "Broker": broker_vta,
                            "Cantidad": cant_matcheada,
                            "Fecha Compra": lote_actual["fecha_compra"],
                            "Precio Compra (USD)": lote_actual["precio_compra_usd"],
                            "Comision Compra (USD)": comision_compra_prop,
                            "Costo Compra (USD)": costo_compra_usd,
                            "TC DOF Compra": tc_compra,
                            "Costo Compra (MXN)": costo_compra_mxn_orig,
                            "Factor INPC": factor_act,
                            "Costo Ajustado (MXN)": costo_compra_mxn_ajustado,
                            "Fecha Venta": fecha_vta,
                            "Precio Venta (USD)": precio_vta,
                            "Comision Venta (USD)": comision_venta_prop,
                            "Ingreso Venta (USD)": ingreso_venta_usd,
                            "TC DOF Venta": tc_venta,
                            "Ingreso Venta (MXN)": ingreso_venta_mxn,
                            "Ganancia/Perdida (USD)": ganancia_usd,
                            "Ganancia/Perdida (MXN)": ganancia_mxn
                        })

                        cant_por_vender = round(cant_por_vender - cant_matcheada, 9)
                        lote_actual["cantidad_disponible"] = round(lote_actual["cantidad_disponible"] - cant_matcheada, 9)

                        if lote_actual["cantidad_disponible"] <= 0:
                            lotes.pop(0)

                    if not venta_incompleta and round(cant_por_vender, 9) > 0:
                        st.error(f"⚠️ Error: Intentas vender acciones de {t} sin registro de compra previo.")

            # Mostrar y Filtrar Resultados del analisis FIFO
            if resultados_fifo:
                df_res = pd.DataFrame(resultados_fifo)

                # Aplicar filtro por BROKER
                if broker_filtro != "Todos":
                    df_res = df_res[df_res["Broker"] == broker_filtro]

                # Aplicar filtro por Año
                if anio_filtro != "Todos":
                    df_res = df_res[df_res["Fecha Venta"].str.startswith(str(anio_filtro))]

                st.session_state["df_fifo"] = df_res
                st.session_state["fifo_anio"] = anio_filtro
                st.session_state["fifo_broker"] = broker_filtro

                if df_res.empty:
                    st.info(f"No hay ventas registradas para los filtros seleccionados.")
                else:
                    st.toast("Cálculo FIFO realizado con éxito", icon=":material/check_circle:", duration=5)
            else:
                st.session_state["df_fifo"] = None
                st.warning("No se generó ninguna fila del cálculo FIFO. Revisa los errores de TC DOF o INPC.")

    df_res = st.session_state["df_fifo"]
    if df_res is not None and not df_res.empty:
        anio_mostrar = st.session_state.get("fifo_anio", "")
        broker_mostrar = st.session_state.get("fifo_broker", "")

        tot_ingreso_usd = round(df_res["Ingreso Venta (USD)"].sum(), 2)
        tot_ingreso_mxn = round(df_res["Ingreso Venta (MXN)"].sum(), 2)
        tot_costo_usd = round(df_res["Costo Compra (USD)"].sum(), 2)
        tot_costo_mxn = round(df_res["Costo Ajustado (MXN)"].sum(), 2)
        tot_ganancia_usd = round(df_res["Ganancia/Perdida (USD)"].sum(), 2)
        tot_ganancia_mxn = round(df_res["Ganancia/Perdida (MXN)"].sum(), 2)
        tot_isr = round(max(0.0, tot_ganancia_mxn * 0.10), 2)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(label="Ingresos Totales", value=f"${tot_ingreso_mxn:,.2f} MXN", delta=f"${tot_ingreso_usd:,.2f} USD", delta_color="off")
        c2.metric(label="Costo Ajustado INPC", value=f"${tot_costo_mxn:,.2f} MXN", delta=f"{tot_costo_usd:,.2f} USD", delta_color="off")
        c3.metric(label="Ganancia/Pérdida Neta", value=f"{tot_ganancia_mxn:,.2f} MXN", delta=f"{tot_ganancia_usd:,.2f} USD", delta_color="normal")
        c4.metric(label="ISR Estimado (10%)", value=f"{tot_isr:,.2f} MXN")

        # BOTON PARA EXPORTAR EXCEL
        # Crear un buffer en memoria para no saturar el disco del servidor
        buffer_excel = io.BytesIO()

        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            # Exportar el DataFrame limpio a Excel
            df_res.to_excel(writer, sheet_name=f"FIFO {anio_mostrar}", index=False)

            # Acceder a las propiedades de openpyxl para darle formato estetico rapido
            workbook = writer.book
            worksheet = writer.sheets[f"FIFO {anio_mostrar}"]
            worksheet.views.sheetView[0].showGridLines = True

            # Autoajustar el ancho de las columnas para que no se corten los numeros
            for i, col in enumerate(worksheet.columns, start=1):
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(i)
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        st.download_button(
            label="📥 Descargar Papel de Trabajo (Excel)",
            data=buffer_excel.getvalue(),
            file_name=f"Papel_de_Trabajo_FIFO_{broker_mostrar}_{anio_mostrar}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_descarga_excel"
        )

        st.divider()
        st.subheader(f"Desglose Fiscal de Ventas ({anio_mostrar})")
        st.dataframe(df_res, width="stretch",
                    column_config={
                        "Cantidad": st.column_config.NumberColumn(format="%.9f"),
                        "Precio Compra (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "Comision Compra (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "Costo Compra (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "TC DOF Compra": st.column_config.NumberColumn(format="%.4f"),
                        "Costo Compra (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                        "Factor INPC": st.column_config.NumberColumn(format="%.5f"),
                        "Costo Ajustado (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                        "Precio Venta (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "Comision Venta (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "Ingreso Venta (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "TC DOF Venta": st.column_config.NumberColumn(format="%.4f"),
                        "Ingreso Venta (MXN)": st.column_config.NumberColumn(format="$%.2f"),
                        "Ganancia/Perdida (USD)": st.column_config.NumberColumn(format="%.2f"),
                        "Ganancia/Perdida (MXN)": st.column_config.NumberColumn(format="$%.2f")
                    })

# =====================================================
# Siderbar: Registrar Compra, Venta, Dividendos e INPC.
# =====================================================
with st.sidebar:
    st.subheader("🛒 Registro de Compra")
    with st.expander("➕ Registrar Nueva Compra", expanded=False):
        with st.form("form_compras", clear_on_submit=True):
            broker_c = st.selectbox("Broker", BROKERS, key="b_compra")
            fecha_c = st.date_input("Fecha de compra", key="f_compra")
            hora_c = st.time_input("Hora de la operacion (ET/EDT/EST)", key="h_compra", step=1)
            ticker_c = st.text_input("Ticker", key="t_compra").upper()
            cantidad_c = st.number_input("Cantidad de Acciones", min_value=0.0, format="%.9f", step=0.000000001, key="c_compra")
            precio_c = st.number_input("Precio de la Acción (USD)", min_value=0.0, format="%.5f", step=0.00001, key="p_compra")
            comision_c = st.number_input("Comision (USD)", min_value=0.0, value=0.0, format="%.5f", step=0.00001, key="com_compra")

            submit_compra = st.form_submit_button("Guardar Compra")

    if submit_compra:
        if ticker_c and cantidad_c > 0 and precio_c > 0:
            datos = {
                "broker": broker_c,
                "fecha": str(fecha_c),
                "fecha_hora": construir_fecha_hora(fecha_c, hora_c),
                "ticker": ticker_c,
                "cantidad": cantidad_c,
                "precio_usd": precio_c,
                "comision_usd": comision_c
            }
            supabase.table("compras").insert(datos).execute()
            st.session_state["mensaje_flash"] = f"¡Compra de {cantidad_c} de {ticker_c} en {broker_c} guardada!"
            st.rerun()
        else:
            st.error("Por favor completa el Ticker, Cantidad y Precio.")
    st.divider()

    st.subheader("🏷️ Registro de Venta")
    with st.expander("➕ Registrar Nueva Venta", expanded=False):
        with st.form("form_ventas", clear_on_submit=True):
            broker_v = st.selectbox("Broker", BROKERS, key="b_venta")
            fecha_v = st.date_input("Fecha de Venta", key="f_venta")
            hora_v = st.time_input("Hora de la operacion (ET/EDT/EST)", key="h_venta", step=1)
            ticker_v = st.text_input("Ticker", key="t_venta").upper()
            cantidad_v = st.number_input("Cantidad Vendida", min_value=0.0, format="%.9f", step=0.000000001, key="c_venta")
            precio_v = st.number_input("Precio de Venta (USD)", min_value=0.0, format="%.5f", step=0.00001, key="p_venta")
            comision_v = st.number_input("Comisión (USD)", min_value=0.0, value=0.0, format="%.5f", step=0.00001, key="com_venta")

            submit_venta = st.form_submit_button("Guardar Venta")

    if submit_venta:
        if ticker_v and cantidad_v > 0 and precio_v > 0:
            datos = {
                "broker": broker_v,
                "fecha": str(fecha_v),
                "fecha_hora": construir_fecha_hora(fecha_v, hora_v),
                "ticker": ticker_v,
                "cantidad": cantidad_v,
                "precio_usd": precio_v,
                "comision_usd": comision_v
            }
            supabase.table("ventas").insert(datos).execute()
            st.session_state["mensaje_flash"] = f"¡Venta de {cantidad_v} de {ticker_v} en {broker_v} guardada!"
            st.rerun()
        else:
            st.error("Por favor completa el Ticker, Cantidad y Precio.")
    st.divider()

    st.subheader("💰 Registro de Dividendo")
    with st.expander("➕ Registrar Nuevo Dividendo", expanded=False):
        with st.form("form_dividendo", clear_on_submit=True):
            broker_div = st.selectbox("Broker", BROKERS, key="div_broker")
            fecha_div = st.date_input("Fecha de Pago", key="div_fecha")
            ticker_div = st.text_input("Ticker", key="div_ticker").upper()
            monto_bruto_div = st.number_input("Monto Bruto (USD)", min_value=0.0, format="%.2f", key="div_bruto")
            retencion_div = st.number_input("Retención IRS (USD)", min_value=0.0, value=0.0, format="%.2f", key="div_retencion")

            submit_dividendo = st.form_submit_button("Guardar Dividendo")

    if submit_dividendo:
        if ticker_div and monto_bruto_div > 0:
            datos = {
                "fecha": str(fecha_div),
                "broker": broker_div,
                "ticker": ticker_div,
                "monto_bruto_usd": monto_bruto_div,
                "retencion_irs_usd": retencion_div
            }
            supabase.table("dividendos").insert(datos).execute()
            st.session_state["mensaje_flash"] = "Dividendo registrado correctamente."
            st.rerun()
        else:
            st.error("Por favor ingresa un Ticker y un Monto Bruto válido.")
    st.divider()

    st.subheader("📜 Registro Mensual INPC (DOF)")
    with st.expander("➕ Registrar Nuevo Valor INPC", expanded=False):
        with st.form("form_inpc", clear_on_submit=True):
            anio_actual = datetime.now().year
            anio_i = st.number_input("Año", min_value=None, max_value=None, value=anio_actual, step=1)
            mes_i = st.number_input("Mes", min_value=1, max_value=12, step=1)
            valor_i = st.number_input("Valor INPC (DOF)", min_value=0.0, format="%.3f")

            submit_inpc = st.form_submit_button("Guardar INPC")

    if submit_inpc:
        if valor_i > 0:
            datos_inpc = {
                "anio": int(anio_i),
                "mes": int(mes_i),
                "valor": round(valor_i, 3)
            }
            supabase.table("inpc").insert(datos_inpc).execute()

            # Limpiar cache para mostrar el nuevo valor INPC
            obtener_tabla_inpc.clear()
            st.session_state["mensaje_flash"] = f"INPC de {anio_i}-{str(mes_i).zfill(2)} ({valor_i}) guardado correctamente."
            st.rerun()
        else:
            st.error("Por favor ingresa un valor de INPC válido.")