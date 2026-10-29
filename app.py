import streamlit as st # Framework para crear la interfaz web de la aplicación
import pandas as pd # Librería para manejar datos, cálculos y exportar a Excel
from datetime import date, datetime, timedelta # Para manejar las fechas del calendario
from io import BytesIO # Para generar el Excel en memoria
from sqlalchemy import text # Para escribir consultas SQL con parámetros seguros
import hmac # Para comparar la contraseña de forma segura

# Configuración de la página en el navegador
st.set_page_config(page_title="Control Lechero Agro", page_icon="🐄", layout="wide")

ROL_SOCIO = "Socio (Solo Lectura)"
ROL_ADMIN = "Administrador (Tú)"


# =====================================================================
# CAPA DE DATOS Y LÓGICA DE NEGOCIO (PostgreSQL en Supabase)
# =====================================================================

def db():
    # Lee la URI desde st.secrets -> [connections.postgresql]
    return st.connection('postgresql', type='sql')

def consultar(sql, **params):
    # ttl=0: siempre datos frescos, así los socios ven al instante lo que registra el administrador
    return db().query(sql, params=params, ttl=0)

def es_admin():
    return st.session_state.get('es_admin', False)

def ejecutar(sql, **params):
    # Toda escritura pasa por aquí; si no es administrador se bloquea aunque la interfaz fallara
    if not es_admin():
        raise PermissionError("Solo el administrador puede modificar datos.")
    with db().session as s:
        resultado = s.execute(text(sql), params)
        fila = resultado.first() if resultado.returns_rows else None
        s.commit()
        return fila

def obtener_precio():
    return float(consultar("SELECT precio_litro FROM config ORDER BY id LIMIT 1").iloc[0, 0])

def actualizar_precio(precio):
    ejecutar("UPDATE config SET precio_litro = :precio", precio=precio)

def crear_periodo(fecha_inicio, fecha_fin, precio):
    fila = ejecutar("INSERT INTO periodos (fecha_inicio, fecha_fin, precio_aplicado) "
                    "VALUES (:inicio, :fin, :precio) RETURNING id",
                    inicio=fecha_inicio, fin=fecha_fin, precio=precio)
    return fila[0]

def listar_periodos():
    # Todos los periodos, del más reciente al más antiguo
    return consultar("SELECT id, fecha_inicio, fecha_fin, precio_aplicado FROM periodos ORDER BY id DESC")

def obtener_registro(periodo_id, fecha):
    df = consultar("SELECT litros_manana, litros_tarde FROM registros WHERE periodo_id = :pid AND fecha = :fecha",
                   pid=periodo_id, fecha=fecha)
    return None if df.empty else (float(df.iloc[0, 0]), float(df.iloc[0, 1]))

def guardar_registro(periodo_id, fecha, litros_m, litros_t):
    # Si el día ya existe en el periodo se actualiza; si no, se inserta
    ejecutar('''INSERT INTO registros (periodo_id, fecha, litros_manana, litros_tarde)
                VALUES (:pid, :fecha, :m, :t)
                ON CONFLICT (periodo_id, fecha)
                DO UPDATE SET litros_manana = EXCLUDED.litros_manana,
                              litros_tarde = EXCLUDED.litros_tarde,
                              actualizado_en = now()''',
             pid=periodo_id, fecha=fecha, m=litros_m, t=litros_t)

def obtener_registros(periodo_id=None):
    # Registros con los cálculos de litros y pago; si se pasa periodo_id, solo los de esa semana
    consulta = '''
        SELECT r.periodo_id, r.fecha, r.litros_manana, r.litros_tarde, p.precio_aplicado
        FROM registros r
        JOIN periodos p ON r.periodo_id = p.id
    '''
    params = {}
    if periodo_id is not None:
        consulta += " WHERE r.periodo_id = :pid"
        params['pid'] = periodo_id
    consulta += " ORDER BY r.fecha"
    df = consultar(consulta, **params)
    df['fecha'] = pd.to_datetime(df['fecha']).dt.date
    df['Total_Diario_Lts'] = df['litros_manana'] + df['litros_tarde']
    df['Pago_Diario_$'] = df['Total_Diario_Lts'] * df['precio_aplicado']
    return df

def generar_excel(hojas):
    # hojas: {"Nombre de la hoja": DataFrame}
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for nombre, df in hojas.items():
            df.to_excel(writer, index=False, sheet_name=nombre)
    return output.getvalue()

def respaldo_completo():
    # Copia de las 3 tablas tal como están en Supabase
    return generar_excel({
        'config': consultar("SELECT * FROM config ORDER BY id"),
        'periodos': consultar("SELECT * FROM periodos ORDER BY id"),
        'registros': consultar("SELECT * FROM registros ORDER BY id"),
    })


# =====================================================================
# UTILIDADES DE PRESENTACIÓN
# =====================================================================

def pesos(valor):
    # Formato colombiano: $1.005.480
    return "$" + f"{valor:,.0f}".replace(",", ".")

def litros(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " L"

def etiqueta_periodo(fila):
    inicio = pd.to_datetime(fila['fecha_inicio']).strftime('%d/%m/%Y')
    fin = pd.to_datetime(fila['fecha_fin']).strftime('%d/%m/%Y')
    return f"Semana #{int(fila['id'])}  ·  {inicio} → {fin}  ·  {pesos(fila['precio_aplicado'])}/L"


# =====================================================================
# SECCIÓN A: 📝 REGISTRO OPERATIVO (solo administrador)
# =====================================================================

def seccion_registro():
    st.title("📝 Registro Operativo")
    st.caption("Crea el plazo de recolección y anota los litros de cada ordeño.")

    # --- Crear periodo ---
    with st.container(border=True):
        st.subheader("📅 Nuevo plazo de recolección")
        col1, col2, col3 = st.columns([2, 2, 1.3], vertical_alignment="bottom")
        fecha_inicio = col1.date_input("Día de inicio", date.today(), format="DD/MM/YYYY")
        fecha_fin = col2.date_input("Día de fin", date.today() + timedelta(days=6), format="DD/MM/YYYY") # Por defecto 7 días
        if col3.button("➕ Crear periodo", type="primary", width="stretch"):
            if fecha_fin < fecha_inicio:
                st.error("El día de fin no puede ser anterior al día de inicio.")
            else:
                precio = obtener_precio()
                nuevo_id = crear_periodo(fecha_inicio, fecha_fin, precio)
                st.toast(f"Semana #{nuevo_id} creada con precio {pesos(precio)}/L", icon="✅")
        st.caption(f"Se aplicará el precio vigente: **{pesos(obtener_precio())}** por litro (se cambia en ⚙️ Configuración).")

    st.divider()

    # --- Registro diario ---
    periodos = listar_periodos()
    if periodos.empty:
        st.info("Todavía no hay periodos. Crea el primero arriba para empezar a registrar litros.")
        return

    activo = periodos.iloc[0] # El último periodo creado es el activo
    p_id = int(activo['id']) # int(): evita enviar a la base el tipo numpy.int64 que devuelve pandas
    p_inicio = pd.to_datetime(activo['fecha_inicio']).date()
    p_fin = pd.to_datetime(activo['fecha_fin']).date()

    with st.container(border=True):
        st.subheader("🥛 Registro diario de leche")
        st.markdown(f"**Periodo activo:** {etiqueta_periodo(activo)}")

        # Si hoy está dentro del periodo, se propone hoy; si no, el primer día
        sugerida = date.today() if p_inicio <= date.today() <= p_fin else p_inicio
        fecha_registro = st.date_input("Fecha a registrar", sugerida, min_value=p_inicio, max_value=p_fin,
                                       format="DD/MM/YYYY")

        # Si ese día ya tiene datos, se cargan para poder corregirlos
        existente = obtener_registro(p_id, fecha_registro)
        if existente:
            st.info("Este día ya tiene un registro. Si guardas, se reemplazará.")
        valor_m, valor_t = existente if existente else (0.0, 0.0)

        col_m, col_t = st.columns(2)
        litros_m = col_m.number_input("☀️ Litros Mañana", min_value=0.0, step=1.0, value=valor_m,
                                      key=f"m_{p_id}_{fecha_registro}")
        litros_t = col_t.number_input("🌙 Litros Tarde", min_value=0.0, step=1.0, value=valor_t,
                                      key=f"t_{p_id}_{fecha_registro}")

        st.metric("Total del día", litros(litros_m + litros_t))

        if st.button("💾 Guardar registro diario", type="primary", width="stretch"):
            guardar_registro(p_id, fecha_registro, litros_m, litros_t)
            st.success(f"Guardado: {litros(litros_m + litros_t)} en total para el {fecha_registro.strftime('%d/%m/%Y')}")

    # Resumen rápido de lo registrado en el periodo activo
    df_activo = obtener_registros(p_id)
    if not df_activo.empty:
        st.divider()
        dias_periodo = (p_fin - p_inicio).days + 1
        c1, c2, c3 = st.columns(3)
        c1.metric("Total días registrados", f"{len(df_activo)} de {dias_periodo}", border=True)
        c2.metric("Total litros del periodo", litros(df_activo['Total_Diario_Lts'].sum()), border=True)
        c3.metric("Total a pagar hasta hoy", pesos(df_activo['Pago_Diario_$'].sum()), border=True)


# =====================================================================
# SECCIÓN B: 💰 FACTURACIÓN Y ANÁLISIS (todos los roles)
# =====================================================================

def seccion_facturacion():
    st.title("💰 Facturación y Análisis")

    periodos = listar_periodos()
    if periodos.empty:
        st.info("No hay periodos creados todavía.")
        return

    # --- Liquidación semanal ---
    with st.container(border=True):
        st.subheader("🧾 Liquidación de la semana")
        # Las opciones son los IDs reales de la tabla 'periodos'; la etiqueta solo es lo que se ve
        etiquetas = {int(fila['id']): etiqueta_periodo(fila) for _, fila in periodos.iterrows()}
        semana_id = st.selectbox("Semana facturada", options=list(etiquetas.keys()),
                                 format_func=lambda pid: etiquetas[pid])

        df_semana = obtener_registros(semana_id)
        if df_semana.empty:
            st.warning("Esta semana todavía no tiene registros de leche.")
        else:
            total_m = df_semana['litros_manana'].sum()
            total_t = df_semana['litros_tarde'].sum()
            total_lts = df_semana['Total_Diario_Lts'].sum()
            total_pago = df_semana['Pago_Diario_$'].sum()

            c1, c2 = st.columns(2)
            c1.metric("💵 Total a pagar de la semana", pesos(total_pago), border=True)
            c2.metric("🥛 Total litros de la semana", litros(total_lts), border=True)

            c3, c4, c5 = st.columns(3)
            c3.metric("☀️ Total litros mañana", litros(total_m), border=True)
            c4.metric("🌙 Total litros tarde", litros(total_t), border=True)
            c5.metric("📈 Total promedio diario", litros(total_lts / len(df_semana)), border=True)

            st.dataframe(
                df_semana[['fecha', 'litros_manana', 'litros_tarde', 'Total_Diario_Lts', 'Pago_Diario_$']],
                hide_index=True, width="stretch",
                column_config={
                    'fecha': st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                    'litros_manana': st.column_config.NumberColumn("Mañana (L)", format="%.1f"),
                    'litros_tarde': st.column_config.NumberColumn("Tarde (L)", format="%.1f"),
                    'Total_Diario_Lts': st.column_config.NumberColumn("Total del día (L)", format="%.1f"),
                    'Pago_Diario_$': st.column_config.NumberColumn("Total pago del día ($)", format="localized"),
                },
            )

    st.divider()

    # --- Histórico ---
    df_historico = obtener_registros()
    with st.container(border=True):
        st.subheader("📊 Rendimiento histórico de las vacas")
        if df_historico.empty:
            st.info("Aún no hay registros para graficar.")
        else:
            serie = (df_historico.assign(fecha=pd.to_datetime(df_historico['fecha']))
                     .groupby('fecha')[['litros_manana', 'litros_tarde', 'Total_Diario_Lts']].sum()
                     .rename(columns={'litros_manana': 'Mañana', 'litros_tarde': 'Tarde',
                                      'Total_Diario_Lts': 'Total del día'}))
            st.line_chart(serie, y_label="Litros", x_label="Fecha")

            h1, h2, h3 = st.columns(3)
            h1.metric("Total litros histórico", litros(df_historico['Total_Diario_Lts'].sum()))
            h2.metric("Total pagado histórico", pesos(df_historico['Pago_Diario_$'].sum()))
            h3.metric("Total días registrados", len(df_historico))

    # --- Exportación ---
    with st.container(border=True):
        st.subheader("📥 Exportar")
        st.caption("El Excel incluye una hoja con la semana seleccionada y otra con todo el histórico.")
        st.download_button(label="📥 Descargar Excel",
                           data=generar_excel({'Semana Facturada': df_semana, 'Histórico': df_historico}),
                           file_name=f"Reporte_Lechero_Semana_{semana_id}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")


# =====================================================================
# SECCIÓN C: ⚙️ CONFIGURACIÓN Y RESPALDO (solo administrador)
# =====================================================================

def seccion_configuracion():
    st.title("⚙️ Configuración y Respaldo")

    col1, col2 = st.columns(2)

    with col1, st.container(border=True):
        st.subheader("💲 Precio por litro")
        precio_actual = obtener_precio()
        st.metric("Total precio vigente por litro", pesos(precio_actual))
        nuevo_precio = st.number_input("Nuevo precio por litro ($)", value=precio_actual, min_value=0.0, step=10.0)
        st.caption("El cambio se aplica a los periodos que crees desde ahora; los ya creados conservan su precio.")
        if st.button("Actualizar precio", type="primary", width="stretch"):
            actualizar_precio(nuevo_precio)
            st.success(f"Precio actualizado a {pesos(nuevo_precio)} por litro.")

    with col2, st.container(border=True):
        st.subheader("☁️ Respaldo")
        st.markdown("Los datos ya viven en la nube (Supabase). Aquí puedes descargar una copia completa "
                    "de las 3 tablas para guardarla en tu PC o en Drive.")
        st.download_button("💾 Descargar respaldo completo",
                           data=respaldo_completo(),
                           file_name=f"Respaldo_Lecheria_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")


# =====================================================================
# ROLES Y MENÚ DE NAVEGACIÓN
# =====================================================================

with st.sidebar:
    st.markdown("## 🐄 Control Lechero")
    st.caption("Sistema de Administración Lechera")
    st.divider()

    rol = st.selectbox("👤 Rol", [ROL_SOCIO, ROL_ADMIN])
    st.session_state['es_admin'] = False
    if rol == ROL_ADMIN:
        clave = st.text_input("Contraseña de administrador", type="password")
        if clave and hmac.compare_digest(clave, st.secrets["auth"]["admin_password"]):
            st.session_state['es_admin'] = True
            st.success("Modo administrador activo")
        elif clave:
            st.error("Contraseña incorrecta. Estás en modo lectura.")
        else:
            st.info("Escribe la contraseña para editar.")

    st.divider()
    if es_admin():
        SECCIONES = {
            "📝 Registro Operativo": seccion_registro,
            "💰 Facturación y Análisis": seccion_facturacion,
            "⚙️ Configuración y Respaldo": seccion_configuracion,
        }
    else:
        # Los socios solo ven la facturación y pueden exportar el Excel
        SECCIONES = {"💰 Facturación y Análisis": seccion_facturacion}
    seccion = st.radio("Menú", list(SECCIONES.keys()), label_visibility="collapsed")

    st.divider()
    try:
        st.caption(f"Precio vigente: {pesos(obtener_precio())}/L")
    except Exception as e:
        st.error(f"No se pudo conectar a la base de datos ({type(e).__name__}). Revisa la URI en los secrets.")
        st.stop()

SECCIONES[seccion]()
