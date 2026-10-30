import streamlit as st # Framework para crear la interfaz web de la aplicación
import pandas as pd # Librería para manejar datos, cálculos y exportar a Excel
from datetime import date, datetime, timedelta # Para manejar las fechas del calendario
from io import BytesIO # Para generar el Excel en memoria
from sqlalchemy import text, func, select # Consultas SQL con parámetros seguros y consultas ORM
import hmac # Para comparar la contraseña de forma segura
from modelos import Base, GastoSemanal, Registro # Modelos ORM de las tablas

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

def exigir_admin():
    # Toda escritura llama a esto; si no es administrador se bloquea aunque la interfaz fallara
    if not es_admin():
        raise PermissionError("Solo el administrador puede modificar datos.")

def ejecutar(sql, **params):
    exigir_admin()
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
        'gastos_semanales': consultar("SELECT * FROM gastos_semanales ORDER BY id"),
    })


# --- Control financiero y punto de equilibrio ---

# Precio de referencia por bulto; en pantalla se puede editar para cada semana
CUIDOS = {
    "Levante (40kg)": 84071,
    "Finalizador (40kg)": 82060,
    "Gold Harina (40kg)": 75206,
    "Gold Pastilla (40kg)": 76267,
    "Terneras Nodrizas (40kg)": 89280,
    "Cinta Azul (40kg)": 85036,
    "Cuido de Perros (10kg)": 108267,
    "Sal Mineralizada al 8%": 39242,
    "Destete Preinicio": 167894,
    "Otro": 0,
}

@st.cache_resource
def preparar_tablas_financieras():
    # Crea gastos_semanales si no existe (una sola vez por servidor) y la protege de la API pública de Supabase
    engine = db().engine
    Base.metadata.create_all(engine, tables=[GastoSemanal.__table__])
    with engine.begin() as c:
        c.execute(text("ALTER TABLE gastos_semanales ENABLE ROW LEVEL SECURITY"))

def litros_producidos(periodo_id, session):
    # Suma de (mañana + tarde) de todos los ordeños del periodo
    total = session.scalar(
        select(func.coalesce(func.sum(Registro.litros_manana + Registro.litros_tarde), 0))
        .where(Registro.periodo_id == periodo_id)
    )
    return float(total)

# Tabla editable de cuidos comprados en la semana
COLUMNAS_CUIDOS = ['Insumo', 'Detalle', 'Bultos', 'Precio bulto', 'Subtotal']

def tabla_cuidos(detalle=None):
    # Convierte la lista guardada en la BD (detalle_cuidos) en la tabla que se edita en pantalla
    filas = [{'Insumo': c['insumo'], 'Detalle': c.get('detalle', ''), 'Bultos': float(c['cantidad_bultos']),
              'Precio bulto': float(c['precio_bulto']), 'Subtotal': float(c['subtotal'])} for c in detalle or []]
    df = pd.DataFrame(filas, columns=COLUMNAS_CUIDOS)
    return df.astype({'Insumo': 'object', 'Detalle': 'object', 'Bultos': 'float', 'Precio bulto': 'float',
                      'Subtotal': 'float'})

def completar_cuidos(editada, anterior):
    # Autocompleta el precio cuando el insumo es nuevo o cambió (si no, respeta el precio editado)
    # y recalcula el subtotal de cada fila
    filas = []
    for idx, fila in editada.iterrows():
        insumo = fila['Insumo'] if isinstance(fila['Insumo'], str) and fila['Insumo'] in CUIDOS else None
        bultos = 0.0 if pd.isna(fila['Bultos']) else float(fila['Bultos'])
        precio = None if pd.isna(fila['Precio bulto']) else float(fila['Precio bulto'])
        insumo_cambio = idx not in anterior.index or anterior.at[idx, 'Insumo'] != insumo
        # 'Otro' vale $0 en el diccionario: nunca borra un precio que el usuario ya escribió
        if insumo and (precio is None or (insumo_cambio and CUIDOS[insumo] > 0)):
            precio = float(CUIDOS[insumo])
        filas.append({'Insumo': insumo, 'Detalle': fila['Detalle'] if isinstance(fila['Detalle'], str) else '',
                      'Bultos': bultos, 'Precio bulto': precio,
                      'Subtotal': bultos * precio if precio is not None else 0.0})
    df = pd.DataFrame(filas, columns=COLUMNAS_CUIDOS)
    return df.astype({'Insumo': 'object', 'Detalle': 'object', 'Bultos': 'float', 'Precio bulto': 'float',
                      'Subtotal': 'float'})

def firma_cuidos(df):
    # Valores de la tabla sin NaN, para saber si el autocompletado cambió algo
    return [tuple(None if (isinstance(v, float) and pd.isna(v)) else v for v in fila)
            for fila in df[COLUMNAS_CUIDOS].itertuples(index=False)]

def detalle_desde_tabla(df):
    # Solo se guardan las filas con insumo y bultos
    return [{'insumo': f['Insumo'], 'detalle': f['Detalle'], 'cantidad_bultos': f['Bultos'],
             'precio_bulto': f['Precio bulto'], 'subtotal': f['Subtotal']}
            for _, f in df.iterrows() if f['Insumo'] and f['Bultos'] > 0]

def resumen_cuidos(detalle):
    # Ej.: "2 × Levante (40kg); 1 × Sal Mineralizada al 8%"
    return "; ".join(f"{c['cantidad_bultos']:g} × {c['insumo']}" + (f" ({c['detalle']})" if c.get('detalle') else "")
                     for c in detalle or [])

def calcular_liquidacion(litros_totales, precio_venta_litro, costo_cuidos, costos_fijos):
    ingresos_leche = litros_totales * precio_venta_litro
    costo_total_semana = costo_cuidos + costos_fijos
    return {
        'ingresos_leche': ingresos_leche,
        'costo_total_semana': costo_total_semana,
        # Costo de producir 1 litro (sin litros no se puede calcular)
        'punto_equilibrio_litro': costo_total_semana / litros_totales if litros_totales else 0.0,
        'utilidad_neta_semana': ingresos_leche - costo_total_semana,
    }

def obtener_liquidacion(periodo_id):
    df = consultar("SELECT * FROM gastos_semanales WHERE periodo_id = :pid", pid=periodo_id)
    return None if df.empty else df.iloc[0].to_dict()

def guardar_liquidacion(datos):
    # Una liquidación por periodo: si ya existe se reemplaza, para no contar dos veces la misma semana
    exigir_admin()
    with db().session as s:
        existente = s.scalar(select(GastoSemanal).where(GastoSemanal.periodo_id == datos['periodo_id']))
        if existente:
            for campo, valor in datos.items():
                setattr(existente, campo, valor)
        else:
            s.add(GastoSemanal(**datos))
        s.commit()
    return existente is not None

def listar_liquidaciones():
    # Liquidaciones con las fechas de su periodo, para mostrarlas y exportarlas
    df = consultar('''SELECT g.*, p.fecha_inicio, p.fecha_fin
                      FROM gastos_semanales g JOIN periodos p ON p.id = g.periodo_id
                      ORDER BY p.fecha_inicio, g.periodo_id''')
    for col in ('fecha_inicio', 'fecha_fin'):
        df[col] = pd.to_datetime(df[col]).dt.date
    df['resumen_cuidos'] = df['detalle_cuidos'].apply(resumen_cuidos)
    return df

def escribir_hoja(writer, df, nombre, columnas, sumables):
    # columnas: {columna_df: (título, formato)}; sumables: columnas que llevan total al final
    reporte = df[list(columnas)].rename(columns={k: v[0] for k, v in columnas.items()})
    reporte.to_excel(writer, index=False, sheet_name=nombre)
    libro, hoja = writer.book, writer.sheets[nombre]
    formatos = {'pesos': libro.add_format({'num_format': '$#,##0'}),
                'litros': libro.add_format({'num_format': '#,##0.0'}),
                'fecha': libro.add_format({'num_format': 'dd/mm/yyyy'}), None: None}
    negrita = {k: libro.add_format({'bold': True, 'top': 1, **({'num_format': f.num_format} if f else {})})
               for k, f in formatos.items()}
    for i, (_, (titulo, tipo)) in enumerate(columnas.items()):
        hoja.set_column(i, i, max(len(titulo) + 2, 14), formatos[tipo])

    # Fila de totales (solo columnas que tiene sentido sumar)
    fila = len(reporte) + 1
    hoja.write(fila, 0, 'TOTAL', negrita[None])
    for i, col in enumerate(columnas):
        if col in sumables:
            letra = chr(ord('A') + i)
            # Se guarda también el resultado para que visores sin recálculo (celular, vista previa) no muestren 0
            hoja.write_formula(fila, i, f"=SUM({letra}2:{letra}{fila})", negrita[columnas[col][1]],
                               float(reporte.iloc[:, i].sum()))
    hoja.freeze_panes(1, 0)

def excel_financiero(df):
    # Hoja 1: una fila por semana. Hoja 2: cada cuido comprado, por semana.
    detalle = pd.DataFrame(
        [{'periodo_id': f['periodo_id'], 'fecha_inicio': f['fecha_inicio'], 'fecha_fin': f['fecha_fin'], **c}
         for _, f in df.iterrows() for c in f['detalle_cuidos']],
        columns=['periodo_id', 'fecha_inicio', 'fecha_fin', 'insumo', 'detalle', 'cantidad_bultos',
                 'precio_bulto', 'subtotal'])

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', date_format='dd/mm/yyyy') as writer:
        escribir_hoja(writer, df, 'Reporte Financiero', {
            'periodo_id': ('Semana #', None), 'fecha_inicio': ('Fecha inicio', 'fecha'),
            'fecha_fin': ('Fecha fin', 'fecha'), 'litros_totales': ('Litros totales', 'litros'),
            'precio_venta_litro': ('Precio venta litro', 'pesos'), 'ingresos_leche': ('Ingresos leche', 'pesos'),
            'costo_cuidos': ('Costo cuidos', 'pesos'), 'costos_fijos': ('Costos fijos', 'pesos'),
            'costo_total_semana': ('Costo total semana', 'pesos'),
            'punto_equilibrio_litro': ('Punto equilibrio litro', 'pesos'),
            'utilidad_neta_semana': ('Utilidad neta semana', 'pesos'),
            'resumen_cuidos': ('Detalle cuidos', None),
        }, sumables={'litros_totales', 'ingresos_leche', 'costo_cuidos', 'costos_fijos',
                     'costo_total_semana', 'utilidad_neta_semana'})
        escribir_hoja(writer, detalle, 'Detalle Cuidos', {
            'periodo_id': ('Semana #', None), 'fecha_inicio': ('Fecha inicio', 'fecha'),
            'fecha_fin': ('Fecha fin', 'fecha'), 'insumo': ('Insumo', None), 'detalle': ('Detalle', None),
            'cantidad_bultos': ('Bultos', 'litros'), 'precio_bulto': ('Precio bulto', 'pesos'),
            'subtotal': ('Subtotal', 'pesos'),
        }, sumables={'cantidad_bultos', 'subtotal'})
    return output.getvalue()


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
# SECCIÓN D: 📈 CONTROL FINANCIERO Y PUNTO DE EQUILIBRIO
# (administrador registra; socios consultan y exportan)
# =====================================================================

def seccion_financiera():
    st.title("📈 Control Financiero")
    st.caption("Ingresos, costos, punto de equilibrio y utilidad de cada semana.")

    if es_admin():
        periodos = listar_periodos()
        if periodos.empty:
            st.info("No hay periodos creados todavía. Ve a 📝 Registro Operativo para crear el primero.")
            return

        # --- 1. Semana y producción ---
        with st.container(border=True):
            st.subheader("🥛 1. Semana a liquidar")
            col1, col2 = st.columns([3, 1.3], vertical_alignment="bottom")
            # Las opciones son los IDs reales de la tabla 'periodos'; la etiqueta solo es lo que se ve
            etiquetas = {int(fila['id']): etiqueta_periodo(fila) for _, fila in periodos.iterrows()}
            periodo_id = col1.selectbox("Semana (periodo)", options=list(etiquetas.keys()),
                                        format_func=lambda pid: etiquetas[pid], key="fin_periodo")
            precio_periodo = float(periodos.loc[periodos['id'] == periodo_id, 'precio_aplicado'].iloc[0])
            # Se consulta a la base cada vez que cambia la semana
            with db().session as s:
                litros_totales = litros_producidos(periodo_id, s)
            col2.metric("Total litros producidos", litros(litros_totales), border=True)
            if litros_totales == 0:
                st.warning("Esta semana todavía no tiene ordeños registrados.")

        # Si la semana ya tiene liquidación, se cargan sus valores para corregirlos
        guardada = obtener_liquidacion(periodo_id)

        # --- 2. Ingresos y costos ---
        with st.container(border=True):
            st.subheader("💵 2. Ingresos y costos fijos")
            c1, c2 = st.columns(2)
            precio_venta = c1.number_input("Precio de venta por litro ($)", min_value=0.0, step=10.0,
                                           value=float(guardada['precio_venta_litro']) if guardada else precio_periodo,
                                           key=f"fin_precio_{periodo_id}")
            costos_fijos = c2.number_input("Otros costos fijos de la semana ($)", min_value=0.0, step=1000.0,
                                           value=float(guardada['costos_fijos']) if guardada else 0.0,
                                           key=f"fin_fijos_{periodo_id}",
                                           help="Mano de obra, transporte, medicamentos, servicios, etc.")

        # --- 3. Cuidos comprados (varias filas) ---
        with st.container(border=True):
            st.subheader("🌾 3. Cuidos comprados en la semana")
            st.caption("Agrega una fila por cada insumo. El precio se llena solo al elegir el insumo, "
                       "pero puedes cambiarlo. Para borrar una fila, selecciónala y usa el ícono de papelera.")

            # La tabla vive en session_state por semana; 'version' renueva el editor tras autocompletar
            clave = f"cuidos_{periodo_id}"
            if clave not in st.session_state:
                st.session_state[clave] = tabla_cuidos(guardada['detalle_cuidos'] if guardada else None)
                st.session_state[clave + "_version"] = 0
            anterior = st.session_state[clave]

            editada = st.data_editor(
                anterior, num_rows="dynamic", hide_index=True, width="stretch",
                key=f"{clave}_editor_{st.session_state[clave + '_version']}",
                column_config={
                    'Insumo': st.column_config.SelectboxColumn("Insumo", options=list(CUIDOS.keys()),
                                                               required=True, width="large"),
                    'Detalle': st.column_config.TextColumn("Detalle (opcional)",
                                                           help="Ej. el nombre del producto si eliges 'Otro'"),
                    'Bultos': st.column_config.NumberColumn("Bultos", min_value=0.0, step=1.0, format="%.1f"),
                    'Precio bulto': st.column_config.NumberColumn("Precio bulto ($)", min_value=0.0, step=100.0,
                                                                  format="localized"),
                    'Subtotal': st.column_config.NumberColumn("Subtotal ($)", disabled=True, format="localized"),
                },
            )
            completada = completar_cuidos(editada, anterior)
            if firma_cuidos(completada) != firma_cuidos(editada):
                # Hubo autocompletado: se guarda y se vuelve a dibujar la tabla con el precio y el subtotal
                st.session_state[clave] = completada
                st.session_state[clave + "_version"] += 1
                st.rerun()
            st.session_state[clave] = completada

            detalle_cuidos = detalle_desde_tabla(completada)
            costo_cuidos = sum(c['subtotal'] for c in detalle_cuidos)
            if any(c['precio_bulto'] == 0 for c in detalle_cuidos):
                st.warning("Hay insumos con precio $0 (por ejemplo 'Otro'). Escribe su precio por bulto.")
            st.metric("Total costo cuidos", pesos(costo_cuidos))

        # --- 4. Resultado ---
        calculo = calcular_liquidacion(litros_totales, precio_venta, costo_cuidos, costos_fijos)
        with st.container(border=True):
            st.subheader("🧮 4. Resultado de la semana")
            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Total ingresos leche", pesos(calculo['ingresos_leche']), border=True)
            r2.metric("Total costos semana", pesos(calculo['costo_total_semana']), border=True)
            # El signo va al inicio del texto para que Streamlit pinte la flecha verde (+) o roja (-)
            margen = precio_venta - calculo['punto_equilibrio_litro']
            delta_margen = f"{'-' if margen < 0 else '+'}{pesos(abs(margen))} de margen por litro" if litros_totales else None
            utilidad = calculo['utilidad_neta_semana']
            r3.metric("Total punto de equilibrio por litro", pesos(calculo['punto_equilibrio_litro']),
                      delta=delta_margen, border=True)
            r4.metric("Total utilidad neta", pesos(utilidad),
                      delta=("Ganancia" if utilidad >= 0 else "Pérdida") if litros_totales else None,
                      delta_color="green" if utilidad >= 0 else "red",
                      delta_arrow="up" if utilidad >= 0 else "down", border=True)
            st.caption("Punto de equilibrio = lo que cuesta producir 1 litro. "
                       "Si es menor que el precio de venta, la semana deja ganancia.")

            if guardada:
                st.info("Esta semana ya tiene una liquidación guardada (se cargó arriba). Si guardas, se reemplazará.")
            if st.button("💾 Guardar Liquidación Semanal", type="primary", width="stretch",
                         disabled=litros_totales == 0):
                reemplazada = guardar_liquidacion({
                    'periodo_id': periodo_id, 'litros_totales': litros_totales,
                    'precio_venta_litro': precio_venta, 'costo_cuidos': costo_cuidos,
                    'detalle_cuidos': detalle_cuidos, 'costos_fijos': costos_fijos, **calculo,
                })
                st.success("Liquidación actualizada." if reemplazada else "Liquidación guardada.")

        st.divider()

    # --- Histórico financiero (todos los roles) ---
    df = listar_liquidaciones()
    with st.container(border=True):
        st.subheader("📚 Liquidaciones guardadas")
        if df.empty:
            st.info("Todavía no hay liquidaciones guardadas.")
            return

        t1, t2, t3 = st.columns(3)
        t1.metric("Total ingresos acumulados", pesos(df['ingresos_leche'].sum()), border=True)
        t2.metric("Total costos acumulados", pesos(df['costo_total_semana'].sum()), border=True)
        t3.metric("Total utilidad neta acumulada", pesos(df['utilidad_neta_semana'].sum()), border=True)

        moneda = lambda titulo: st.column_config.NumberColumn(titulo, format="localized")
        st.dataframe(
            df[['periodo_id', 'fecha_inicio', 'fecha_fin', 'litros_totales', 'precio_venta_litro', 'ingresos_leche',
                'costo_cuidos', 'resumen_cuidos', 'costos_fijos', 'costo_total_semana',
                'punto_equilibrio_litro', 'utilidad_neta_semana']],
            hide_index=True, width="stretch",
            column_config={
                'periodo_id': st.column_config.NumberColumn("Semana #", format="%d"),
                'fecha_inicio': st.column_config.DateColumn("Inicio", format="DD/MM/YYYY"),
                'fecha_fin': st.column_config.DateColumn("Fin", format="DD/MM/YYYY"),
                'litros_totales': st.column_config.NumberColumn("Litros", format="%.1f"),
                'precio_venta_litro': moneda("Precio venta ($/L)"),
                'ingresos_leche': moneda("Ingresos ($)"),
                'costo_cuidos': moneda("Costo cuidos ($)"),
                'resumen_cuidos': st.column_config.TextColumn("Cuidos comprados", width="large"),
                'costos_fijos': moneda("Costos fijos ($)"),
                'costo_total_semana': moneda("Costo total ($)"),
                'punto_equilibrio_litro': moneda("Equilibrio ($/L)"),
                'utilidad_neta_semana': moneda("Utilidad neta ($)"),
            },
        )

        st.download_button("📥 Descargar Reporte Financiero (Excel)",
                           data=excel_financiero(df),
                           file_name=f"Reporte_Financiero_{date.today().strftime('%Y-%m-%d')}.xlsx",
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
            "📈 Control Financiero": seccion_financiera,
            "⚙️ Configuración y Respaldo": seccion_configuracion,
        }
    else:
        # Los socios solo consultan y exportan: facturación y liquidaciones financieras
        SECCIONES = {
            "💰 Facturación y Análisis": seccion_facturacion,
            "📈 Control Financiero": seccion_financiera,
        }
    seccion = st.radio("Menú", list(SECCIONES.keys()), label_visibility="collapsed")

    st.divider()
    try:
        preparar_tablas_financieras()
        st.caption(f"Precio vigente: {pesos(obtener_precio())}/L")
    except Exception as e:
        st.error(f"No se pudo conectar a la base de datos ({type(e).__name__}). Revisa la URI en los secrets.")
        st.stop()

SECCIONES[seccion]()
