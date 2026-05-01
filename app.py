"""
app.py — Demo Automatización Convenios EAFIT
Equipo Reto 2 · Beca IA EAFIT · 2025
"""
import streamlit as st
from datetime import date, timedelta
import sys, os

MESES_ES = {
    1:"01 - Enero", 2:"02 - Febrero", 3:"03 - Marzo", 4:"04 - Abril",
    5:"05 - Mayo", 6:"06 - Junio", 7:"07 - Julio", 8:"08 - Agosto",
    9:"09 - Septiembre", 10:"10 - Octubre", 11:"11 - Noviembre", 12:"12 - Diciembre"
}
# Mapa inverso para obtener el número del mes desde la etiqueta
MESES_INV = {v: k for k, v in MESES_ES.items()}

def selector_fecha(label, key, default=None):
    """Selector de fecha con día, mes y año en español."""
    if default is None:
        default = date.today()
    st.markdown(f"**{label}**")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        dia = st.selectbox("Día", list(range(1, 32)), index=default.day - 1, key=f"{key}_dia", label_visibility="collapsed")
    with c2:
        mes_nombre = st.selectbox("Mes", list(MESES_ES.values()), index=default.month - 1, key=f"{key}_mes", label_visibility="collapsed")
        mes = MESES_INV[mes_nombre]
    with c3:
        anio = st.selectbox("Año", list(range(2024, 2030)), index=list(range(2024, 2030)).index(default.year) if default.year in range(2024, 2030) else 1, key=f"{key}_anio", label_visibility="collapsed")
    try:
        return date(anio, mes, min(dia, 28 if mes == 2 else 30 if mes in [4,6,9,11] else 31))
    except:
        return date(anio, mes, 1)
import importlib.util as _ilu
_BASE = os.path.dirname(os.path.abspath(__file__))
_spec = _ilu.spec_from_file_location("generador", os.path.join(_BASE, "generador.py"))
_gen  = _ilu.module_from_spec(_spec)
_spec.loader.exec_module(_gen)
generar_documento = _gen.generar_documento

# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Convenios EAFIT · Talento",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { max-width: 780px; margin: auto; }
    .header-box {
        background: linear-gradient(135deg, #003087 0%, #0057a8 100%);
        padding: 24px 28px 18px;
        border-radius: 12px;
        margin-bottom: 24px;
        color: white;
    }
    .header-box h1 { font-size: 22px; margin: 0 0 4px; }
    .header-box p  { font-size: 13px; opacity: .85; margin: 0; }
    .step-label {
        font-size: 11px; font-weight: 700; text-transform: uppercase;
        letter-spacing: .08em; color: #0057a8; margin-bottom: 4px;
    }
    .section-title {
        font-size: 15px; font-weight: 700;
        border-left: 4px solid #0057a8;
        padding-left: 10px; margin: 20px 0 12px;
    }
    .plantilla-box {
        background: #eef4ff; border: 1.5px solid #0057a8;
        border-radius: 8px; padding: 12px 16px; margin: 12px 0;
        font-size: 13px;
    }
    .plantilla-box b { color: #003087; }
    .success-box {
        background: #f0fdf4; border: 1.5px solid #22c55e;
        border-radius: 10px; padding: 16px 20px; margin-top: 16px;
    }
    div[data-testid="stAlert"] { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── Encabezado ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
  <h1>🎓 Generador de Convenios · Talento EAFIT</h1>
  <p>Complete el formulario para generar automáticamente el convenio o acuerdo de práctica/pasantía.</p>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# PASO 1 — Tipo de experiencia y condiciones
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="step-label">Paso 1 de 4 — Tipo de experiencia</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">¿Qué tipo de convenio se va a generar?</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    tipo_exp = st.radio(
        "Tipo de experiencia",
        ["Práctica", "Inmersión / Pasantía"],
        help="Práctica = vinculación formativa. Inmersión = acuerdo de pasantía tripartito."
    )
with col2:
    remunerada = st.radio(
        "¿El estudiante recibe auxilio económico?",
        ["Sí, es remunerada", "No, es gratuita"],
    ) == "Sí, es remunerada"

quien_arl = st.radio(
    "¿Quién paga la ARL del estudiante?",
    ["La organización / empresa", "EAFIT"],
    horizontal=True,
)
quien_paga_arl = "Organización" if "organización" in quien_arl.lower() else "EAFIT"

# Mostrar qué plantilla se usará
NOMBRES_PLANTILLAS = {
    ("Práctica", True,  "Organización"): "Convenio de Vinculación Formativa REMUNERADA · Empresa paga ARL",
    ("Práctica", True,  "EAFIT"):        "Convenio de Vinculación Formativa REMUNERADA · EAFIT paga ARL",
    ("Práctica", False, "Organización"): "Convenio de Vinculación Formativa NO REMUNERADA · Empresa paga ARL",
    ("Práctica", False, "EAFIT"):        "Convenio de Vinculación Formativa NO REMUNERADA · EAFIT paga ARL",
}
tipo_key = tipo_exp if tipo_exp == "Práctica" else "Práctica"
nombre_plantilla = NOMBRES_PLANTILLAS.get(
    (tipo_key, remunerada, quien_paga_arl),
    "Convenio de Vinculación Formativa"
)
st.markdown(f"""
<div class="plantilla-box">
  📄 <b>Plantilla seleccionada automáticamente:</b><br>
  {nombre_plantilla}
</div>
""", unsafe_allow_html=True)

st.divider()

# ═══════════════════════════════════════════════════════════════════════════
# PASO 2 — Datos de la empresa
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="step-label">Paso 2 de 4 — Datos de la organización</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Organización / Empresa</div>', unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    nombre_empresa = st.text_input("Razón social de la empresa *", placeholder="Ej: TECNOLOGÍA INNOVADORA S.A.S.")
with col2:
    tipo_sociedad = st.selectbox("Tipo de sociedad", ["S.A.S.", "S.A.", "Ltda.", "E.U.", "Fundación", "Otro"])

col1, col2 = st.columns(2)
with col1:
    nit_empresa = st.text_input("NIT *", placeholder="Ej: 900.123.456-7")
with col2:
    ciudad_empresa = st.text_input("Ciudad de la empresa *", placeholder="Ej: Medellín, Antioquia")

st.markdown("**Representante legal**")
col1, col2, col3 = st.columns(3)
with col1:
    nombre_rep = st.text_input("Nombre completo *", placeholder="Nombre del representante")
with col2:
    cedula_rep = st.text_input("Cédula *", placeholder="Ej: 71.234.567")
with col3:
    cargo_rep = st.text_input("Cargo *", placeholder="Ej: Gerente General")

st.divider()

# ═══════════════════════════════════════════════════════════════════════════
# PASO 3 — Datos del estudiante y práctica
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="step-label">Paso 3 de 4 — Datos del estudiante y práctica</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Estudiante</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    nombre_est = st.text_input("Nombre completo del estudiante *", placeholder="Nombre completo")
with col2:
    cedula_est = st.text_input("Cédula del estudiante *", placeholder="Ej: 1.001.234.567")

programa = st.text_input("Programa académico *", placeholder="Ej: Ingeniería de Sistemas")

st.markdown('<div class="section-title">Período de la práctica</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    fecha_inicio = selector_fecha("Fecha de inicio *", "inicio", date.today())
with col2:
    fecha_fin = selector_fecha("Fecha de fin *", "fin", date.today() + timedelta(days=180))

if remunerada:
    st.markdown('<div class="section-title">Auxilio económico</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        monto_letras = st.text_input("Monto en letras *", placeholder="Ej: UN MILLÓN TRESCIENTOS MIL")
    with col2:
        monto_numeros = st.text_input("Monto en números *", placeholder="Ej: 1.300.000")
else:
    monto_letras = ""
    monto_numeros = ""

st.markdown('<div class="section-title">Supervisión</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    st.markdown("**Tutor (designado por la empresa)**")
    nombre_tutor = st.text_input("Nombre del tutor *", placeholder="Nombre completo")
    cedula_tutor = st.text_input("Cédula del tutor *", placeholder="Ej: 98.765.432")
with col2:
    st.markdown("**Monitor (designado por EAFIT)**")
    nombre_monitor = st.text_input("Nombre del monitor *", placeholder="Nombre completo")
    cedula_monitor = st.text_input("Cédula del monitor *", placeholder="Ej: 43.111.222")

st.divider()

# ═══════════════════════════════════════════════════════════════════════════
# PASO 4 — Actividades y firma
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="step-label">Paso 4 de 5 — Actividades y fecha de firma</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Actividades del estudiante</div>', unsafe_allow_html=True)
st.caption("Ingrese entre 4 y 8 actividades/funciones que realizará el estudiante.")

actividades = []
cols = st.columns(2)
for i in range(6):
    with cols[i % 2]:
        act = st.text_input(
            f"Actividad {i+1}" + (" *" if i < 4 else " (opcional)"),
            key=f"act_{i}",
            placeholder=f"Describir actividad {i+1}..."
        )
        if act.strip():
            actividades.append(act.strip())

fecha_firma = selector_fecha("Fecha de firma del convenio *", "firma", date.today())

st.divider()

# ═══════════════════════════════════════════════════════════════════════════
# PASO 5 — Correos para el flujo de firmas
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="step-label">Paso 5 de 5 — Flujo de firmas digital</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📧 Correos de los firmantes</div>', unsafe_allow_html=True)
st.caption("El convenio será enviado automáticamente a cada parte en el orden correcto.")

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**1️⃣ Organización**")
    email_org = st.text_input(
        "Email representante legal",
        placeholder="firma@empresa.com",
        key="email_org",
    )
with col2:
    st.markdown("**2️⃣ Estudiante**")
    email_est = st.text_input(
        "Email institucional EAFIT",
        placeholder="estudiante@eafit.edu.co",
        key="email_est",
    )
with col3:
    st.markdown("**3️⃣ EAFIT**")
    email_eafit = st.text_input(
        "Email asesor Talento EAFIT",
        placeholder="asesor@eafit.edu.co",
        key="email_eafit",
    )

# Pipeline visual estático (siempre visible)
import streamlit.components.v1 as components
components.html("""
<div style="
  background:#f8faff;
  border:1.5px solid #d1ddf7;
  border-radius:12px;
  padding:16px 20px;
  font-family: sans-serif;
">
  <div style="font-size:11px; font-weight:700; color:#0057a8; text-transform:uppercase;
              letter-spacing:.06em; margin-bottom:12px;">
    Orden secuencial de firmas
  </div>
  <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">

    <div style="background:white; border:2px solid #0057a8; border-radius:10px;
                padding:10px 14px; min-width:130px; text-align:center;">
      <div style="font-size:20px;">&#127970;</div>
      <div style="font-size:12px; font-weight:700; color:#003087; margin-top:4px;">Organizaci&#243;n</div>
      <div style="font-size:10px; color:#6b7280; margin-top:2px;">Representante legal</div>
      <div style="font-size:10px; color:#059669; font-weight:600; margin-top:4px;">Plazo: 48 h</div>
    </div>

    <div style="color:#9ca3af; font-size:22px;">&rsaquo;</div>

    <div style="background:white; border:2px solid #0057a8; border-radius:10px;
                padding:10px 14px; min-width:130px; text-align:center;">
      <div style="font-size:20px;">&#128100;</div>
      <div style="font-size:12px; font-weight:700; color:#003087; margin-top:4px;">Estudiante</div>
      <div style="font-size:10px; color:#6b7280; margin-top:2px;">Al recibir confirmaci&#243;n</div>
      <div style="font-size:10px; color:#059669; font-weight:600; margin-top:4px;">Plazo: 48 h</div>
    </div>

    <div style="color:#9ca3af; font-size:22px;">&rsaquo;</div>

    <div style="background:white; border:2px solid #0057a8; border-radius:10px;
                padding:10px 14px; min-width:130px; text-align:center;">
      <div style="font-size:20px;">&#127891;</div>
      <div style="font-size:12px; font-weight:700; color:#003087; margin-top:4px;">EAFIT</div>
      <div style="font-size:10px; color:#6b7280; margin-top:2px;">Asesor Talento EAFIT</div>
      <div style="font-size:10px; color:#059669; font-weight:600; margin-top:4px;">Plazo: 24 h</div>
    </div>

    <div style="color:#9ca3af; font-size:22px;">&rsaquo;</div>

    <div style="background:#f0fdf4; border:2px solid #22c55e; border-radius:10px;
                padding:10px 14px; min-width:130px; text-align:center;">
      <div style="font-size:20px;">&#9989;</div>
      <div style="font-size:12px; font-weight:700; color:#16a34a; margin-top:4px;">Convenio listo</div>
      <div style="font-size:10px; color:#6b7280; margin-top:2px;">Archivado en Ceiba</div>
      <div style="font-size:10px; color:#059669; font-weight:600; margin-top:4px;">Autom&#225;tico</div>
    </div>

  </div>
</div>
""", height=160)

st.divider()

# ═══════════════════════════════════════════════════════════════════════════
# BOTÓN GENERAR
# ═══════════════════════════════════════════════════════════════════════════
campos_obligatorios = [
    nombre_empresa, nit_empresa, ciudad_empresa,
    nombre_rep, cedula_rep, cargo_rep,
    nombre_est, cedula_est, programa,
    nombre_tutor, cedula_tutor,
    nombre_monitor, cedula_monitor,
]
if remunerada:
    campos_obligatorios += [monto_letras, monto_numeros]

todo_completo = all(c.strip() for c in campos_obligatorios) and len(actividades) >= 4

if not todo_completo:
    faltantes = []
    if not nombre_empresa: faltantes.append("Razón social de la empresa")
    if not nit_empresa: faltantes.append("NIT")
    if not ciudad_empresa: faltantes.append("Ciudad de la empresa")
    if not nombre_rep: faltantes.append("Nombre del representante legal")
    if not nombre_est: faltantes.append("Nombre del estudiante")
    if not programa: faltantes.append("Programa académico")
    if not nombre_tutor: faltantes.append("Nombre del tutor")
    if not nombre_monitor: faltantes.append("Nombre del monitor")
    if len(actividades) < 4: faltantes.append(f"Actividades (mínimo 4, ingresadas {len(actividades)})")
    if remunerada and not monto_letras: faltantes.append("Monto del auxilio")
    
    if faltantes:
        st.warning("⚠️ Campos pendientes: " + " · ".join(faltantes))

generar = st.button(
    "⚡ Generar Convenio",
    type="primary",
    use_container_width=True,
    disabled=not todo_completo,
)

if generar and todo_completo:
    with st.spinner("Generando documento..."):
        datos = {
            'tipo_experiencia': "Práctica",
            'remunerada': remunerada,
            'quien_paga_arl': quien_paga_arl,
            'nombre_empresa': nombre_empresa.upper(),
            'tipo_sociedad': tipo_sociedad,
            'nit_empresa': nit_empresa,
            'nombre_representante': nombre_rep.upper(),
            'cedula_representante': cedula_rep,
            'cargo_representante': cargo_rep,
            'ciudad_empresa': ciudad_empresa,
            'nombre_estudiante': nombre_est.upper(),
            'cedula_estudiante': cedula_est,
            'programa_academico': programa,
            'fecha_inicio': fecha_inicio,
            'fecha_fin': fecha_fin,
            'nombre_tutor': nombre_tutor.upper(),
            'cedula_tutor': cedula_tutor,
            'nombre_monitor': nombre_monitor.upper(),
            'cedula_monitor': cedula_monitor,
            'actividades': actividades,
            'monto_letras': monto_letras,
            'monto_numeros': monto_numeros,
            'fecha_firma': fecha_firma,
        }
        docx_bytes = generar_documento(datos)
    
    nombre_archivo = f"Convenio_{nombre_est.split()[0]}_{nombre_empresa.split()[0]}_{fecha_inicio.strftime('%Y%m')}.docx"
    nombre_archivo = nombre_archivo.replace(" ", "_")
    
    st.markdown("""
    <div class="success-box">
      <b>✅ ¡Convenio generado exitosamente!</b><br>
      <span style="font-size:13px; color:#374151;">
        El documento está listo. Descárgalo o inicia el flujo de firmas digital.
      </span>
    </div>
    """, unsafe_allow_html=True)

    col_dl, col_sign = st.columns(2)
    with col_dl:
        st.download_button(
            label="⬇️ Descargar convenio (.docx)",
            data=docx_bytes,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with col_sign:
        iniciar_firmas = st.button(
            "🔏 Iniciar flujo de firmas",
            type="primary",
            use_container_width=True,
            disabled=not (email_org.strip() and email_est.strip() and email_eafit.strip()),
            help="Completa los 3 correos de firmantes para activar este botón."
        )

    st.info(f"""
    **Resumen del convenio generado:**
    📋 Tipo: {nombre_plantilla}
    🏢 Empresa: {nombre_empresa}
    👤 Estudiante: {nombre_est}
    📅 Período: {fecha_inicio.strftime('%d/%m/%Y')} → {fecha_fin.strftime('%d/%m/%Y')}
    📝 Actividades: {len(actividades)} registradas
    """)

    # ── Simulación del flujo de firmas ─────────────────────────────────────
    if iniciar_firmas:
        import time

        st.markdown("""
        <div style="background:#00205B; color:white; border-radius:12px;
                    padding:18px 22px; margin:16px 0 8px;">
          <div style="font-size:15px; font-weight:700; margin-bottom:4px;">
            🔏 Flujo de firmas iniciado
          </div>
          <div style="font-size:12px; opacity:.85;">
            El convenio se está enviando secuencialmente a cada firmante.
          </div>
        </div>
        """, unsafe_allow_html=True)

        pasos = [
            ("🏢", "Organización",  email_org,   "Representante legal", "Enviando enlace de firma…",    "✅ Enlace enviado — esperando firma"),
            ("👤", "Estudiante",    email_est,    "Practicante EAFIT",  "Notificando al estudiante…",   "✅ Notificación enviada — esperando firma"),
            ("🎓", "EAFIT",         email_eafit,  "Asesor Talento",     "Notificando al asesor…",       "✅ Convenio en bandeja del asesor"),
        ]

        for icono, nombre, email, rol, msg_cargando, msg_ok in pasos:
            with st.spinner(f"{icono} {nombre} ({email}) — {msg_cargando}"):
                time.sleep(1.4)
            st.markdown(f"""
            <div style="display:flex; align-items:center; gap:12px;
                        background:white; border:1.5px solid #22c55e;
                        border-radius:10px; padding:12px 16px; margin-bottom:8px;">
              <div style="font-size:26px;">{icono}</div>
              <div>
                <div style="font-size:13px; font-weight:700; color:#003087;">{nombre} — {rol}</div>
                <div style="font-size:12px; color:#374151;">📧 {email}</div>
                <div style="font-size:12px; color:#16a34a; font-weight:600; margin-top:2px;">{msg_ok}</div>
              </div>
              <div style="margin-left:auto; font-size:22px;">✅</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("""
        <div style="background:#f0fdf4; border:2px solid #22c55e; border-radius:12px;
                    padding:16px 20px; margin-top:12px; text-align:center;">
          <div style="font-size:22px; margin-bottom:6px;">🎉</div>
          <div style="font-size:15px; font-weight:700; color:#16a34a;">¡Flujo de firmas activado con éxito!</div>
          <div style="font-size:12px; color:#374151; margin-top:6px;">
            Cada firmante recibirá un correo con enlace de firma electrónica.<br>
            El convenio se archivará automáticamente en Ceiba una vez todos hayan firmado.
          </div>
          <div style="margin-top:12px; display:flex; justify-content:center; gap:16px;
                      font-size:11px; color:#6b7280;">
            <span>⏱ Tiempo estimado total: <b>2–4 días hábiles</b></span>
            <span>vs. proceso manual: <b>2–4 semanas</b></span>
          </div>
        </div>
        """, unsafe_allow_html=True)

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("🤖 Reto 2 · Motor de Procesos · Beca IA EAFIT · Equipo SER ANDI 2025")


# ═══════════════════════════════════════════════════════════════════════════
# SECCIÓN CARGA MASIVA (al final del formulario individual)
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<div style="background:#f0fdf4; border:2px solid #22c55e; border-radius:12px; padding:18px 22px; margin-top:8px">
  <h3 style="color:#16a34a; margin:0 0 6px; font-size:16px">⚡ Generación Masiva desde Excel</h3>
  <p style="font-size:12px; color:#374151; margin:0">
    ¿Tiene múltiples estudiantes? Cargue el Excel maestro y genere todos los convenios en segundos.
  </p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-title" style="margin-top:16px">Carga masiva de convenios</div>', unsafe_allow_html=True)

archivo_excel = st.file_uploader(
    "Subir Excel maestro (TALENTO_EAFIT_Carga_Masiva.xlsx)",
    type=["xlsx"],
    help="Use la plantilla Excel con los datos de todos los estudiantes."
)

if archivo_excel:
    import tempfile, sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else '.')
    from procesamiento_masivo import procesar_excel
    import time

    col1, col2 = st.columns([2,1])
    with col1:
        st.success(f"✅ Archivo cargado: **{archivo_excel.name}**")

    if st.button("🚀 Generar TODOS los convenios", type="primary", use_container_width=True):
        with st.spinner("Procesando todos los estudiantes..."):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(archivo_excel.getvalue())
                tmp_path = tmp.name

            t0 = time.time()
            zip_bytes, ok, errores = procesar_excel(tmp_path)
            t1 = time.time()
            os.unlink(tmp_path)

        st.markdown(f"""
        <div style="background:#f0fdf4; border:1.5px solid #22c55e; border-radius:10px; padding:16px 20px; margin:12px 0">
          <b style="color:#16a34a; font-size:15px">✅ {ok} convenios generados en {t1-t0:.1f} segundos</b><br>
          <span style="font-size:12px; color:#374151">
            Comparado con ~30 min por convenio manual = 
            <b style="color:#16a34a">{ok*30:,} minutos ahorrados ({ok*30//60} horas)</b>
          </span>
        </div>
        """, unsafe_allow_html=True)

        if errores:
            with st.expander(f"⚠️ {len(errores)} filas con error"):
                for e in errores:
                    st.write(e)

        st.download_button(
            label=f"⬇️ Descargar {ok} convenios (.zip)",
            data=zip_bytes,
            file_name=f"Convenios_EAFIT_{ok}_estudiantes.zip",
            mime="application/zip",
            use_container_width=True,
        )
