import streamlit as st
from google import genai

# --- INICIALIZAR MEMORIA DE SESIÓN ---
if 'autenticado_seo' not in st.session_state:
    st.session_state.autenticado_seo = False

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="DuoLogic IA | Eleven Agency", page_icon="⚡", layout="wide")

st.title("⚡ DuoLogic IA - Suite para Agencias")
st.markdown("Automatización de flujos de contenido SEO y Paid Media impulsada por IA.")

# --- BARRA LATERAL DE SEGURIDAD ---
with st.sidebar:
    st.header("🔐 Acceso al Sistema SEO")
    st.markdown("Ingresá la clave corporativa para habilitar la IA.")
    
    # El usuario pone la contraseña "visual"
    clave_usuario = st.text_input("Contraseña:", type="password")
    
    if st.button("Ingresar", type="primary"):
        if clave_usuario == "eleven2026":
            st.session_state.autenticado_seo = True
        else:
            st.session_state.autenticado_seo = False
            st.error("❌ Contraseña incorrecta")
    
    # --- LA MAGIA DE LA SEGURIDAD ---
    if st.session_state.autenticado_seo:
        try:
            # Trae la clave real oculta desde la configuración de Streamlit
            api_key_input = st.secrets["GEMINI_API_KEY"]
            st.success("✅ Sistema Desbloqueado")
        except KeyError:
            st.error("⚠️ Falta configurar la API Key secreta en Streamlit Cloud.")
            api_key_input = None
    else:
        api_key_input = None

    st.markdown("---")
    st.markdown("**Módulos Activos:**")
    st.markdown("- 📝 Redactor SEO (Long-form)")
    st.markdown("- 🎯 Creador de Ads (Short-form)")

# --- PESTAÑAS DE LA APLICACIÓN ---
tab1, tab2 = st.tabs(["📝 Redactor SEO", "🎯 Creador de Ads"])

# ==========================================
# PESTAÑA 1: REDACTOR SEO
# ==========================================
with tab1:
    st.header("Generador de Artículos Optimizados")
    st.markdown("Crea artículos listos para posicionar, sin frases cliché y con meta etiquetas incluidas.")
    
    col1, col2 = st.columns(2)
    with col1:
        keyword = st.text_input("Palabra clave principal:", placeholder="Ej: mejores tarjetas de crédito")
        longitud = st.selectbox("Longitud aproximada:", ["Corta (500 palabras)", "Media (800 palabras)", "Larga (1200+ palabras)"])
    with col2:
        tono = st.text_input("Tono de voz / Marca:", placeholder="Ej: Formal, educativo, estilo banco")
    
    if st.button("Generar Artículo SEO", type="primary"):
        if not st.session_state.autenticado_seo or not api_key_input:
            st.error("⚠️ Por favor, ingresa tu contraseña corporativa en la barra lateral para acceder al sistema.")
        elif not keyword:
            st.warning("⚠️ Debes ingresar una palabra clave.")
        else:
            with st.spinner('DuoLogic IA está investigando y redactando...'):
                try:
                    client = genai.Client(api_key=api_key_input)
                    prompt_seo = f"""
                    Actúa como un Redactor SEO Senior de la agencia Eleven. Escribe un artículo sobre "{keyword}".
                    Tono: {tono}. Longitud: {longitud}.
                    
                    REGLAS ESTRICTAS: 
                    1. Párrafos cortos y uso de listas. Cero frases IA cliché ("En conclusión", "En resumen", etc.).
                    2. PROHIBIDO el relleno conversacional. NO incluyas saludos, ni te presentes (ej: "Hola, soy...", "Aquí tienes el artículo").
                    3. El texto DEBE iniciar directamente con el título principal H1 (#).
                    4. Al final del artículo, añade obligatoriamente una línea divisoria (---) y el título exacto "⚙️ META ETIQUETAS SEO".
                    
                    Incluye debajo de ese título 3 opciones de Meta Títulos (máx 60 carac.) y 3 Meta Descripciones (máx 155 carac.).
                    Formato: Markdown estricto.
                    """
                    respuesta = client.models.generate_content(model='gemini-2.5-flash', contents=prompt_seo)
                    
                    st.success("¡Artículo generado con éxito!")
                    st.markdown("---")
                    st.markdown(respuesta.text)
                    
                    # --- BOTÓN DE DESCARGA (SEO) ---
                    st.markdown("---")
                    nombre_archivo = keyword.replace(" ", "_").lower() + ".md"
                    st.download_button(
                        label="📥 Descargar Artículo (.md)",
                        data=respuesta.text,
                        file_name=nombre_archivo,
                        mime="text/markdown"
                    )
                    
                except Exception as e:
                    st.error(f"Error al conectar con la IA: {e}")

# ==========================================
# PESTAÑA 2: CREADOR DE ADS
# ==========================================
with tab2:
    st.header("Generador de Campañas (Paid Media)")
    st.markdown("Crea grillas de anuncios con límites de caracteres perfectos para Google y Meta Ads.")
    
    col3, col4 = st.columns(2)
    with col3:
        producto = st.text_input("Producto o Servicio:", placeholder="Ej: Curso Online de Excel")
        oferta = st.text_input("Oferta o Diferenciador:", placeholder="Ej: 50% de descuento por 48hs")
    with col4:
        publico = st.text_input("Público Objetivo:", placeholder="Ej: Analistas junior")
        
    if st.button("Generar Campaña de Ads", type="primary"):
         if not st.session_state.autenticado_seo or not api_key_input:
            st.error("⚠️ Por favor, ingresa tu contraseña corporativa en la barra lateral para acceder al sistema.")
         elif not producto:
            st.warning("⚠️ Debes ingresar un producto.")
         else:
            with st.spinner('Aplicando fórmulas de persuasión (AIDA/PAS)...'):
                try:
                    client = genai.Client(api_key=api_key_input)
                    prompt_ads = f"""
                    Actúa como Media Buyer de la agencia Eleven. Crea copys para: {producto}.
                    Público: {publico}. Oferta: {oferta}.
                    
                    REGLAS ESTRICTAS:
                    1. PROHIBIDO el relleno conversacional. NO incluyas saludos ni te presentes (ej: "¡Absolutamente!", "Aquí tienes las propuestas", etc.).
                    2. Empieza el texto directamente con el título "## 1. Google Ads".
                    
                    Entrega:
                    ## 1. Google Ads
                    5 Títulos (máx 30 caracteres) y 3 Descripciones (máx 90 caracteres). Indica la longitud de cada uno al final entre paréntesis.
                    
                    ## 2. Meta Ads
                    2 opciones largas de copy para el feed (Opción A: Fórmula AIDA, Opción B: Storytelling).
                    
                    Formato: Markdown estricto.
                    """
                    respuesta = client.models.generate_content(model='gemini-2.5-flash', contents=prompt_ads)
                    
                    st.success("¡Campaña generada con éxito!")
                    st.markdown("---")
                    st.markdown(respuesta.text)
                    
                    # --- BOTÓN DE DESCARGA (ADS) ---
                    st.markdown("---")
                    nombre_archivo_ads = "ADS_" + producto.replace(" ", "_").lower() + ".md"
                    st.download_button(
                        label="📥 Descargar Campaña (.md)",
                        data=respuesta.text,
                        file_name=nombre_archivo_ads,
                        mime="text/markdown"
                    )
                    
                except Exception as e:
                    st.error(f"Error al conectar con la IA: {e}")