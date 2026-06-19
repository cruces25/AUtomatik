import streamlit as st
import pdfplumber
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="PDF a Excel Converter", page_icon="📊", layout="centered")

st.title("📊 Convertidor de PDF a Excel")
st.write("Sube uno o varios archivos PDF y los uniremos en un solo archivo de Excel, donde cada PDF será una pestaña diferente.")

# Componente para subir múltiples archivos
uploaded_files = st.file_uploader("Elige tus archivos PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.success(f"¡{len(uploaded_files)} archivo(s) cargado(s) con éxito!")
    
    if st.button("Procesar y Generar Excel"):
        output = io.BytesIO()
        
        try:
            with st.spinner("Procesando los PDFs... Por favor espera."):
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    for uploaded_file in uploaded_files:
                        sheet_name = uploaded_file.name.replace(".pdf", "")[:31]
                        
                        pdf_text_lines = []
                        with pdfplumber.open(uploaded_file) as pdf:
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    pdf_text_lines.extend(text.split('\n'))
                        
                        if not pdf_text_lines:
                            pdf_text_lines = ["No se pudo extraer texto de este PDF (¿es escaneado/imagen?)"]
                        
                        df = pd.DataFrame(pdf_text_lines, columns=["Contenido del PDF"])
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                output.seek(0)
                st.balloons()
                st.success("¡Archivo Excel generado correctamente!")
                
                st.download_button(
                    label="📥 Descargar Archivo Excel",
                    data=output,
                    file_name="PDFs_Convertidos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {e}")
