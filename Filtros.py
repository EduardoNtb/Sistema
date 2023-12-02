import pickle
from pathlib import Path
import pandas as pd
#import plotly.express as px
import streamlit as st
import streamlit_authenticator as stauth
import os
import io
#from PIL import Image
import datetime
#import subprocess
#import smtplib
#import time 
from docx import Document
from docxtpl import DocxTemplate
import base64
#from docx2pdf import convert
import win32com.client

win32com.client.Dispatch("WScript.Shell")
fecha_actual = datetime.datetime.now().strftime("%d/%m/%Y")
st.set_page_config(page_title='Sistema', page_icon='🌍', layout='wide')
names = ['Salvador Jair Ocampo','Rodrigo Manzano']
usernames = ['jocampo', 'rmanzano']
file_path = Path(__file__).parent / 'hashed_pw.pkl'
with file_path.open('rb') as file:
    hashed_passwords = pickle.load(file)

authenticator = stauth.Authenticate(names, usernames, hashed_passwords,
                                    'principal_dashboard', 'abcdef', cookie_expiry_days=2/24)  # 2 horas


name, authentication_status, username = authenticator.login('Login','main')

if authentication_status == False:
    st.error('El Usuario/Constraseña es incorrecta')

if authentication_status == None:
    st.warning('Complete todo los campos')

if authentication_status:
    imagen = "Foto4.jpg"
    imagen2 = 'Foto2.jpg'
    st.markdown(
    f'<div style="display: flex; justify-content: space-between;">'
    f'    <img src="data:image2/png;base64,{base64.b64encode(open(imagen2, "rb").read()).decode()}" '
    f'        style="float: left; margin-right: 10px; margin-top: 10px;" />'
    f'    <img src="data:image/png;base64,{base64.b64encode(open(imagen, "rb").read()).decode()}" '
    f'        style="float: right; margin-right: 10px; margin-top: 10px;" />'
    f'</div>',
    unsafe_allow_html=True
    )

    #st.sidebar.success("Selecciona la Opcion Arriba")
    
    st.sidebar.title(f'Bienvenido {name}')
    authenticator.logout('Logout', 'sidebar')
    
    st.subheader('Actualizacion de archivos')

    new_word_file = st.file_uploader("Cargar o Reemplazar Archivo Word Docentes ", type=["docx"])

    if new_word_file is not None:
        # Procesar el nuevo archivo Word cargado y reemplazar el archivo existente
        with st.spinner('Procesando el nuevo archivo, por favor espera...'):
            # Leer el nuevo archivo Word
            new_doc = Document(new_word_file)

            new_file_path = "docentes.docx"
            new_doc.save(new_file_path)
            
            st.success('Archivo Word actualizado exitosamente.')

    new_word_file = st.file_uploader("Cargar o Reemplazar Archivo Word Tutores ", type=["docx"])

    if new_word_file is not None:
        # Procesar el nuevo archivo Word cargado y reemplazar el archivo existente
        with st.spinner('Procesando el nuevo archivo, por favor espera...'):
            # Leer el nuevo archivo Word
            new_doc = Document(new_word_file)

            new_file_path = "tutorados.docx"
            new_doc.save(new_file_path)
            
            st.success('Archivo Word actualizado exitosamente.')

    st.header('Filtrado del Excel')
    st.subheader('Vistas')
    # carga de excel
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx", "xls","csv"])

    # Si se carga un archivo, cargarlo en un DataFrame de Pandas
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file).dropna()
    else:
        # Si no se carga un archivo, usar un DataFrame vacío
        df = pd.DataFrame()

    if not df.empty:
        st.markdown('---')
        st.subheader('Excel Original')
        st.dataframe(df)

        st.sidebar.header('Filtros:')

        profesor = df['Docente'].unique().tolist()
        tutor = df['TUTOR'].unique().tolist()
        materia = df['Materia'].unique().tolist()
        curso = df['Curso'].unique().tolist()

        maestro_select = st.sidebar.multiselect('Docente:',
                                        profesor,
                                        default=profesor)

        tutor_select = st.sidebar.multiselect('Tutor:',
                                        tutor,
                                        default=tutor)

        materia_select = st.sidebar.multiselect('Materia:',
                                        materia,
                                        default=materia)

        curso_select = st.sidebar.multiselect('Curso:',
                                        curso,
                                        default=curso)

        st.markdown('---')
        
        mask = (df['Docente'].isin(maestro_select)) & (df['TUTOR'].isin(tutor_select)) & (df['Materia'].isin(materia_select)) & (df['Curso'].isin(curso_select)) 
        result = df[mask].shape[0]
        st.markdown(f'*Todal de alumnos que cumples las condiciones: {result}*')

        st.subheader('Excel filtrado')
        st.dataframe(df[mask])

        if st.button("Generar Archivos para Tutores"):
            with st.spinner('Generando archivos, por favor espera...'):
                for index, row in df[mask].iterrows():
                    template = DocxTemplate("tutorados.docx")
                    context = {
                        "nombre": row['Estudiante'],
                        "control": row['Control'],
                        "semestre": row['Semestre'],
                        "materia": row['Materia'],
                        "gpo": row['Grupo'],
                        "curso": row['Curso'],
                        "docente": row['Docente'],
                        "tutor": row['TUTOR'],
                        "date": fecha_actual
                    }
                    template.render(context)
                    output = io.BytesIO()
                    template.save(output)
                    download_link = f'[Descargar {row["Estudiante"]}](data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{base64.b64encode(output.getvalue()).decode()})'
                    # Mostrar el enlace de descarga
                    st.markdown(download_link, unsafe_allow_html=True)
            st.success('Archivos generados exitosamente.')

        if st.button("Generar Archivos para Docentes"):
            with st.spinner('Generando archivos, por favor espera...'):
                for index, row in df[mask].iterrows():
                    # Cargar la plantilla de Word
                    template = DocxTemplate("docentes.docx")
                    context = {
                        "nombre": row['Estudiante'],
                        "control": row['Control'],
                        "semestre": row['Semestre'],
                        "materia": row['Materia'],
                        "gpo": row['Grupo'],
                        "curso": row['Curso'],
                        "docente": row['Docente'],
                        "tutor": row['TUTOR'],
                        "date": fecha_actual
                    }
                    template.render(context)
                    output = io.BytesIO()
                    template.save(output)
                    download_link = f'[Descargar {row["Estudiante"]}](data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{base64.b64encode(output.getvalue()).decode()})'
                    st.markdown(download_link, unsafe_allow_html=True)
                st.success('Archivos generados exitosamente.')

        if st.button("Generar Archivos Excel por Docente"):
            with st.spinner('Generando archivos, por favor espera...'):
                # Añadir columnas vacías
                columns_to_save = ['Estudiante', 'Control', 'Semestre','Curso', 'Materia']
                
                # Iterate over unique values in 'Docente' column in the original DataFrame
                for docente_value in df['Docente'].unique():
                    # Filtrar por el valor actual de 'Docente'
                    docente_mask = df['Docente'] == docente_value
                    df_for_docente = df[docente_mask][columns_to_save]

                    # Generar un nombre de archivo único para el archivo Excel
                    excel_file = f"{docente_value}_filtered_data.xlsx"

                    # Guardar el DataFrame en un archivo Excel
                    df_for_docente.to_excel(excel_file, index=False)

                    # Crear el enlace de descarga para el archivo Excel
                    download_link_excel = f'[Descargar Excel de {docente_value}](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{base64.b64encode(open(excel_file, "rb").read()).decode()})'

                    # Mostrar el enlace de descarga
                    st.markdown(download_link_excel, unsafe_allow_html=True)

                    # Eliminar el archivo Excel después de mostrar el enlace (opcional)
                    os.remove(excel_file)
            st.success('Archivos generados exitosamente.')
