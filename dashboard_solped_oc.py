"""
Dashboard de SOLPED versus Ordenes de Compra (OC)
=================================================

Este módulo implementa una aplicación de Streamlit diseñada para ayudar a los
usuarios a controlar las solicitudes de pedido (SOLPED) y verificar si ya
cuentan o no con una orden de compra (OC). El objetivo es proporcionar
visibilidad y control sobre los registros de solicitud para que el equipo
pueda identificar fácilmente aquellas solicitudes que requieren seguimiento.

Instrucciones de uso
--------------------

1. Ejecute el script con `streamlit run dashboard_solped_oc.py` en su
   terminal. Asegúrese de tener instaladas las dependencias `streamlit` y
   `pandas`.
2. Desde la barra lateral de la aplicación, cargue su archivo Excel que
   contenga las columnas "Fecha Sol.", "SOLPED", "Descripción del
   Material", "Doc.Compra", "Proveedor", "Solicitante", "Fecha Mod.",
   "Cantidad", "Centro" y "Almacén". La plantilla proporcionada en este
   repositorio (`SOLPED_VS_OC.xlsx`) sirve como referencia.
3. Una vez cargados los datos, el panel mostrará métricas resumidas del
   número total de SOLPED, aquellas que tienen OC asociada y aquellas que
   no la tienen. También se muestran filtros interactivos por solicitante,
   centro y estado de OC (con o sin orden), así como una tabla con el
   resultado filtrado y un botón para descargar el resultado en CSV.

Esta versión también permite cargar automáticamente la hoja `REVISION_SOLPED`
de un documento de Google Sheets, definido mediante las constantes
``DEFAULT_SHEET_ID`` y ``DEFAULT_GID`` en este archivo. Si el documento es
público, la aplicación leerá los datos sin que el usuario deba subir un
archivo o ingresar identificadores manualmente. Además se incluyen secciones
de análisis adicionales para las solicitudes sin orden de compra (OC), con
gráficos que resumen los solicitantes implicados, la evolución por fechas
de solicitud y modificación, y la distribución de las cantidades pedidas.

Autor: Asistente ChatGPT
Fecha: 20 de febrero de 2026
"""

import io
from typing import Tuple, Optional

import pandas as pd
import streamlit as st

# Identificadores por defecto para la hoja de cálculo de Google. Al definir
# estos valores, la aplicación puede obtener los datos automáticamente sin
# necesidad de que el usuario ingrese manualmente el ID y el GID de la pestaña.
# Para personalizar, reemplace los valores de DEFAULT_SHEET_ID y DEFAULT_GID.
DEFAULT_SHEET_ID = "1MT28ElFN2_nEPBc8sgKfqe7toWoht2ng"
DEFAULT_GID = "220782066"


def load_solped_data(file_like: io.BytesIO) -> pd.DataFrame:
    """Carga y transforma los datos de SOLPED.

    El archivo Excel generado por el sistema SAP suele contener dos filas
    vacías al inicio y una fila con nombres de columna. Esta función toma
    el archivo cargado por el usuario, extrae la fila 3 (índice 2) como
    cabecera y devuelve un DataFrame con columnas limpias. Además se crea
    una columna adicional llamada ``Tiene OC`` que indica si la solicitud
    tiene una orden de compra asociada (valor ``Con OC``) o no la tiene
    (valor ``Sin OC``). Se considera que no tiene OC cuando el campo
    ``Doc.Compra`` contiene textos como ``(en blanco)``, ``nan`` o está
    vacío.

    Args:
        file_like: Un objeto de tipo BytesIO que representa el archivo
            Excel cargado a través de Streamlit.

    Returns:
        DataFrame con los datos transformados y la columna ``Tiene OC``.
    """
    # Leer el archivo sin cabecera para poder identificar la fila de nombres
    raw = pd.read_excel(file_like, header=None)
    # Tomar la tercera fila como cabecera (índice 2 en cero-basado)
    header = raw.iloc[2].tolist()
    data = raw.iloc[3:].reset_index(drop=True)
    data.columns = header
    # Limpiar la columna Doc.Compra
    data['Doc.Compra'] = data['Doc.Compra'].astype(str).str.strip()
    # Determinar si existe OC
    sin_oc_mask = data['Doc.Compra'].isin(['(en blanco)', 'nan', '', 'None'])
    data['Tiene OC'] = sin_oc_mask.map({True: 'Sin OC', False: 'Con OC'})
    return data


def load_solped_from_google(sheet_id: str, gid: str) -> Optional[pd.DataFrame]:
    """Carga los datos desde una hoja de cálculo de Google.

    Utiliza el formato de exportación de Google Sheets en CSV para obtener
    directamente los datos de la pestaña especificada. Para que esto funcione,
    la hoja debe estar configurada con permisos de acceso público o bien
    compartida con la cuenta de servicio que se utilice.

    Google permite exportar una pestaña específica de un documento
    mediante la siguiente URL:

    ``https://docs.google.com/spreadsheets/d/<ID>/export?format=csv&id=<ID>&gid=<GID>``

    Donde ``<ID>`` es el identificador del documento y ``<GID>`` es el
    identificador de la pestaña. Estos parámetros se observan en la URL del
    documento cuando se navega entre pestañas.

    Documentación de referencia: según Ben Collins, se pueden añadir
    parámetros ``gid`` al final del enlace de exportación para seleccionar la
    pestaña deseada【723629888661203†L160-L176】【723629888661203†L270-L279】.

    Args:
        sheet_id: Identificador único del archivo de Google Sheets.
        gid: Identificador único de la pestaña dentro del documento.

    Returns:
        DataFrame con los datos cargados o ``None`` si la descarga falla.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        f"&id={sheet_id}&gid={gid}"
    )
    try:
        data = pd.read_csv(url)
    except Exception as exc:
        return None
    # Si la primera fila es el encabezado real, se devuelve directamente
    # Algunas hojas pueden incluir filas vacías iniciales; en tal caso se puede
    # reutilizar la función de transformación `load_solped_data` convirtiendo
    # el DataFrame en un buffer de Bytes e invocándola.
    # Aquí asumimos que el encabezado es correcto.
    # Crear columna Tiene OC
    if 'Doc.Compra' in data.columns:
        data['Doc.Compra'] = data['Doc.Compra'].astype(str).str.strip()
        sin_oc_mask = data['Doc.Compra'].isin(['(en blanco)', 'nan', '', 'None'])
        data['Tiene OC'] = sin_oc_mask.map({True: 'Sin OC', False: 'Con OC'})
    return data


def compute_metrics(data: pd.DataFrame) -> Tuple[int, int, int]:
    """Calcula métricas básicas a partir del DataFrame.

    Args:
        data: DataFrame de solicitudes procesadas.

    Returns:
        Una tupla con el total de registros, el número de solicitudes con
        orden de compra y el número de solicitudes sin orden de compra.
    """
    total = len(data)
    con_oc = (data['Tiene OC'] == 'Con OC').sum()
    sin_oc = (data['Tiene OC'] == 'Sin OC').sum()
    return total, con_oc, sin_oc


def main() -> None:
    """Punto de entrada principal de la aplicación Streamlit."""
    st.set_page_config(page_title='Dashboard SOLPED vs OC', layout='wide')
    st.title('Dashboard SOLPED vs Ordenes de Compra (OC)')
    st.write(
        'Esta aplicación le ayuda a identificar qué solicitudes de pedido '
        '(SOLPED) cuentan con una orden de compra (OC) y cuáles no. Puede '
        'filtrar por solicitante, centro y estado de OC, visualizar la '
        'información en una tabla y descargar los resultados filtrados.'
    )

    # Opciones de origen de datos: URL predefinida de Google Sheet o carga manual
    st.sidebar.header('Origen de datos')
    source_option = st.sidebar.radio(
        'Seleccione el origen de datos:',
        options=['Google Sheet (predefinido)', 'Archivo local', 'Google Sheet personalizado'],
        index=0
    )

    data: Optional[pd.DataFrame] = None
    if source_option == 'Google Sheet (predefinido)':
        # Utilizar los identificadores por defecto definidos arriba
        data = load_solped_from_google(DEFAULT_SHEET_ID, DEFAULT_GID)
        if data is None:
            st.error(
                'No se pudieron descargar los datos desde la hoja predefinida. '
                'Verifique que el documento sea público o ajuste DEFAULT_SHEET_ID y DEFAULT_GID.'
            )
            return
    elif source_option == 'Archivo local':
        st.sidebar.subheader('Cargar archivo Excel')
        uploaded_file = st.sidebar.file_uploader(
            label='Sube tu archivo Excel (p. ej. SOLPED_VS_OC.xlsx)',
            type=['xlsx', 'xls']
        )
        if uploaded_file is not None:
            try:
                data = load_solped_data(uploaded_file)
            except Exception as e:
                st.error(f'No se pudo leer el archivo: {e}')
                return
    else:  # Google Sheet personalizado
        st.sidebar.subheader('Cargar desde Google Sheets')
        sheet_id = st.sidebar.text_input(
            'ID del documento',
            value=DEFAULT_SHEET_ID,
            help='El identificador que aparece en la URL después de "/spreadsheets/d/"'
        )
        gid = st.sidebar.text_input(
            'GID de la pestaña',
            value=DEFAULT_GID,
            help='El parámetro "gid" que aparece al final de la URL cuando seleccionas la pestaña deseada'
        )
        if sheet_id and gid:
            data = load_solped_from_google(sheet_id, gid)
            if data is None:
                st.error('No se pudieron descargar los datos. Verifique que el documento sea público y que los identificadores sean correctos.')
                return

    # Si se cargaron datos correctamente, mostrar contenido y filtros
    if data is not None:
        # Asegurar que exista la columna 'Tiene OC'
        if 'Tiene OC' not in data.columns and 'Doc.Compra' in data.columns:
            data['Doc.Compra'] = data['Doc.Compra'].astype(str).str.strip()
            sin_oc_mask = data['Doc.Compra'].isin(['(en blanco)', 'nan', '', 'None'])
            data['Tiene OC'] = sin_oc_mask.map({True: 'Sin OC', False: 'Con OC'})

        total, con_oc, sin_oc = compute_metrics(data)
        col1, col2, col3 = st.columns(3)
        col1.metric('Total SOLPED', total)
        col2.metric('Con OC', con_oc)
        col3.metric('Sin OC', sin_oc)

        counts = data['Tiene OC'].value_counts().rename_axis('Estado').reset_index(name='Cantidad')
        st.subheader('Distribución de solicitudes con y sin OC')
        st.bar_chart(data=counts.set_index('Estado'))

        # Filtros comunes
        st.sidebar.header('Filtros')
        solicitante_col = 'Solicitante' if 'Solicitante' in data.columns else None
        centro_col = 'Centro' if 'Centro' in data.columns else None
        solicitantes = sorted(data[solicitante_col].dropna().unique().tolist()) if solicitante_col else []
        selected_solicitantes = st.sidebar.multiselect(
            'Solicitante', options=solicitantes, default=solicitantes
        ) if solicitante_col else []
        centros = sorted(data[centro_col].dropna().unique().tolist()) if centro_col else []
        selected_centros = st.sidebar.multiselect(
            'Centro', options=centros, default=centros
        ) if centro_col else []
        estado_oc = st.sidebar.radio(
            'Estado de OC', options=['Todos', 'Con OC', 'Sin OC'], index=0
        )

        # Aplicar filtros
        filtered_data = data.copy()
        if solicitante_col and selected_solicitantes:
            filtered_data = filtered_data[filtered_data[solicitante_col].isin(selected_solicitantes)]
        if centro_col and selected_centros:
            filtered_data = filtered_data[filtered_data[centro_col].isin(selected_centros)]
        if estado_oc != 'Todos':
            filtered_data = filtered_data[filtered_data['Tiene OC'] == estado_oc]

        st.subheader('Detalle de SOLPED filtradas')
        st.dataframe(filtered_data, use_container_width=True)

        csv = filtered_data.to_csv(index=False).encode('utf-8')
        st.download_button(
            label='Descargar datos filtrados en CSV',
            data=csv,
            file_name='solped_filtrado.csv',
            mime='text/csv'
        )

        # ----- Análisis específico para SOLPED sin OC -----
        # Este apartado genera gráficos y tablas adicionales para las
        # solicitudes que no tienen orden de compra relacionada. La intención
        # es proporcionar pistas sobre cuáles solicitantes, fechas o cantidades
        # requieren mayor atención.
        if 'Sin OC' in data['Tiene OC'].unique():
            missing = data[data['Tiene OC'] == 'Sin OC'].copy()
            st.subheader('Análisis de SOLPED sin OC')

            # Gráfica por solicitante
            if 'Solicitante' in missing.columns:
                counts_solic = missing['Solicitante'].value_counts().reset_index()
                counts_solic.columns = ['Solicitante', 'Cantidad']
                if not counts_solic.empty:
                    st.markdown('**Solicitudes sin OC por solicitante**')
                    st.bar_chart(counts_solic.set_index('Solicitante'))

            # Conversión de fechas de mod y solicitud si existen
            # Fecha de modificación
            if 'Fecha Mod.' in missing.columns:
                try:
                    missing['Fecha Mod.'] = pd.to_datetime(missing['Fecha Mod.'], errors='coerce', dayfirst=True)
                    counts_fmod = (
                        missing.dropna(subset=['Fecha Mod.'])
                        .groupby(pd.Grouper(key='Fecha Mod.', freq='M'))
                        .size()
                        .reset_index(name='Cantidad')
                    )
                    if not counts_fmod.empty:
                        counts_fmod['Periodo'] = counts_fmod['Fecha Mod.'].dt.to_period('M').dt.to_timestamp()
                        st.markdown('**Evolución mensual de SOLPED sin OC (Fecha Mod.)**')
                        st.line_chart(counts_fmod.set_index('Periodo')['Cantidad'])
                except Exception:
                    pass

            # Fecha de solicitud
            if 'Fecha Sol.' in missing.columns:
                try:
                    missing['Fecha Sol.'] = pd.to_datetime(missing['Fecha Sol.'], errors='coerce', dayfirst=True)
                    counts_fsol = (
                        missing.dropna(subset=['Fecha Sol.'])
                        .groupby(pd.Grouper(key='Fecha Sol.', freq='M'))
                        .size()
                        .reset_index(name='Cantidad')
                    )
                    if not counts_fsol.empty:
                        counts_fsol['Periodo'] = counts_fsol['Fecha Sol.'].dt.to_period('M').dt.to_timestamp()
                        st.markdown('**Evolución mensual de SOLPED sin OC (Fecha Sol.)**')
                        st.line_chart(counts_fsol.set_index('Periodo')['Cantidad'])
                except Exception:
                    pass

            # Cantidad de pedido
            if 'Cantidad' in missing.columns:
                try:
                    # Convertir cantidad a numérica en caso de ser cadena
                    missing['Cantidad'] = pd.to_numeric(missing['Cantidad'], errors='coerce')
                    counts_qty = (
                        missing.dropna(subset=['Cantidad'])
                        .groupby('Cantidad')
                        .size()
                        .reset_index(name='Frecuencia')
                        .sort_values(by='Cantidad')
                    )
                    if not counts_qty.empty:
                        st.markdown('**Distribución de cantidades en SOLPED sin OC**')
                        st.bar_chart(counts_qty.set_index('Cantidad'))
                except Exception:
                    pass
    else:
        st.info('Seleccione un origen de datos y proporcione la información necesaria para cargar los registros.')


if __name__ == '__main__':
    main()
