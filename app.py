import streamlit as st
import pandas as pd
import io

# Configuración básica de la página
st.set_page_config(page_title="Generador de Reportes - Terra", layout="centered")

st.title("Generador de Reporte: Cantidad de Pasajes por Origen")
st.write("Sube el archivo de resumen de ventas descargado de Terra para generar automáticamente el reporte agrupado por ciudad.")

# 1. Cargador de archivos
archivo_subido = st.file_uploader("Sube el archivo original (Ej: cruzdelsur_Resumen_ventas_...xlsx)", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        # Leer el archivo excel en memoria
        df_ventas = pd.read_excel(archivo_subido)
        
        # Validación de seguridad: Comprobar que el Excel tenga las columnas que necesitamos
        columnas_requeridas = ['Estado', 'Origen (ciudad)', 'Valor en moneda principal', 'Código pasaje']
        faltantes = [col for col in columnas_requeridas if col not in df_ventas.columns]
        
        if faltantes:
            st.error(f"❌ El archivo subido no parece ser el reporte de Terra. Faltan las siguientes columnas: {', '.join(faltantes)}")
        else:
            # 2. Filtrar únicamente los pasajes "Pagado"
            df_pagados = df_ventas[df_ventas['Estado'] == 'Pagado']
            
            # 3. Agrupar por ciudad y realizar los cálculos
            df_agrupado = df_pagados.groupby('Origen (ciudad)').agg(
                Valor_total=('Valor en moneda principal', 'sum'),
                Cantidad_de_pasajes=('Código pasaje', 'count')
            ).reset_index()
            
            # 4. Renombrar columnas para igualar al formato deseado
            df_agrupado = df_agrupado.rename(columns={
                'Origen (ciudad)': 'Ciudad',
                'Valor_total': 'Valor total',
                'Cantidad_de_pasajes': 'Cantidad de pasajes'
            })
            
            # Mostrar una pequeña vista previa en la pantalla
            st.subheader("Vista previa del reporte")
            st.dataframe(df_agrupado, use_container_width=True)
            
            # 5. Escribir el nuevo Excel en la memoria RAM del servidor (BytesIO)
            output = io.BytesIO()
            # Usamos xlsxwriter como motor para asegurar compatibilidad
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_agrupado.to_excel(writer, sheet_name='Cantidad de pasajes', index=False)
            
            st.success("✅ ¡El reporte se procesó con éxito! Ya puedes descargarlo.")
            
            # 6. Botón para que el usuario descargue el reporte final
            st.download_button(
                label="📥 Descargar Reporte Generado",
                data=output.getvalue(),
                file_name="Reporte cantidad de pasajes por origen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"⚠️ Ocurrió un error inesperado al leer el archivo: {e}")
