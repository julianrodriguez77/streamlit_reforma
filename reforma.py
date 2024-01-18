import os
import streamlit as st
import pandas as pd
from st_aggrid import JsCode, AgGrid, GridOptionsBuilder,GridUpdateMode, DataReturnMode
import matplotlib.pyplot as plt
from fpdf import FPDF
import base64
import numpy as np
from tempfile import NamedTemporaryFile
from sklearn.datasets import load_iris
from datetime import datetime
import hydralit_components as hc
from PIL import Image
import gspread
import state


#para que reconosca la tabla xlsx
import pip
pip.main(["install", "openpyxl"])
#nombre de pagina
st.set_page_config(page_title= 'Reformas CPI',
                    page_icon='moneybag:',
                    layout='wide' )

#@st.cache_data()
def create_download_link(val, filename):
    b64 = base64.b64encode(val)  
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}.pdf">Download file</a>'

# Para la seccion 1 y 2
def agregar_columnas(df):
    df['Movimiento'] = 0
    df['TOTAL'] = df['Codificado'] + df['Movimiento']
    return df
# Para la seccion 2 segunda tabla
def agregar_column(dfd):
    dfd['Movimiento'] = 0
    dfd['TOTAL'] = dfd['Codificado'] + dfd['Movimiento']
    return dfd

def Inicio():
    #left_co, cent_co,last_co = st.columns(3)
    #with cent_co:
        #st.image(logo)
    #st.image("P_Invencible.png")
    #st.markdown("<h1 style='text-aling: center'> Reformas : </h1>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; background-color: #000045; color: #ece5f6'>Unidad DE PLANIFICACIÓN</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; background-color: #000045; color: #ece5f6'>Sistema de reformas</h4>", unsafe_allow_html=True)
    menu_data = [
    {'id': 1, 'label': "Información", 'key': "md_how_to", 'icon': "fa fa-home"},
    {'id': 2, 'label': "Documentación", 'key': "md_run_analysis"}
    #{'id': 3, 'label': "Document Collection", 'key': "md_document_collection"},
    #{'id': 4, 'label': "Semantic Q&A", 'key': "md_rag"}
    ]

    int(hc.nav_bar(
        menu_definition=menu_data,
        hide_streamlit_markers=False,
        sticky_nav=True,
        sticky_mode='pinned',
        override_theme={'menu_background': '#4c00a5'}
    ))
    st.header("**📖 Información general de la Aplicación para realizar reformas**")
    st.markdown("""
                    
                    General:

                    Las reformas al Plan Operativo Anual (POA) del Gobierno Autónomo Descentralizado de la Prefectura de Pichincha (GADPP) consisten en el cambio/modificación a las partidas presupuestarias, proyectos, y metas establecidas
                    por cada una de las Unidades/Unidades del GADPP.
                    Este proceso implica en líneas generales el procesamiento de información disponible en el sistema odoo, así como bases sueltas y el seguimiento a metas en la plataforma
                    de seguimiento, con el propósito de mantener una congruencia de la información para proceder aceptar dichos cambios/modificaciones. 
                    Para llevar a cabo de manera efectiva el proceso y diagnóstico, es crucial tener en cuenta tres elementos esenciales: la normativa, el proceso y las bases. 

                    Este aplicativo web se presenta como una estrategia efectiva para optimizar y agilizar el proceso de gestión de reformas en la institución, al proporcionar una plataforma accesible y dinámica que permitirá a los usuarios
                    navegar a través de los datos, realizar análisis en tiempo real y tomar decisiones informadas de manera eficiente reduciendo el trabajo manual y posibles errores humanos
             
                    Además, el diseño del aplicativo web presenta un panel interactivo y de fácil intuición, garantizando que las diversas Unidades
                    de la institución puedan utilizar la herramienta de manera eficiente
                    """)
        
    st.info("""
                Se describen con más detalle estos componentes en el manual de uso [Documentacion Reformas](https://docs.snowflake.com/). 
                Ademas esta incluida la información del sistema utilizado y sus beneficios.
                """)
        
    st.markdown("""
                    En el caso de tener algun tipo de problema comuniquese con la coordinacion de Planificación. 🚀
                
                    A continuación se detalla cada opción de reforma:

                """)
    
    
    st.header("🔄 Reforma interna", help="Reforma en la misma Unidad")
            
    create_tab, tips_tab = \
        st.tabs(["Resumen", "❄️Pasos"])
    with create_tab:
        st.markdown("""
                    **Interna**

                    En esta sección las reformas/movimientos que se solicitan se realizaran unicamente entre la misma Unidad.
                    Es decir, se puede modificar los proyectos sumando y restando el codificado o cambiar el nombre de las metas en la Unidad seleccionada.
                    Una ves modificado los valores y metas deseadas se procede a guardar la información esto generara automaticamente un archivo pdf con un codigo y con información de las 
                    modificaciones realizadas ya sea solo codificado, metas o los 2.
                    
                    """
                )
                    
    with tips_tab:
        st.markdown("""
                💡 **Pasos para realizar una reforma interna**
                    
                - Selecione la `Unidad` en la cual desea realizar las modificaciones, aparecera 2 tablas con la información de la Unidad: `Presupuesto` y `Metas`. 
                - En la primera tabla `Presupuesto` se puede editar los valores que afectan al `codificado`, con la columna `movimiento` tomando muy en cuenta que se puede restar valores a las partidas unicamente que tengan `saldo diponible` y sumar el valor restado a cualquier partida deseada. 
                - En la segunda tabla `Metas` se pueden realizar modificaciones a la ultima meta registrada, en la columna `nueva meta` se ingresa el nombre de la nueva meta.
                - En los widgets de la parte inferior tiene los totales del `codificado`, `nuevo codificado` y `movimiento`. Esta información permite verificar que la información se a ingresado correctamente ya que el codificado debe tener el mismo valor y el valor de movimiento simpre debe ser cero.
                - El boton de guardar información se activara si todo el proceso se encuentra bien realizado caso contrario no se podra descargar la informacion de los datos modificados.
                - Se descarga un archivo pdf del movimiento del presupuesto y de los cambios de las metas, si el documento se encuentra vacio es decir, que no se a realizado cambios sea en el presupuesto o en las metas.
                """)
                #- Utilize dynamic tables to simplify data transformation and avoid complex pipeline management, making them ideal for materializing query results from multiple base tables in ETL processes.
    st.header("🌀 Reforma Externa", help="Reforma entre Unidades")
            
    create_tab, tips_tab = \
        st.tabs(["Resumen", "❄️Pasos"])
    with create_tab:
        st.markdown("""
                    **Externa**

                    En esta sección las reformas/movimientos que se solicitan se realizaran unicamente entre la misma Unidad.
                    Es decir, se puede modificar los proyectos sumando y restando el codificado o cambiar el nombre de las metas en la Unidad seleccionada.
                    Una ves modificado los valores y metas deseadas se procede a guardar la información esto generara automaticamente un archivo pdf con un codigo y con información de las 
                    modificaciones realizadas ya sea solo codificado, metas o los 2.
                    
                    """
                )
                    
    with tips_tab:
        st.markdown("""
                💡 **Pasos para realizar una reforma externa**
                    
                - Selecione la `Unidad` en la cual desea realizar las modificaciones, aparecera 2 tablas con la información de la Unidad: `Presupuesto` y `Metas`. 
                - En la primera tabla `Presupuesto` se puede editar los valores que afectan al `codificado`, con la columna `movimiento` tomando muy en cuenta que se puede restar valores a las partidas unicamente que tengan `saldo diponible` y sumar el valor restado a cualquier partida deseada. 
                - En la segunda tabla `Metas` se pueden realizar modificaciones a la ultima meta registrada, en la columna `nueva meta` se ingresa el nombre de la nueva meta.
                - En los widgets de la parte inferior tiene los totales del `codificado`, `nuevo codificado` y `movimiento`. Esta información permite verificar que la información se a ingresado correctamente ya que el codificado debe tener el mismo valor y el valor de movimiento simpre debe ser cero.
                - El boton de guardar información se activara si todo el proceso se encuentra bien realizado caso contrario no se podra descargar la informacion de los datos modificados.
                - Se descarga un archivo pdf del movimiento del presupuesto y de los cambios de las metas, si el documento se encuentra vacio es decir, que no se a realizado cambios sea en el presupuesto o en las metas.
                """)
        
    st.header("➖ Liberación ", help="Reforma entre Unidades")
            
    create_tab, tips_tab = \
        st.tabs(["Resumen", "❄️Pasos"])
    with create_tab:
        st.markdown("""
                    **Externa**

                    En esta sección las reformas/movimientos que se solicitan se realizaran unicamente entre la misma Unidad.
                    Es decir, se puede modificar los proyectos sumando y restando el codificado o cambiar el nombre de las metas en la Unidad seleccionada.
                    Una ves modificado los valores y metas deseadas se procede a guardar la información esto generara automaticamente un archivo pdf con un codigo y con información de las 
                    modificaciones realizadas ya sea solo codificado, metas o los 2.
                    
                    """
                )
                    
    with tips_tab:
        st.markdown("""
                💡 **Pasos para realizar una reforma externa**
                    
                - Selecione la `Unidad` en la cual desea realizar las modificaciones, aparecera 2 tablas con la información de la Unidad: `Presupuesto` y `Metas`. 
                - En la primera tabla `Presupuesto` se puede editar los valores que afectan al `codificado`, con la columna `movimiento` tomando muy en cuenta que se puede restar valores a las partidas unicamente que tengan `saldo diponible` y sumar el valor restado a cualquier partida deseada. 
                - En la segunda tabla `Metas` se pueden realizar modificaciones a la ultima meta registrada, en la columna `nueva meta` se ingresa el nombre de la nueva meta.
                - En los widgets de la parte inferior tiene los totales del `codificado`, `nuevo codificado` y `movimiento`. Esta información permite verificar que la información se a ingresado correctamente ya que el codificado debe tener el mismo valor y el valor de movimiento simpre debe ser cero.
                - El boton de guardar información se activara si todo el proceso se encuentra bien realizado caso contrario no se podra descargar la informacion de los datos modificados.
                - Se descarga un archivo pdf del movimiento del presupuesto y de los cambios de las metas, si el documento se encuentra vacio es decir, que no se a realizado cambios sea en el presupuesto o en las metas.
                """)
        
    st.header("➕ Solicitud", help="Reforma entre Unidades")
            
    create_tab, tips_tab = \
        st.tabs(["Resumen", "❄️Pasos"])
    with create_tab:
        st.markdown("""
                    **Externa**

                    En esta sección las reformas/movimientos que se solicitan se realizaran unicamente entre la misma Unidad.
                    Es decir, se puede modificar los proyectos sumando y restando el codificado o cambiar el nombre de las metas en la Unidad seleccionada.
                    Una ves modificado los valores y metas deseadas se procede a guardar la información esto generara automaticamente un archivo pdf con un codigo y con información de las 
                    modificaciones realizadas ya sea solo codificado, metas o los 2.
                    
                    """
                )
                    
    with tips_tab:
        st.markdown("""
                💡 **Pasos para realizar una reforma externa**
                    
                - Selecione la `Unidad` en la cual desea realizar las modificaciones, aparecera 2 tablas con la información de la Unidad: `Presupuesto` y `Metas`. 
                - En la primera tabla `Presupuesto` se puede editar los valores que afectan al `codificado`, con la columna `movimiento` tomando muy en cuenta que se puede restar valores a las partidas unicamente que tengan `saldo diponible` y sumar el valor restado a cualquier partida deseada. 
                - En la segunda tabla `Metas` se pueden realizar modificaciones a la ultima meta registrada, en la columna `nueva meta` se ingresa el nombre de la nueva meta.
                - En los widgets de la parte inferior tiene los totales del `codificado`, `nuevo codificado` y `movimiento`. Esta información permite verificar que la información se a ingresado correctamente ya que el codificado debe tener el mismo valor y el valor de movimiento simpre debe ser cero.
                - El boton de guardar información se activara si todo el proceso se encuentra bien realizado caso contrario no se podra descargar la informacion de los datos modificados.
                - Se descarga un archivo pdf del movimiento del presupuesto y de los cambios de las metas, si el documento se encuentra vacio es decir, que no se a realizado cambios sea en el presupuesto o en las metas.
                """)
    

def Interna():
    def main():
        #para conectar a google drive
        #gc = gspread.service_account(filename= 'reformas-402915-4d33dedfb202.json')
        #abrir el archivo de drive
        #sh = gc.open('reforma')
        #mt = gc.open('REFORMA_POA')
        #la hoja de sheets
        #wks = sh.get_worksheet(0)
        #mts = mt.get_worksheet(0)
        #para editar
        odoo = pd.read_excel("tabla_presupuesto.xlsx")
        #odoo = wks.get_all_records()
        metas = pd.read_excel("tabla_metas.xlsx")
        #metas = mts.get_all_records()
        #odoo = [{k: v.encode('utf-8') if isinstance(v, str) else v for k, v in row.items()} for row in odoo]

        df_odoo = pd.DataFrame(odoo)
        df_mt = pd.DataFrame(metas)
        st.markdown("<h1 style='text-align:center;background-color: #95b8d1; color: #ffffff'>🔄 REFORMA INTERNA </h1>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center; background-color: #f0efeb; color: #080200'>(En la misma Unidad)</h4>", unsafe_allow_html=True)
        st.markdown("---")
        #reload---
        reload_data = False
        #filtramos solo PAI
        df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
        df_odoo['Codificado'] = df_odoo['Codificado'].round(2)
        #agrupamos las Unidades
        direc = st.selectbox('Escoja la Unidad', options=df_odoo['Unidad'].unique())
        #filtrar columnas 
        df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()##.agg({'Codificado':'sum'},{'Saldo_Disponible':'sum'}) #
        df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
        df_mt= df_mt.loc[df_mt.Unidad == direc]
        #Creamos los Datos ha editar
        #data=df_od
        df = pd.DataFrame(df_od)
        df = agregar_columnas(df)
        st.markdown(f"<h2 style='text-align:center;'> {direc} </h2>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Presupuesto</h3>", unsafe_allow_html=True, help="Se realiza cambios en la columna movimientos, tomar encuenta el saldo disponible para restar un valor caso contrario no se tomara encuenta la reforma.")
        gb = GridOptionsBuilder.from_dataframe(df) 
        gb.configure_column('Unidad', hide=True)#, rowGroup=True, cellRenderer= "agGroupCellRenderer", )
        gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
        gb.configure_column('Estructura')
        gb.configure_column(field ='Codificado', column_name = "Codificado", maxWidth=150, aggFunc="sum", valueFormatter="data.Codificado.toLocaleString('en-US');")
        gb.configure_column('Saldo_Disponible', aggFunc='sum',maxWidth=150, valueFormatter="data.Saldo_Disponible.toLocaleString('en-US');" )
        cellsytle_jscode = JsCode("""
            function(params) {
                if (params.value > '0') {
                    return {
                        'color': 'white',
                        'backgroundColor': 'green'
                    }
                } 
                else if (params.value < '0'){
                    return {
                        'color': 'white',
                        'backgroundColor': 'darkred'
                    }
                }                 
                else {
                    return {
                        'color': 'black',
                        'backgroundColor': 'white'
                    }
                }
            };
        """)
       
        change = JsCode("""
            function isCellEditable(params){
                if(params.data.Saldo_Disponible >= 0 && params.data.Movimiento >= 0 ){
                    return true
                }
                else{
                    return false

                }
                
            }
        """)    
        gb.configure_column('Movimiento', editable= change ,type=['numericColumn'], cellStyle=cellsytle_jscode, aggFunc='sum',maxWidth=150, valueFormatter="data.Movimiento.toLocaleString('en-US');")
        
        gb.configure_column('Nuevo Codificado', valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                            editable=True, type=['numericColumn'], aggFunc='sum',maxWidth=150, valueFormatter="data.Nuevo Codificado.toLocaleString('en-US');")
        gb.configure_column('TOTAL', hide=True)

        go = gb.build()
        go['alwaysShowHorizontalScroll'] = True
        go['scrollbarWidth'] = 1
        reload_data = False
        #return_mode_value = DataReturnMode.FILTERED_AND_SORTED
        #update_mode_value = GridUpdateMode.GRID_CHANGED

        edited_df = AgGrid(
            df,
            editable= True,
            gridOptions=go,
            width=1000, 
            height=400, 
            fit_columns_on_grid_load=True,
            theme="streamlit",
            #data_return_mode=return_mode_value, 
            #update_mode=update_mode_value,
            allow_unsafe_jscode=True, 
            #key='an_unique_key_xZs151',
            reload_data=reload_data,
            #no agregar cambia la columna de float a str
            #try_to_convert_back_to_original_types=False
        )
        st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Metas </h3>", unsafe_allow_html=True, help="En la columna Nueva Meta se puede agregar la modificación a la meta actual")
        # Mostrar la tabla con la extensión st_aggrid
        edit_df = AgGrid(df_mt, editable=True)
                        #reload_data=reload_data,)
        edit_df = pd.DataFrame(edit_df['data'])

        # Si se detectan cambios, actualiza el DataFrame
        if edited_df is not None:
            # Convierte el objeto AgGridReturn a DataFrame
            edited_df = pd.DataFrame(edited_df['data'])
            edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
            #st.write('Tabla editada:', edited_df)
            #st.write(f'<div style="width: 100%; margin: auto;">{df.to_html(index=False)}</div>',unsafe_allow_html=True)
        

        #CERRAMOS LA SECCIÓN
        st.markdown("---")
        st.subheader("Totales ") 
        total_cod = int(edited_df['Codificado'].sum())
        total_mov = int(edited_df['Movimiento'].sum())
        total_tot = int(edited_df['TOTAL'].sum())

        #DIVIDIMOS EL TABLERO EN 3 SECCIONES
        left_column, center_column, right_column = st.columns(3)
        with left_column:
            st.subheader("Total Codificado: ")
            st.subheader(f"US $ {total_cod:,}")
            #st.dataframe(df_od.stack())
            #st.write(df_od)
        
        with center_column:
            st.subheader("Nuevo Codificado: ")
            st.subheader(f"US $ {total_tot:,}")

        with right_column:
            st.subheader("Total movimiento: ")
            st.subheader(f"US $ {total_mov:,}")


        if total_cod != total_tot:
            st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor total del codificado y nuevo codificado son diferentes</h3></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor total del codificado y nuevo codificado son iguales</h3></div>', unsafe_allow_html=True)

        try:
            edited_rows = edited_df[edited_df['Movimiento'] != 0]
            edit_rows = edit_df[edit_df['Nueva Meta'] != '-']
        except:
            st.write('No se realizaron cambios en de la información')

        st.markdown("---")
        #st.markdown(type(Codificado))
        export_as_pdf = st.button("Guardar y Descargar")
        #Creamos una nueva tabla para el presupuesto
        columnas_filtradas = ['PROYECTO','Código','Estructura','Movimiento']
        nuevo_df = edited_rows[columnas_filtradas]
        nuevo_df['Proyecto, Partida Presupuestaria, Estructura'] = nuevo_df['PROYECTO'] + ' | ' + nuevo_df['Código'] + ' | ' + nuevo_df['Estructura']
        columnas_filtradas2 = ['Proyecto, Partida Presupuestaria, Estructura','Movimiento']
        nuevo_df = nuevo_df[columnas_filtradas2]
        sum_row = nuevo_df[['Movimiento']].sum()
        # Agregar la fila 'Total' al DataFrame mostrado en AGGrid
        total_row = pd.DataFrame([['Total', sum_row['Movimiento']]], 
                            columns=['Proyecto, Partida Presupuestaria, Estructura','Movimiento'], index=['Total'])
        nuevo_df = pd.concat([nuevo_df, total_row])
        #Creamos una nueva tabla para las metas
        meta_filtro = ['Proyecto','Metas','Nueva Meta']
        ed_df = edit_rows[meta_filtro]

        if total_cod == total_tot:
            if export_as_pdf:
                now = datetime.now()
                fecha_hora = now.strftime("%Y%m%d%H%M")
                excel_file = f'presupuesto{fecha_hora}.xlsx'
                excel_file2 = f'metas{fecha_hora}.xlsx'        
                st.write('Descargando... ¡Espere un momento!')
                edited_rows.to_excel(excel_file, index= False)
                #edited_rows.to_csv('data.csv', index=False)
                edit_rows.to_excel(excel_file2, index= False)
                st.success(f'¡Archivo Excel "{excel_file}" descargado correctamente!')
                st.success(f'¡Archivo Excel "{excel_file2}" descargado correctamente!')

                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                image_path = "logo GadPP.png"
                image = Image.open(image_path)
                img_width, img_height = image.size

                # Definir el tamaño de la imagen en el PDF (puedes ajustar según sea necesario)
                pdf.image(image_path, x=10, y=10, w=38, h=0)

                # Obtener fecha y hora actual para el título
                pdf.set_title(f"Reforma Presupuesto - {fecha_hora}")
                # Escribir el título en el PDF
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 30, txt=f"Reforma Presupuesto - {fecha_hora}", ln=True, align="C")
                pdf.cell(200, 0, txt=f"{direc}", ln=True, align="C")
                pdf.ln(10)
                # Anchos de columna para el DataFrame en el PDF
                col_widths = [150, 20]  # Anchos de columna fijos
                # Obtener anchos de columna dinámicos basados en el contenido
                pdf.set_font("Arial", size=7)
                for i, col in enumerate(nuevo_df.columns):
                    pdf.cell(col_widths[i], 10, str(col), border=1, align='C')
                pdf.ln()

                for _, row in nuevo_df.iterrows():
                    a=0
                    b=0
                    for i, value in enumerate(row):
                        # Convertir el valor a string antes de la verificación
                        value = str(value)
                        if len(value) > 25:
                            x = pdf.get_x()  # Guardar la posición X actual
                            y = pdf.get_y()  # Guardar la posición Y actual
                            pdf.multi_cell(col_widths[i], 5, txt=value, border=1)
                            pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        #   b = len(value)/4
                        #elif a == 1:
                        #    x = pdf.get_x()  # Guardar la posición X actual
                        #    y = pdf.get_y()  # Guardar la posición Y actual
                        #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                        #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        else:
                            pdf.cell(col_widths[i], 10, txt=value, border=1, align='C')
                    pdf.ln()
                
                pdf.add_page()
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=f"Reforma metas - {fecha_hora}", ln=True, align="C")
                pdf.cell(200, 15, txt=f"{direc}", ln=True, align="C")
                pdf.ln(10)
                # Anchos de columna para el DataFrame en el PDF
                col_widths2 = [60, 60,60]  # Anchos de columna fijos
                # Obtener anchos de columna dinámicos basados en el contenido
                pdf.set_font("Arial", size=7)
                for i, col in enumerate(ed_df.columns):
                    pdf.cell(col_widths2[i], 14, str(col), border=1, align='C')
                pdf.ln()

                for _, row in ed_df.iterrows():
                    a=0
                    b=0
                    for i, value in enumerate(row):
                        # Convertir el valor a string antes de la verificación
                        value = str(value)
                        if len(value) > 25:
                            x = pdf.get_x()  # Guardar la posición X actual
                            y = pdf.get_y()  # Guardar la posición Y actual
                            pdf.multi_cell(col_widths2[i], 4, txt=value, border=1)
                            pdf.set_xy(x + col_widths2[i], y)  # Restablecer la posición XY
                        #    a = 1
                        #   b = len(value)/4
                        #elif a == 1:
                        #    x = pdf.get_x()  # Guardar la posición X actual
                        #    y = pdf.get_y()  # Guardar la posición Y actual
                        #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                        #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        else:
                            pdf.cell(col_widths2[i], 10, txt=value, border=1, align='C')
                    pdf.ln()
                


                # Guardar el PDF
                pdf_output = f"Reforma Presupuesto_{fecha_hora}.pdf"
                pdf.output(pdf_output)
                st.success(f"Se ha generado el PDF: {pdf_output}")
        else:
            st.markdown('<div style=" margin: 0 auto; background-color:#D9FDFF; padding:10px; text-align: center;"><h4 style="color:#01464A;">Los movimientos tienen incosistencia revisar para descargar.</h3></div>', unsafe_allow_html=True)
                

    if __name__ == '__main__':
        main()

def Externa():
    #para conectar a google drive
    
    st.markdown("<h1 style='text-align:center;background-color: #028d96; color: #ffffff'>🌀 REFORMA EXTERNA </h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; background-color: #f0efeb; color: #080200'>(Entre diferentes Unidades)</h4>", unsafe_allow_html=True)
    st.markdown("---")
    #htmlstr = f"<h2 style='background-color: #A7727D; color: #F9F5E7; border-radius: 7px; padding-left: 8px; text-align: center'> Selecione la Unidad donde se retira el movimiento</style></h2>"
    #st.markdown(htmlstr, unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; background-color: #f1f6f7; color: #080200'> Paso 1: Selecione la Unidad donde se retira el movimiento </h2>", unsafe_allow_html=True)
    #para conectar a google drive
    #gc = gspread.service_account(filename= 'reformas-402915-4d33dedfb202.json')
    #abrir el archivo de drive
    #sh = gc.open('reforma')
    #mt = gc.open('REFORMA_POA')
    #la hoja de sheets
    #wks = sh.get_worksheet(0)
    #mts = mt.get_worksheet(0)
    #para editar
    odoo = pd.read_excel("tabla_presupuesto.xlsx")
    #odoo = wks.get_all_records()
    metas = pd.read_excel("tabla_metas.xlsx")
    #metas = mts.get_all_records()
    df_odoo = pd.DataFrame(odoo)
    df_mt = pd.DataFrame(metas)
    df_mt2 = pd.DataFrame(metas)
    #df_odoo['Codificado'] = df_odoo['Codificado'].replace(",", ".").astype(float)
    #df_odoo['Saldo_Disponible'] = df_odoo['Saldo_Disponible'].replace(",", ".").astype(float)

    #reload---
    reload_data = False
    #filtramos solo PAI
    df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
    odf = df_odoo
    #agrupamos las Unidades
    direc = st.selectbox('Escoje la Unidad', options=df_odoo['Unidad'].unique())
    #filtrar columnas 
    df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()##.agg({'Codificado':'sum'},{'Saldo_Disponible':'sum'}) #
    df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mt= df_mt.loc[df_mt.Unidad == direc]

    #Creamos los Datos ha editar
    data1=df_od
    df = pd.DataFrame(data1)
    df = agregar_columnas(df)
    st.markdown(f"<h2 style='text-align:center;'> {direc} </h2>", unsafe_allow_html=True)
    st.title('')
    st.markdown(f"<h3 >Tabla de Presupuestos</h3>", unsafe_allow_html=True)
    gb = GridOptionsBuilder.from_dataframe(df) 
    gb.configure_column('Unidad', hide=True)#, rowGroup=True, cellRenderer= "agGroupCellRenderer", )
    gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
    gb.configure_column('Estructura')
    gb.configure_column('Codificado', header_name = "Codificado", aggFunc='sum')
    gb.configure_column('Saldo_Disponible', aggFunc='sum')
    cellsytle_jscode = JsCode("""
        function(params) {
            if (params.value > '0') {
                return {
                    'color': 'white',
                    'backgroundColor': 'green'
                }
            } 
            else if (params.value < '0'){
                return {
                    'color': 'white',
                    'backgroundColor': 'darkred'
                }
            }                 
            else {
                return {
                    'color': 'black',
                    'backgroundColor': 'white'
                }
            }
        };
    """)
    
    gb.configure_column('Movimiento', editable= True,type=['numericColumn'], cellStyle=cellsytle_jscode, aggFunc='sum')
    gb.configure_column('TOTAL2', valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                         editable=True, type=['numericColumn'], aggFunc='sum')
    gb.configure_column('TOTAL', hide=True)
    go = gb.build()
    reload_data = False
    #return_mode_value = DataReturnMode.FILTERED_AND_SORTED
    #update_mode_value = GridUpdateMode.GRID_CHANGED

    edited_df = AgGrid(
        df,
        editable= True,
        gridOptions=go,
        width=1000, 
        height=400, 
        fit_columns_on_grid_load=True,
        theme="streamlit",
        #data_return_mode=return_mode_value, 
        #update_mode=update_mode_value,
        allow_unsafe_jscode=True, 
        #key='an_unique_key_xZs151',
        reload_data=reload_data,
        #no agregar cambia la columna de float a str
        #try_to_convert_back_to_original_types=False
    )
    st.markdown(f"<h3 > Tabla de Metas </h3>", unsafe_allow_html=True)
    # Mostrar la tabla con la extensión st_aggrid
    edit_df = AgGrid(df_mt, editable=True,
                      reload_data=reload_data,)
    edit_df = pd.DataFrame(edit_df['data'])

    # Si se detectan cambios, actualiza el DataFrame
    if edited_df is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edited_df = pd.DataFrame(edited_df['data'])
        edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
        #st.write('Tabla editada:', edited_df)
        #st.write(f'<div style="width: 100%; margin: auto;">{df.to_html(index=False)}</div>',unsafe_allow_html=True)
    

    #CERRAMOS LA SECCIÓN
    st.markdown("---")
    st.subheader("Totales: ") 
    total_cod = int(edited_df['Codificado'].sum())
    total_mov = int(edited_df['Movimiento'].sum())
    total_tot = int(edited_df['TOTAL'].sum())

    #DIVIDIMOS EL TABLERO EN 3 SECCIONES
    left_column, center_column, right_column = st.columns(3)
    with left_column:
        st.subheader("Total Codificado: ")
        st.subheader(f"US $ {total_cod:,}")
        #st.dataframe(df_od.stack())
        #st.write(df_od)
    
    with center_column:
        st.subheader("Nuevo Codificado: ")
        st.subheader(f"US $ {total_tot:,}")

    with right_column:
        st.subheader("Total movimiento: ")
        st.subheader(f"US $ {total_mov:,}")


    if total_mov < 0:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">Se a restado el valor de:  ${total_mov:}</h3></div>', unsafe_allow_html=True)
    elif total_mov == 0:
        st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">No se a realizado ningun movimiento del codificado</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">Se a sumado el valor de:  ${total_mov:}</h3></div>', unsafe_allow_html=True)
        
    ##################################################
    #################     TABLA 2  ###################
    ################################################## 
    st.markdown("---")
    st.markdown("<h2 style='text-align: center; background-color: #f1f6f7; color: #080200'> Paso 2: Unidad a que se le asigna el movimiento  </h2>", unsafe_allow_html=True)
    
    # Obtener las opciones para el segundo selectbox excluyendo la opción seleccionada en el primero    
    opci = odf['Unidad'][odf['Unidad'] != direc]
    selec= st.selectbox('Escoja la Unidad donde se agregara el movimiento', options= opci.unique())
    #filtrar columnas 
    dfff= odf.loc[odf.Unidad == selec].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()
    dfff.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mtt= df_mt2.loc[df_mt2.Unidad == selec]
    #Creamos los Datos ha editar
    data2=dfff
    dfd2 = pd.DataFrame(data2)
    dfd2 = agregar_column(dfd2)
    st.markdown(f"<h2 style='text-align:center;'> {selec} </h2>", unsafe_allow_html=True)
    st.title('')
    st.markdown(f"<h3 > Tabla de Presupuestos</h3>", unsafe_allow_html=True)
   
    edi = AgGrid(
        dfd2,
        editable= True,
        gridOptions=go,
        width=1000, 
        height=400, 
        #fit_columns_on_grid_load=True,
        theme="streamlit",
        #data_return_mode=return_mode_value, 
        
        #update_mode=update_mode_value,
        allow_unsafe_jscode=True, 
        #key='an_unique_key_xZs151',
        reload_data=reload_data,
        #no agregar cambia la columna de float a str
        #try_to_convert_back_to_original_types=False
    )
    st.markdown(f"<h3> Tabla de Metas </h3>", unsafe_allow_html=True)
    # Mostrar la tabla con la extensión st_aggrid
    edit_dfd = AgGrid(df_mtt, 
                      editable=True,
                      #reload_data=reload_data
                      )
    edit_dfd = pd.DataFrame(edit_dfd['data'])

    # Si se detectan cambios, actualiza el DataFrame
    if edi is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edi = pd.DataFrame(edi['data'])
        edi['TOTAL'] = edi['Codificado'] + edi['Movimiento']
        #st.write('Tabla editada:', edited_df)
        #st.write(f'<div style="width: 100%; margin: auto;">{df.to_html(index=False)}</div>',unsafe_allow_html=True)
    
    
    #CERRAMOS LA SECCIÓN
    st.markdown("---")
    st.subheader("Totales: ") 
    total_cod2 = int(edi['Codificado'].sum())
    total_mov2 = int(edi['Movimiento'].sum())
    total_tot2 = int(edi['TOTAL'].sum())

    #DIVIDIMOS EL TABLERO EN 3 SECCIONES
    left_column, center_column, right_column = st.columns(3)
    with left_column:
        st.subheader("Total Codificado: ")
        st.subheader(f"US $ {total_cod2:,}")
        #st.dataframe(df_od.stack())
        #st.write(df_od)
    
    with center_column:
        st.subheader("Nuevo Codificado: ")
        st.subheader(f"US $ {total_tot2:,}")

    with right_column:
        st.subheader("Total movimiento: ")
        st.subheader(f"US $ {total_mov2:,}")


    if total_mov2 < 0:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">Se a restado el valor de:  ${total_mov2:}</h3></div>', unsafe_allow_html=True)
    elif total_mov2 == 0:
        st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">No se a realizado ningun movimiento del codificado</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">Se a sumado el valor de:  ${total_mov2:}</h3></div>', unsafe_allow_html=True)

    st.markdown("---")

    if total_mov2 != -total_mov:
        st.markdown(f'<div style="max-width: auto; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor del movimiento en las Unidades es diferente, por tanto no se puede proseguir. </h3></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="max-width: auto; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor del movimiento en las Unidades es el mismo.</h3></div>', unsafe_allow_html=True)

    try:
        edited_rows = edited_df[edited_df['Movimiento'] != 0]
        edited_rows2 = edi[edi['Movimiento'] != 0]
        edit_rows = edit_df[edit_df['Nueva Meta'] != '']
        edit_rows2 = edit_dfd[edit_dfd['Nueva Meta'] != '']
    except:
        st.write('No se realizaron cambios en la información')

    #if st.button('Descargar como Excel'):
    st.markdown("---")    
    export_as_pdf = st.button("Guardar y Descargar")
    #Creamos una nueva tabla para el presupuesto
    columnas_filtradas = ['PROYECTO','Código','Estructura','Movimiento']
    nuevo_df = edited_rows[columnas_filtradas]
    nuevo_df2 = edited_rows2[columnas_filtradas]
    nuevo_df['Proyecto, Partida Presupuestaria, Estructura'] = nuevo_df['PROYECTO'] + ' | ' + nuevo_df['Código'] + ' | ' + nuevo_df['Estructura']
    nuevo_df2['Proyecto, Partida Presupuestaria, Estructura'] = nuevo_df2['PROYECTO'] + ' | ' + nuevo_df2['Código'] + ' | ' + nuevo_df2['Estructura']
    columnas_filtradas2 = ['Proyecto, Partida Presupuestaria, Estructura','Movimiento']
    nuevo_df = nuevo_df[columnas_filtradas2]
    nuevo_df2 = nuevo_df2[columnas_filtradas2]
    sum_row = nuevo_df[['Movimiento']].sum()
    sum_row2 = nuevo_df2[['Movimiento']].sum()
    # Agregar la fila Total al DataFrame mostrado en AGGrid
    total_row = pd.DataFrame([['Total', sum_row['Movimiento']]], 
                         columns=['Proyecto, Partida Presupuestaria, Estructura','Movimiento'], index=['Total'])
    total_row2 = pd.DataFrame([['Total', sum_row2['Movimiento']]], 
                         columns=['Proyecto, Partida Presupuestaria, Estructura','Movimiento'], index=['Total'])
    
    nuevo_df = pd.concat([nuevo_df, total_row])
    nuevo_df2 = pd.concat([nuevo_df2, total_row2])
    #Creamos una nueva tabla para las metas
    meta_filtro = ['Proyecto','Metas','Nueva Meta']
    ed_df = edit_rows[meta_filtro]
    ed_df2 = edit_rows2[meta_filtro]

    if total_mov2 == -(total_mov):
        if export_as_pdf:
            now = datetime.now()
            fecha_hora = now.strftime("%Y%m%d%H%M")
            
            excel_file = f'presupuesto-{fecha_hora}.xlsx'
            excel_file2 = f'metas-{fecha_hora}.xlsx'
            excel_file3 = f'presupuesto2-{fecha_hora}.xlsx'
            excel_file4 = f'metas2-{fecha_hora}.xlsx'             
            st.write('Descargando... ¡Espere un momento!')
            edited_rows.to_excel(excel_file, index=False)
            edit_rows.to_excel(excel_file2, index=False)
            edited_rows2.to_excel(excel_file3, index=False)
            edit_rows2.to_excel(excel_file4, index=False)
            st.success(f'¡Archivo Excel "{excel_file}" descargado correctamente!')
            st.success(f'¡Archivo Excel "{excel_file2}" descargado correctamente!')

            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            # Obtener fecha y hora actual para el título
            pdf.set_title(f"Reforma Presupuesto - {fecha_hora}")
            # Escribir el título en el PDF
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, txt=f"Reforma Presupuesto - {fecha_hora}", ln=True, align="C")
            pdf.cell(200, 15, txt=f"De: {direc}", ln=True, align="C")
            pdf.ln(10)
            # Anchos de columna para el DataFrame en el PDF
            col_widths = [150, 20]  # Anchos de columna fijos
            # Obtener anchos de columna dinámicos basados en el contenido
            pdf.set_font("Arial", size=7)
            for i, col in enumerate(nuevo_df.columns):
                pdf.cell(col_widths[i], 10, str(col), border=1, align='C')
            pdf.ln()

            for _, row in nuevo_df.iterrows():
                a=0
                b=0
                for i, value in enumerate(row):
                    # Convertir el valor a string antes de la verificación
                    value = str(value)
                    if len(value) > 25:
                        x = pdf.get_x()  # Guardar la posición X actual
                        y = pdf.get_y()  # Guardar la posición Y actual
                        pdf.multi_cell(col_widths[i], 5, txt=value, border=1)
                        pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                    else:
                        pdf.cell(col_widths[i], 10, txt=value, border=1, align='C')
                pdf.ln()
            
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, txt=f"Reforma Presupuesto - {fecha_hora}", ln=True, align="C")
            pdf.cell(200, 15, txt=f"Para:  {selec}", ln=True, align="C")
            pdf.ln(10)
            # Anchos de columna para el DataFrame en el PDF
            #col_widths = [150, 20]  # Anchos de columna fijos
            # Obtener anchos de columna dinámicos basados en el contenido
            pdf.set_font("Arial", size=7)
            for i, col in enumerate(nuevo_df2.columns):
                pdf.cell(col_widths[i], 10, str(col), border=1, align='C')
            pdf.ln()

            for _, row in nuevo_df2.iterrows():
                a=0
                b=0
                for i, value in enumerate(row):
                    # Convertir el valor a string antes de la verificación
                    value = str(value)
                    if len(value) > 25:
                        x = pdf.get_x()  # Guardar la posición X actual
                        y = pdf.get_y()  # Guardar la posición Y actual
                        pdf.multi_cell(col_widths[i], 5, txt=value, border=1)
                        pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                    else:
                        pdf.cell(col_widths[i], 10, txt=value, border=1, align='C')
                pdf.ln()

            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, txt=f"Reforma metas - {fecha_hora}", ln=True, align="C")
            pdf.cell(200, 15, txt=f"{direc}", ln=True, align="C")
            pdf.ln(10)
            # Anchos de columna para el DataFrame en el PDF
            col_widths2 = [60, 60,60]  # Anchos de columna fijos
            # Obtener anchos de columna dinámicos basados en el contenido
            pdf.set_font("Arial", size=7)
            for i, col in enumerate(ed_df.columns):
                pdf.cell(col_widths2[i], 10, str(col), border=1, align='C')
            pdf.ln()

            for _, row in ed_df.iterrows():
                a=0
                b=0
                for i, value in enumerate(row):
                    # Convertir el valor a string antes de la verificación
                    value = str(value)
                    if len(value) > 25:
                        x = pdf.get_x()  # Guardar la posición X actual
                        y = pdf.get_y()  # Guardar la posición Y actual
                        pdf.multi_cell(col_widths2[i], 5, txt=value, border=1)
                        pdf.set_xy(x + col_widths2[i], y)  # Restablecer la posición XY
                    #    a = 1
                    #   b = len(value)/4
                    #elif a == 1:
                    #    x = pdf.get_x()  # Guardar la posición X actual
                    #    y = pdf.get_y()  # Guardar la posición Y actual
                    #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                    #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                    #    a = 1
                    else:
                        pdf.cell(col_widths2[i], 10, txt=value, border=1, align='C')
                pdf.ln()

            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, txt=f"Reforma metas - {fecha_hora}", ln=True, align="C")
            pdf.cell(200, 15, txt=f"{selec}", ln=True, align="C")
            pdf.ln(10)
            # Anchos de columna para el DataFrame en el PDF
            #col_widths2 = [60, 60,60]  # Anchos de columna fijos
            # Obtener anchos de columna dinámicos basados en el contenido
            pdf.set_font("Arial", size=7)
            for i, col in enumerate(ed_df2.columns):
                pdf.cell(col_widths2[i], 10, str(col), border=1, align='C')
            pdf.ln()

            for _, row in ed_df2.iterrows():
                a=0
                b=0
                for i, value in enumerate(row):
                    # Convertir el valor a string antes de la verificación
                    value = str(value)
                    if len(value) > 25:
                        x = pdf.get_x()  # Guardar la posición X actual
                        y = pdf.get_y()  # Guardar la posición Y actual
                        pdf.multi_cell(col_widths2[i], 5, txt=value, border=1)
                        pdf.set_xy(x + col_widths2[i], y)  # Restablecer la posición XY
                    else:
                        pdf.cell(col_widths2[i], 10, txt=value, border=1, align='C')
                pdf.ln()
            


            # Guardar el PDF
            pdf_output = f"Reforma Presupuesto_{fecha_hora}.pdf"
            pdf.output(pdf_output)
            st.success(f"Se ha generado el PDF: {pdf_output}")
    else:
        st.markdown('<div style=" margin: 0 auto; background-color:#D9FDFF; padding:10px; text-align: center;"><h4 style="color:#01464A;">Los movimientos tienen incosistencia revisar para descargar.</h3></div>', unsafe_allow_html=True)
        

def Liberación ():
        #para conectar a google drive
        #gc = gspread.service_account(filename= 'reformas-402915-4d33dedfb202.json')
        #abrir el archivo de drive
        #sh = gc.open('reforma')
        #mt = gc.open('REFORMA_POA')
        #la hoja de sheets
        #wks = sh.get_worksheet(0)
        #mts = mt.get_worksheet(0)
        #para editar
        odoo = pd.read_excel("tabla_presupuesto.xlsx")        
        #odoo = wks.get_all_records()
        metas = pd.read_excel("tabla_metas.xlsx")
        #metas = mts.get_all_records()
        #odoo = [{k: v.encode('utf-8') if isinstance(v, str) else v for k, v in row.items()} for row in odoo]

        df_odoo = pd.DataFrame(odoo)
        df_mt = pd.DataFrame(metas)
        #df_odoo['Codificado'] = df_odoo['Codificado'].replace(",", ".").astype(float)
        #df_odoo['Saldo_Disponible'] = df_odoo['Saldo_Disponible'].replace(",", ".").astype(float)
        st.markdown("<h1 style='text-align:center;background-color: #4d4c01; color: #ffffff'>➖ REFORMA LIBERACIÓN </h1>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center; background-color: #ebebda; color: #080200'>(Se solicita liberar presupuesto a la Institución)</h4>", unsafe_allow_html=True)
        st.markdown("---")
        #reload---
        reload_data = False
        #filtramos solo PAI
        df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
        #agrupamos las Unidades
        direc = st.selectbox('Escoje la Unidad', options=df_odoo['Unidad'].unique())
        #filtrar columnas 
        df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()##.agg({'Codificado':'sum'},{'Saldo_Disponible':'sum'}) #
        df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
        df_mt= df_mt.loc[df_mt.Unidad == direc]
        #Creamos los Datos ha editar
        data=df_od
        df = pd.DataFrame(data)
        df = agregar_columnas(df)
        st.markdown(f"<h2 style='text-align:center;'> {direc} </h2>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Presupuesto</h3>", unsafe_allow_html=True, help="Se realiza cambios en la columna movimientos, tomar encuenta el saldo disponible para restar un valor caso contrario no se tomara encuenta la reforma.")
        gb = GridOptionsBuilder.from_dataframe(df) 
        gb.configure_column('Unidad', hide=True)#, rowGroup=True, cellRenderer= "agGroupCellRenderer", )
        gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
        gb.configure_column('Estructura')
        gb.configure_column(field ='Codificado', column_name = "Codificado", maxWidth=150, aggFunc="sum")
        gb.configure_column('Saldo_Disponible', aggFunc='sum',maxWidth=150 )
        cellsytle_jscode = JsCode("""
            function(params) {
                if (params.value > '0') {
                    return {
                        'color': 'white',
                        'backgroundColor': 'green'
                    }
                } 
                else if (params.value < '0'){
                    return {
                        'color': 'white',
                        'backgroundColor': 'darkred'
                    }
                }                 
                else {
                    return {
                        'color': 'black',
                        'backgroundColor': 'white'
                    }
                }
            };
        """)
       
        change = JsCode("""
            function isCellEditable(params){
                if(params.data.Saldo_Disponible >= 0 && params.data.Movimiento >= 0 ){
                    return true
                }
                else{
                    return false

                }
                
            }
        """)    
        gb.configure_column('Movimiento', editable= change ,type=['numericColumn'], cellStyle=cellsytle_jscode, aggFunc='sum',maxWidth=150)
        
        gb.configure_column('Nuevo Codificado', valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                            editable=True, type=['numericColumn'], aggFunc='sum',maxWidth=150)
        gb.configure_column('TOTAL', hide=True)

        go = gb.build()
        go['alwaysShowHorizontalScroll'] = True
        go['scrollbarWidth'] = 1
        reload_data = False
        #return_mode_value = DataReturnMode.FILTERED_AND_SORTED
        #update_mode_value = GridUpdateMode.GRID_CHANGED

        edited_df = AgGrid(
            df,
            editable= True,
            gridOptions=go,
            width=1000, 
            height=400, 
            fit_columns_on_grid_load=True,
            theme="streamlit",
            #data_return_mode=return_mode_value, 
            #update_mode=update_mode_value,
            allow_unsafe_jscode=True, 
            #key='an_unique_key_xZs151',
            reload_data=reload_data,
            #no agregar cambia la columna de float a str
            #try_to_convert_back_to_original_types=False
        )
        st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Metas </h3>", unsafe_allow_html=True, help="En la columna Nueva Meta se puede agregar la modificación a la meta actual")
        # Mostrar la tabla con la extensión st_aggrid
        edit_df = AgGrid(df_mt, editable=True,
                        reload_data=reload_data,)
        edit_df = pd.DataFrame(edit_df['data'])

        # Si se detectan cambios, actualiza el DataFrame
        if edited_df is not None:
            # Convierte el objeto AgGridReturn a DataFrame
            edited_df = pd.DataFrame(edited_df['data'])
            edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
            #st.write('Tabla editada:', edited_df)
            #st.write(f'<div style="width: 100%; margin: auto;">{df.to_html(index=False)}</div>',unsafe_allow_html=True)
        

        #CERRAMOS LA SECCIÓN
        st.markdown("---")
        st.subheader("Totales ") 
        total_cod = int(edited_df['Codificado'].sum())
        total_mov = int(edited_df['Movimiento'].sum())
        total_tot = int(edited_df['TOTAL'].sum())

        #DIVIDIMOS EL TABLERO EN 3 SECCIONES
        left_column, center_column, right_column = st.columns(3)
        with left_column:
            st.subheader("Total Codificado: ")
            st.subheader(f"US $ {total_cod:,}")
            #st.dataframe(df_od.stack())
            #st.write(df_od)
        
        with center_column:
            st.subheader("Nuevo Codificado: ")
            st.subheader(f"US $ {total_tot:,}")

        with right_column:
            st.subheader("Total movimiento: ")
            st.subheader(f"US $ {total_mov:,}")


        if total_cod > total_tot:
            st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor total del codificado es mayor al nuevo codificado</h3></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor total del codificado es menor o igual al nuevo codificado por tanto no se esta liberando el presupuesto</h3></div>', unsafe_allow_html=True)

        try:
            edited_rows = edited_df[edited_df['Movimiento'] != 0]
            edit_rows = edit_df[edit_df['Nueva Meta'] != '']
        except:
            st.write('No se realizaron cambios en de la información')

        st.markdown("---")
        export_as_pdf = st.button("Guardar y Descargar")
        #Creamos una nueva tabla para el presupuesto
        columnas_filtradas = ['PROYECTO','Código','Estructura','Movimiento']
        nuevo_df = edited_rows[columnas_filtradas]
        nuevo_df['Proyecto, Partida Presupuestaria, Estructura'] = nuevo_df['PROYECTO'] + ' | ' + nuevo_df['Código'] + ' | ' + nuevo_df['Estructura']
        columnas_filtradas2 = ['Proyecto, Partida Presupuestaria, Estructura','Movimiento']
        nuevo_df = nuevo_df[columnas_filtradas2]
        sum_row = nuevo_df[['Movimiento']].sum()
        # Agregar la fila 'Total' al DataFrame mostrado en AGGrid
        total_row = pd.DataFrame([['Total', sum_row['Movimiento']]], 
                            columns=['Proyecto, Partida Presupuestaria, Estructura','Movimiento'], index=['Total'])
        nuevo_df = pd.concat([nuevo_df, total_row])
        #Creamos una nueva tabla para las metas
        meta_filtro = ['Proyecto','Metas','Nueva Meta']
        ed_df = edit_rows[meta_filtro]

        if total_cod > total_tot:
            if export_as_pdf:
                now = datetime.now()
                fecha_hora = now.strftime("%Y%m%d%H%M")
                excel_file = f'presupuesto{fecha_hora}.xlsx'
                excel_file2 = f'metas{fecha_hora}.xlsx'        
                st.write('Descargando... ¡Espere un momento!')
                edited_rows.to_excel(excel_file, index= False)
                #edited_rows.to_csv('data.csv', index=False)
                edit_rows.to_excel(excel_file2, index= False)
                st.success(f'¡Archivo Excel "{excel_file}" descargado correctamente!')
                st.success(f'¡Archivo Excel "{excel_file2}" descargado correctamente!')

                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()

                # Obtener fecha y hora actual para el título
                pdf.set_title(f"Reforma Presupuesto - {fecha_hora}")
                # Escribir el título en el PDF
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=f"Reforma Presupuesto - {fecha_hora}", ln=True, align="C")
                pdf.cell(200, 15, txt=f"{direc}", ln=True, align="C")
                pdf.ln(10)
                # Anchos de columna para el DataFrame en el PDF
                col_widths = [150, 20]  # Anchos de columna fijos
                # Obtener anchos de columna dinámicos basados en el contenido
                pdf.set_font("Arial", size=7)
                for i, col in enumerate(nuevo_df.columns):
                    pdf.cell(col_widths[i], 10, str(col), border=1, align='C')
                pdf.ln()

                for _, row in nuevo_df.iterrows():
                    a=0
                    b=0
                    for i, value in enumerate(row):
                        # Convertir el valor a string antes de la verificación
                        value = str(value)
                        if len(value) > 25:
                            x = pdf.get_x()  # Guardar la posición X actual
                            y = pdf.get_y()  # Guardar la posición Y actual
                            pdf.multi_cell(col_widths[i], 5, txt=value, border=1)
                            pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        #   b = len(value)/4
                        #elif a == 1:
                        #    x = pdf.get_x()  # Guardar la posición X actual
                        #    y = pdf.get_y()  # Guardar la posición Y actual
                        #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                        #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        else:
                            pdf.cell(col_widths[i], 10, txt=value, border=1, align='C')
                    pdf.ln()
                
                pdf.add_page()
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=f"Reforma metas - {fecha_hora}", ln=True, align="C")
                pdf.cell(200, 15, txt=f"{direc}", ln=True, align="C")
                pdf.ln(10)
                # Anchos de columna para el DataFrame en el PDF
                col_widths2 = [60, 60,60]  # Anchos de columna fijos
                # Obtener anchos de columna dinámicos basados en el contenido
                pdf.set_font("Arial", size=7)
                for i, col in enumerate(ed_df.columns):
                    pdf.cell(col_widths2[i], 14, str(col), border=1, align='C')
                pdf.ln()

                for _, row in ed_df.iterrows():
                    a=0
                    b=0
                    for i, value in enumerate(row):
                        # Convertir el valor a string antes de la verificación
                        value = str(value)
                        if len(value) > 25:
                            x = pdf.get_x()  # Guardar la posición X actual
                            y = pdf.get_y()  # Guardar la posición Y actual
                            pdf.multi_cell(col_widths2[i], 4, txt=value, border=1)
                            pdf.set_xy(x + col_widths2[i], y)  # Restablecer la posición XY
                        #    a = 1
                        #   b = len(value)/4
                        #elif a == 1:
                        #    x = pdf.get_x()  # Guardar la posición X actual
                        #    y = pdf.get_y()  # Guardar la posición Y actual
                        #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                        #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        else:
                            pdf.cell(col_widths2[i], 10, txt=value, border=1, align='C')
                    pdf.ln()
                


                # Guardar el PDF
                pdf_output = f"Reforma Presupuesto_{fecha_hora}.pdf"
                pdf.output(pdf_output)
                st.success(f"Se ha generado el PDF: {pdf_output}")
        else:
            st.warning('No se puede descargar, porque no se esta liberando el presupuesto')
                
def Solicitud ():
        #para conectar a google drive
        #gc = gspread.service_account(filename= 'reformas-402915-4d33dedfb202.json')
        #abrir el archivo de drive
        #sh = gc.open('reforma')
        #mt = gc.open('REFORMA_POA')
        #la hoja de sheets
        #wks = sh.get_worksheet(0)
        #mts = mt.get_worksheet(0)
        #para editar
        odoo = pd.read_excel("tabla_presupuesto.xlsx")
        #odoo = wks.get_all_records()
        metas = pd.read_excel("tabla_metas.xlsx")
        #metas = mts.get_all_records()
        #odoo = [{k: v.encode('utf-8') if isinstance(v, str) else v for k, v in row.items()} for row in odoo]

        df_odoo = pd.DataFrame(odoo)
        df_mt = pd.DataFrame(metas)
        #df_odoo['Codificado'] = df_odoo['Codificado'].replace(",", ".").astype(float)
        #df_odoo['Saldo_Disponible'] = df_odoo['Saldo_Disponible'].replace(",", ".").astype(float)
        st.markdown("<h1 style='text-align:center;background-color: #4e0061; color: #ffffff'>➕ REFORMA SOLICITUD </h1>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center; background-color: #f6f1f7; color: #080200'>(Se solicita presupuesto a la Institución)</h4>", unsafe_allow_html=True)
        st.markdown("---")
        
        #reload---
        reload_data = False
        #filtramos solo PAI
        df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
        #agrupamos las Unidades
        direc = st.selectbox('Escoje la Unidad', options=df_odoo['Unidad'].unique())
        #filtrar columnas 
        df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()##.agg({'Codificado':'sum'},{'Saldo_Disponible':'sum'}) #
        df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
        df_mt= df_mt.loc[df_mt.Unidad == direc]
        #Creamos los Datos ha editar
        data=df_od
        df = pd.DataFrame(data)
        df = agregar_columnas(df)
        st.markdown(f"<h2 style='text-align:center;'> {direc} </h2>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Presupuesto</h3>", unsafe_allow_html=True, help="Se realiza cambios en la columna movimientos, tomar encuenta el saldo disponible para restar un valor caso contrario no se tomara encuenta la reforma.")
        gb = GridOptionsBuilder.from_dataframe(df) 
        gb.configure_column('Unidad', hide=True)#, rowGroup=True, cellRenderer= "agGroupCellRenderer", )
        gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
        gb.configure_column('Estructura')
        gb.configure_column(field ='Codificado', column_name = "Codificado", maxWidth=150, aggFunc="sum")
        gb.configure_column('Saldo_Disponible', aggFunc='sum',maxWidth=150 )
        cellsytle_jscode = JsCode("""
            function(params) {
                if (params.value > '0') {
                    return {
                        'color': 'white',
                        'backgroundColor': 'green'
                    }
                } 
                else if (params.value < '0'){
                    return {
                        'color': 'white',
                        'backgroundColor': 'darkred'
                    }
                }                 
                else {
                    return {
                        'color': 'black',
                        'backgroundColor': 'white'
                    }
                }
            };
        """)
       
        change = JsCode("""
            function isCellEditable(params){
                if(params.data.Saldo_Disponible >= 0 && params.data.Movimiento >= 0 ){
                    return true
                }
                else{
                    return false

                }
                
            }
        """)    
        gb.configure_column('Movimiento', editable= change ,type=['numericColumn'], cellStyle=cellsytle_jscode, aggFunc='sum',maxWidth=150)
        
        gb.configure_column('Nuevo Codificado', valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                            editable=True, type=['numericColumn'], aggFunc='sum',maxWidth=150)
        gb.configure_column('TOTAL', hide=True)

        go = gb.build()
        go['alwaysShowHorizontalScroll'] = True
        go['scrollbarWidth'] = 1
        reload_data = False
        #return_mode_value = DataReturnMode.FILTERED_AND_SORTED
        #update_mode_value = GridUpdateMode.GRID_CHANGED

        edited_df = AgGrid(
            df,
            editable= True,
            gridOptions=go,
            width=1000, 
            height=400, 
            fit_columns_on_grid_load=True,
            theme="streamlit",
            #data_return_mode=return_mode_value, 
            #update_mode=update_mode_value,
            allow_unsafe_jscode=True, 
            #key='an_unique_key_xZs151',
            reload_data=reload_data,
            #no agregar cambia la columna de float a str
            #try_to_convert_back_to_original_types=False
        )
        st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Metas </h3>", unsafe_allow_html=True, help="En la columna Nueva Meta se puede agregar la modificación a la meta actual")
        # Mostrar la tabla con la extensión st_aggrid
        edit_df = AgGrid(df_mt, editable=True,
                        reload_data=reload_data,)
        edit_df = pd.DataFrame(edit_df['data'])

        # Si se detectan cambios, actualiza el DataFrame
        if edited_df is not None:
            # Convierte el objeto AgGridReturn a DataFrame
            edited_df = pd.DataFrame(edited_df['data'])
            edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
            #st.write('Tabla editada:', edited_df)
            #st.write(f'<div style="width: 100%; margin: auto;">{df.to_html(index=False)}</div>',unsafe_allow_html=True)
        

        #CERRAMOS LA SECCIÓN
        st.markdown("---")
        st.subheader("Totales ") 
        total_cod = int(edited_df['Codificado'].sum())
        total_mov = int(edited_df['Movimiento'].sum())
        total_tot = int(edited_df['TOTAL'].sum())

        #DIVIDIMOS EL TABLERO EN 3 SECCIONES
        left_column, center_column, right_column = st.columns(3)
        with left_column:
            st.subheader("Total Codificado: ")
            st.subheader(f"US $ {total_cod:,}")
            #st.dataframe(df_od.stack())
            #st.write(df_od)
        
        with center_column:
            st.subheader("Nuevo Codificado: ")
            st.subheader(f"US $ {total_tot:,}")

        with right_column:
            st.subheader("Total movimiento: ")
            st.subheader(f"US $ {total_mov:,}")


        if total_cod < total_tot:
            st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ccffcc; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor total del codificado es menor al nuevo codificado</h4></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor total del codificado es mayor o igual al nuevo codificado</h4></div>', unsafe_allow_html=True)

        try:
            edited_rows = edited_df[edited_df['Movimiento'] != 0]
            edit_rows = edit_df[edit_df['Nueva Meta'] != '']
        except:
            st.write('No se realizaron cambios en de la información')

        st.markdown("---")
        export_as_pdf = st.button("Guardar y Descargar")
        #Creamos una nueva tabla para el presupuesto
        columnas_filtradas = ['PROYECTO','Código','Estructura','Movimiento']
        nuevo_df = edited_rows[columnas_filtradas]
        nuevo_df['Proyecto, Partida Presupuestaria, Estructura'] = nuevo_df['PROYECTO'] + ' | ' + nuevo_df['Código'] + ' | ' + nuevo_df['Estructura']
        columnas_filtradas2 = ['Proyecto, Partida Presupuestaria, Estructura','Movimiento']
        nuevo_df = nuevo_df[columnas_filtradas2]
        sum_row = nuevo_df[['Movimiento']].sum()
        # Agregar la fila 'Total' al DataFrame mostrado en AGGrid
        total_row = pd.DataFrame([['Total', sum_row['Movimiento']]], 
                            columns=['Proyecto, Partida Presupuestaria, Estructura','Movimiento'], index=['Total'])
        nuevo_df = pd.concat([nuevo_df, total_row])
        #Creamos una nueva tabla para las metas
        meta_filtro = ['Proyecto','Metas','Nueva Meta']
        ed_df = edit_rows[meta_filtro]

        if total_cod < total_tot:
            if export_as_pdf:
                now = datetime.now()
                fecha_hora = now.strftime("%Y%m%d%H%M")
                excel_file = f'presupuesto{fecha_hora}.xlsx'
                excel_file2 = f'metas{fecha_hora}.xlsx'        
                st.write('Descargando... ¡Espere un momento!')
                edited_rows.to_excel(excel_file, index= False)
                #edited_rows.to_csv('data.csv', index=False)
                edit_rows.to_excel(excel_file2, index= False)
                st.success(f'¡Archivo Excel "{excel_file}" descargado correctamente!')
                st.success(f'¡Archivo Excel "{excel_file2}" descargado correctamente!')

                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()

                # Obtener fecha y hora actual para el título
                pdf.set_title(f"Reforma Presupuesto - {fecha_hora}")
                # Escribir el título en el PDF
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=f"Reforma Presupuesto - {fecha_hora}", ln=True, align="C")
                pdf.cell(200, 15, txt=f"{direc}", ln=True, align="C")
                pdf.ln(10)
                # Anchos de columna para el DataFrame en el PDF
                col_widths = [150, 20]  # Anchos de columna fijos
                # Obtener anchos de columna dinámicos basados en el contenido
                pdf.set_font("Arial", size=7)
                for i, col in enumerate(nuevo_df.columns):
                    pdf.cell(col_widths[i], 10, str(col), border=1, align='C')
                pdf.ln()

                for _, row in nuevo_df.iterrows():
                    a=0
                    b=0
                    for i, value in enumerate(row):
                        # Convertir el valor a string antes de la verificación
                        value = str(value)
                        if len(value) > 25:
                            x = pdf.get_x()  # Guardar la posición X actual
                            y = pdf.get_y()  # Guardar la posición Y actual
                            pdf.multi_cell(col_widths[i], 5, txt=value, border=1)
                            pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        #   b = len(value)/4
                        #elif a == 1:
                        #    x = pdf.get_x()  # Guardar la posición X actual
                        #    y = pdf.get_y()  # Guardar la posición Y actual
                        #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                        #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        else:
                            pdf.cell(col_widths[i], 10, txt=value, border=1, align='C')
                    pdf.ln()
                
                pdf.add_page()
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=f"Reforma metas - {fecha_hora}", ln=True, align="C")
                pdf.cell(200, 15, txt=f"{direc}", ln=True, align="C")
                pdf.ln(10)
                # Anchos de columna para el DataFrame en el PDF
                col_widths2 = [60, 60,60]  # Anchos de columna fijos
                # Obtener anchos de columna dinámicos basados en el contenido
                pdf.set_font("Arial", size=7)
                for i, col in enumerate(ed_df.columns):
                    pdf.cell(col_widths2[i], 14, str(col), border=1, align='C')
                pdf.ln()

                for _, row in ed_df.iterrows():
                    a=0
                    b=0
                    for i, value in enumerate(row):
                        # Convertir el valor a string antes de la verificación
                        value = str(value)
                        if len(value) > 25:
                            x = pdf.get_x()  # Guardar la posición X actual
                            y = pdf.get_y()  # Guardar la posición Y actual
                            pdf.multi_cell(col_widths2[i], 4, txt=value, border=1)
                            pdf.set_xy(x + col_widths2[i], y)  # Restablecer la posición XY
                        #    a = 1
                        #   b = len(value)/4
                        #elif a == 1:
                        #    x = pdf.get_x()  # Guardar la posición X actual
                        #    y = pdf.get_y()  # Guardar la posición Y actual
                        #    pdf.multi_cell(col_widths[i], 5+b, txt=value, border=1,  align='C')
                        #    pdf.set_xy(x + col_widths[i], y)  # Restablecer la posición XY
                        #    a = 1
                        else:
                            pdf.cell(col_widths2[i], 10, txt=value, border=1, align='C')
                    pdf.ln()
                


                # Guardar el PDF
                pdf_output = f"Reforma Presupuesto_{fecha_hora}.pdf"
                pdf.output(pdf_output)
                st.success(f"Se ha generado el PDF: {pdf_output}")
        else:
             st.warning('No se puede descargar, porque no se esta solicitando presupuesto')

page_names_to_funcs = {

    "Inicio": Inicio,    
    "Interna": Interna,
    "Externa": Externa,
    "Liberación": Liberación,
    "Solicitud": Solicitud
}
#st.markdown("""
#<style>
#    [data-testid=stSidebar] {
#        background-color: #020f69;
#        text-color: #ffffff;
#    }
#</style>
#""", unsafe_allow_html=True)


odoo = pd.read_excel("tabla_presupuesto.xlsx")        
df_odoo = pd.DataFrame(odoo)
FECHA = df_odoo.iloc[2]["Fecha"]
st.sidebar.image('logo GadPP.png', caption='Unidad de Planificación')
st.sidebar.title("Reformas:")
demo_name = st.sidebar.selectbox('Escoja el tipo de Reforma', page_names_to_funcs.keys())
page_names_to_funcs[demo_name]()

with st.sidebar.expander("🗺 Datos", expanded=True):
    st.markdown(f"""
    - La información del `presupuesto` se actualiza cada día a las 10 de la mañana.
    - Fecha actual de la información presupuestaria: (dd/mm/aa) `{FECHA}`  
        """)

contrasena_correcta = "CPI"

# Configuración de la aplicación
st.sidebar.title("ADMIN")

# Entrada de contraseña
contrasena = st.sidebar.text_input("contraseña:", type="password")

# Verificar si la contraseña es correcta
if contrasena == contrasena_correcta:
    gc = gspread.service_account(filename= 'reformas-402915-4d33dedfb202.json')
        #abrir el archivo de drive
    sh = gc.open('reforma')
    wks = sh.get_worksheet(0)
    odoo = wks.get_all_records()
    df_odoo = pd.DataFrame(odoo)

    mt = gc.open('REFORMA_POA')
    mts = mt.get_worksheet(0)
    metas = mts.get_all_records()
    df_metas = pd.DataFrame(metas)

    st.sidebar.success("Contraseña correcta. ¡Bienvenido!")
    df_odoo.to_excel('tabla_presupuesto.xlsx', index= False)
    df_metas.to_excel('tabla_metas.xlsx', index= False)
else:
    st.sidebar.error("Contraseña incorrecta. Acceso denegado.")

