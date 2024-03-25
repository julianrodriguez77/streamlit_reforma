import streamlit as st
import pandas as pd
from st_aggrid import JsCode, AgGrid, GridOptionsBuilder
from datetime import datetime
import hydralit_components as hc
#from PIL import Image
from io import BytesIO
from reportlab.lib.pagesizes import letter,A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph,Image,  PageTemplate, Frame


#para que reconosca la tabla xlsx
import pip
pip.main(["install", "openpyxl"])
#nombre de pagina
st.set_page_config(page_title= 'Reformas CPI',
                    page_icon='moneybag:',
                    layout='wide' )
#ocultar menu streamlit
hide_st_style = """
<style>
MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
# DEFINIMOS LA VARIABLE MOVIMIENTO
# Para la seccion 1,2,3 y 4
def agregar_columnas(df):
    df['Movimiento'] = 0
    df['TOTAL'] = df['Codificado'] + df['Movimiento']
    return df
# Para la seccion 2 segunda tabla
def agregar_column(dfd):
    dfd['Movimiento'] = 0
    dfd['TOTAL'] = dfd['Codificado'] + dfd['Movimiento']
    return dfd
#Cargamos la Data
odoo = pd.read_excel("tabla_presupuesto.xlsx")
metas = pd.read_excel("tabla_metas.xlsx")
#Items de inicio 

#Titulo 
st.markdown("<h1 style='text-align: center; background-color: #000045; color: #ece5f6'>Sistema de Reformas al POA</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; background-color: #000045; color: #ece5f6'>Dirección de Planificación</h4>", unsafe_allow_html=True)
#Pestañas
menu_data = [
{'id': 1, 'label': "Información", 'icon': "fa fa-home"},
{'id': 2, 'label': "Documentación"},
{'id': 3, 'label': "Interna"},
{'id': 4, 'label': "Externa"},
{'id': 5, 'label': "Liberación"},
{'id': 6, 'label': "Solicitud"}
]
#Diseño de pestañas
paginas=int(hc.nav_bar(
    menu_definition=menu_data,
    hide_streamlit_markers=False,
    sticky_nav=True,
    sticky_mode='pinned',
    override_theme={'menu_background': '#4c00a5'}
))

if paginas ==1:           
    st.header("**📖 REFORMAS AL PLAN OPERATIVO ANUAL**")
    st.markdown("""     
                El Plan Operativo Anual (POA) es un instrumento estratégico que enlaza la planificación con el presupuesto. 
                Su propósito es guiar la ejecución de objetivos institucionales, estableciendo pasos claros hacia metas 
                evaluables a corto plazo. Se caracteriza por concretar lo planificado de acuerdo con las capacidades y recursos 
                reales disponibles, asegurando una gestión eficiente y alineada con los objetivos estratégicos. 
                     
                **Esta plataforma permite reformar esta planificación siempre y cuando haya las debidas justificaciones técnicas
                 y la documentación de respaldo o sustento.**
                """)
        
    st.info("""
                Para cualquier inconveniente o duda adicional en la utilización de la plataforma, se invita a ponerse en contacto con Cecilia Sosa,
                a la extensión institucional 12022.
            """)
if paginas == 2:
        st.markdown('A continuación, puede encontrar la normativa legal vigente y demás documentos de interés para el tema de reformas al presupuesto.')
        st.markdown("""
                💡 **Enlaces**
                    
                - [**RESOLUCIÓN ORDENANZA PROVINCIAL No. 06-CPP-2023-2027**](https://docs.snowflake.com/) 
                - [**ADMINISTRATIVA No. 03-SG-2022**](https://docs.snowflake.com/) 
                - [**RESOLUCIÓN ADMINISTRATIVA No. 08-DGSG-2022**](https://docs.snowflake.com/) 
                - [**MEMORANDO 55-DP-23**](https://docs.snowflake.com/) 
                """)

if paginas == 3:
    #Llamamos a la base
    df_odoo = pd.DataFrame(odoo)
    df_mt = pd.DataFrame(metas)
    #ENCABEZADO
    st.markdown("<h1 style='text-align:center;background-color: #000045; color: #ffffff'>🔄 REFORMA INTERNA </h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; background-color: #f0efeb; color: #080200'>Dentro de la misma Dirección/Unidad</h4>", unsafe_allow_html=True)
    st.markdown("")
    with st.expander(f"Resumen", expanded=False):
         st.markdown("""
                    
                    Corresponde a la Reforma al POA donde una Unidad/Dirección ejecutora realiza cambios a su propia planificación presupuestaria o física.
                    """
                    )
         st.info(""" **NOTA**: Las modificaciones se hacen sobre saldos disponibles y metas vigentes. """)
     
    #st.markdown("<h2 style='text-align: left;  color: #ccccc'>Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; background-color: #26469C; color: #ffffff'> Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E'>1.- Escoga la unidad en la cual desea realizar la reforma </h5>", unsafe_allow_html=True)
    #FILTRAMOS SOLO PAI
    df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
    #AGRUPAMOS LAS UNIDADES
    direc = st.selectbox('',    options=df_odoo['Unidad'].unique())
    #FILTRAMOS COLUMNAS 
    df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()
    df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mt= df_mt.loc[df_mt.Unidad == direc]
    df_mtfil = ['Proyecto','Metas','Nueva Meta','Observación']
    df_mt = df_mt[df_mtfil]        
    df = pd.DataFrame(df_od)
    df = agregar_columnas(df)
    #SUBTITULOS
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>2.- Escoja el Proyecto y ajuste los valores en la casilla de “Incremento / Disminución”. </h5>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Presupuesto de {direc} </h3>", unsafe_allow_html=True, help="Se realiza cambios en la columna Movimientos, tomar encuenta el saldo disponible para restar un valor caso contrario no se tomara encuenta la reforma.")
    #FORMATO DE COLUMNAS
    gb = GridOptionsBuilder.from_dataframe(df) 
    gb.configure_column('Unidad', hide=True)
    gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True, cellRenderer= "agGroupCellRenderer")
    gb.configure_column('Estructura', header_name="Actividad")
    gb.configure_column(field ='Codificado', maxWidth=150, aggFunc='sum', valueFormatter="data.Codificado.toLocaleString('en-US');")
    gb.configure_column('Saldo_Disponible', header_name="Saldo", maxWidth=120, valueFormatter="data.Saldo_Disponible.toLocaleString('en-US');", aggFunc='sum' )
    cellsytle_jscode = JsCode("""
        function(params) {
            if (params.value > '0') {
                return {
                    'color': 'black',
                    'backgroundColor': 'green'
                }
            } 
            else if (params.value < '0'){
                return {
                    'color': 'black',
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

    gb.configure_column('Movimiento', header_name='Increm/Dismi' , editable= True ,type=['numericColumn'], cellStyle=cellsytle_jscode, maxWidth=120, valueFormatter="data.Movimiento.toLocaleString('en-US');")
    gb.configure_column('Nuevo Codificado',header_name='Nuevo Cod' , valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                        type=['numericColumn'],maxWidth=150, valueFormatter="data.Nuevo Codificado.toLocaleString('en-US');", aggFunc='sum', enableValue=True)
    gb.configure_column('TOTAL', hide=True)
    go = gb.build()

    edited_df = AgGrid(
        df,
        gridOptions=go,
        width=1000, 
        height=350, 
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=True
    )
    
    # Si se detectan cambios, actualiza el DataFrame
    if edited_df is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edited_df = pd.DataFrame(edited_df['data'])
        edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
    #Barra despegable para crear una nueva partida
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>3.- Crear nuevas actividades presupuestarias: Seleccione el proyecto, coloque el nombre de la partida y el valor a incrementar.</h5>", unsafe_allow_html=True)
    with st.expander(f"🆕  Crear una partida nueva para {direc} ", expanded=False): 
        st.markdown("<p style='text-align: center; background-color: #B5E6FC;'> Agregar nueva partida </p>", unsafe_allow_html=True)
        dfnuevop = pd.DataFrame(columns=['Proyecto','Estructura','Incremento','Parroquia'])
        config = {
            'Proyecto' : st.column_config.SelectboxColumn('Proyecto',width='large', options=df_od['PROYECTO'].unique()),
            'Estructura' : st.column_config.TextColumn('Estructura', width='large', required=True),
            'Incremento' : st.column_config.NumberColumn('Incremento', min_value=0,width='medium', required=True),
            'Parroquia' : st.column_config.TextColumn('Parroquia', width='medium', required=True)
        }
        result = st.data_editor(dfnuevop, column_config = config, num_rows='dynamic')
        
        if st.button('Crear partida:'):
            st.write(result)
    #TOTALES
    total_cod = int(edited_df['Codificado'].sum())
    total_tot = int(edited_df['TOTAL'].sum())
    nuevo_p = int(result['Incremento'].sum())
    #Creamos una tabla de los totales
    total_row = {
        'PROYECTO': 'Total',  # No se calcula el total para la columna de texto
        'Total_Codificado': df['Codificado'].sum(),
        'Total_Saldo': df['Saldo_Disponible'].sum(),
        'Tot_Increm/Dismi': edited_df['Movimiento'].sum() + nuevo_p,
        'Total_Nuev_Codif':  nuevo_p + edited_df['TOTAL'].sum() 
    }
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>4.- Verifica que los valores totales estén correctos. </h5>", unsafe_allow_html=True)
    #Diseño de las tabla totales
    total_df = pd.DataFrame([total_row])
    gbt = GridOptionsBuilder.from_dataframe(total_df)
    gbt.configure_column('PROYECTO', minWidth =810 )
    gbt.configure_column('Total_Saldo', header_name='Total Saldo', maxWidth =120, valueFormatter="data.Total_Saldo.toLocaleString('en-US');" )
    gbt.configure_column('Tot_Increm/Dismi', header_name='Tot Incr/Dismi', maxWidth =140, valueFormatter="data.Tot_Increm/Dismi.toLocaleString('en-US');" )
    gbt.configure_column('Total_Nuev_Codif', header_name='Tot Nuev Cod', maxWidth =135, valueFormatter="data.Total_Nuev_Codif.toLocaleString('en-US');" )
    gbt.configure_column('Total_Codificado', header_name='Tot Codificado',  maxWidth =145, valueFormatter="data.Total_Codificado.toLocaleString('en-US');" )
    #Tabla totales
    AgGrid(total_df,
            gridOptions=gbt.build(),
            theme='alpine',
            height=120)
    #st.markdown("---")
    
    if total_cod != total_tot+nuevo_p:
        st.markdown('<div style="max-width: 900px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">Revisa. El valor codificado es distinto antes y después de la reforma.</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="max-width: 900px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">Correcto. El valor codificado es el mismo antes y después de la reforma</h3></div>', unsafe_allow_html=True)
    
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>5.- Escoja el proyecto, revise las metas y de ser necesario modifíquelas de ser necesario. </h5>", unsafe_allow_html=True)
    #TABLA DE METAS
    st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Metas de {direc} </h3>", unsafe_allow_html=True, help="En la columna Nueva Meta se puede agregar la modificación a la meta actual")
    #Diseño de tablas
    gbmt = GridOptionsBuilder.from_dataframe(df_mt)
    gbmt.configure_column('Proyecto', minWidth =408, editable=False )
    gbmt.configure_column('Metas', minWidth =350, editable=False )
    gbmt.configure_column('Nueva Meta', minWidth =300, editable=True )
    gbmt.configure_column('Observación', minWidth =300, editable=True )
    #Tabla Metas
    edit_df = AgGrid(df_mt,
                        gridOptions=gbmt.build(),
                        height=350)
    
    edit_df = pd.DataFrame(edit_df['data'])
    
    #Condición para que se agregen unicamente las partidas y metas editadas en el excel y pdf
    try:
        edited_rows = edited_df[edited_df['Movimiento'] != 0]
        edit_rows = edit_df[edit_df['Nueva Meta'] != '-']
    except:
        st.write('No se realizaron cambios en de la información')

    st.markdown('<hr style="border: 1.5px solid red; margin-bottom: 5px;">', unsafe_allow_html=True)

    st.info("""
            Descargue adjunto:
                
            En el caso de que no haya anteriores revisiones, guarde la información y descargue el documento PDF para adjuntarlo al memorando de solicitud de reforma.
            * Recuerde que el memorando de solicitud debe contener las justificaciones técnicas y demás documentación de respaldo.

            """)
    #función para descargar el excel
    def descargar_xlsx(edited_rows, edit_rows, result):
            # Guardar los DataFrames en dos hojas de un archivo XLSX en memoria
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    edited_rows.to_excel(writer, sheet_name='Presupuesto', index=False)
                    edit_rows.to_excel(writer, sheet_name='Metas', index=False)
                    result.to_excel(writer, sheet_name='Nueva_partida', index=False)
                output.seek(0)
                return output
    #Boton para descargar la información
    export_as_pdf = st.columns(3)[1].button("Guardar información")
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
    dfp1 = nuevo_df
    #Creamos una nueva tabla para las metas
    meta_filtro = ['Proyecto','Metas','Nueva Meta','Observación']
    dfp3 = edit_rows[meta_filtro]
    result2 = result
    #nueva partida
    result2['Proyecto, Estructura']=result2['Proyecto']+ ' | ' + result2['Estructura']
    resul_filtro=['Proyecto, Estructura','Incremento']
    result2=result2[resul_filtro]
    sum_res=result2[['Incremento']].sum()
    total_res = pd.DataFrame([['Total', sum_res['Incremento']]], 
                        columns=['Proyecto, Estructura','Incremento'], index=['Total'])
    result2 = pd.concat([result2, total_res])
    dfp2=result2
    # Condición para poder descargar
    if total_cod == total_tot+nuevo_p:
        if export_as_pdf:
            now = datetime.now()
            fecha_hora = now.strftime("%Y%m%d%H%M")    
            st.write('Descargando... ¡Espere un momento!')
            
            
            def export_to_pdf(dfp1, dfp2, dfp3):
                # Crear un objeto BytesIO para almacenar el PDF
                pdf_buffer = BytesIO()
                # Crear un objeto SimpleDocTemplate para el PDF
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
                # Obtener estilos de texto predefinidos
                styles = getSampleStyleSheet()
                # Agregar título a la página
                title = f"Reforma Interna-{fecha_hora}"
                title_style = styles['Title']
                title_style.spaceAfter = 12
                title_style.spaceBefore = 22
                title_paragraph = Paragraph(title, title_style)
                
                subtitle_text = f"{direc} - Presupuesto"
                subtitle_paragraph = Paragraph(subtitle_text, title_style)

                title2 = f"{direc} - Nueva Partida"
                title2_paragraph = Paragraph(title2, title_style)

                title3 = f"{direc} - Metas"
                title3_paragraph = Paragraph(title3, title_style)         
                
                # Agregar imagen a la página
                img_path = "logo GadPP.png"  # Reemplaza con la ruta de tu imagen
                #image = Image(img_path, width=100, height=100) 

                # Convertir DataFrame a lista de listas para la tabla
                data1 = [dfp1.columns.tolist()] + [configure_cell(dfp1, row) for _, row in dfp1.iterrows()]  # Agregar una fila con los nombres de las variables
                # Crear la tabla con los datos del DataFrame
                table1 = Table(data1, repeatRows=1, colWidths=[400] + [None] * (len(dfp1.columns) - 1))

                        # Establecer estilos para la tabla
                table1.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinear el contenido al centro
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  # Fuente en negrita para la primera fila (encabezado)
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Agregar espacio inferior a la primera fila
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  # Agregar bordes a la tabla
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  # Espacio después de cada fila
                ]))

                data2 = [dfp2.columns.tolist()] + [configure_cell(dfp2, row) for _, row in dfp2.iterrows()]
                # Crear la segunda tabla con los datos del DataFrame 2
                table2 = Table(data2, repeatRows=1, colWidths=[400] + [None] * (len(dfp2.columns) - 1))

                # Establecer estilos para la segunda tabla
                table2.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                ]))

                data3 = [dfp3.columns.tolist()] + [configure_cellmetas(dfp3, row) for _, row in dfp3.iterrows()]
                # Crear la segunda tabla con los datos del DataFrame 2
                table3 = Table(data3, repeatRows=1, colWidths=[200] + [None] * (len(dfp3.columns) - 1))

                # Establecer estilos para la segunda tabla
                table3.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                ]))
            # Construir el PDF con la tabla
                frames = [Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')]
                template = PageTemplate(id='encabezado_izquierdo', frames=frames, onPage=lambda canvas, doc, **kwargs: canvas.drawImage(img_path, doc.leftMargin-40, doc.height+55, width=120, height=90, preserveAspectRatio=True))
                doc.addPageTemplates([template])
                doc.build([title_paragraph, subtitle_paragraph, table1,title2_paragraph, table2,title3_paragraph, table3])
                # Obtener el contenido del BytesIO
                pdf_content = pdf_buffer.getvalue()
                # Cerrar el BytesIO
                pdf_buffer.close()

                return pdf_content

            def configure_cell(df, row):
                styles = getSampleStyleSheet()
                styles['Normal'].fontSize = 8
                first_col_text = str(row[df.columns[0]])
                first_col_text = '\n'.join([first_col_text[j:j+20] for j in range(0, len(first_col_text), 20)])
                return [Paragraph(first_col_text, styles['Normal'], encoding='utf-8')] + [str(row[col]) for col in df.columns[1:]]
        
            def configure_cellmetas(df, row):
                styles = getSampleStyleSheet()
                styles['Normal'].fontSize = 8
                # Aplicar el formato con '\n' a todas las columnas
                formatted_columns = [('\n'.join(str(row[col])[j:j+20] for j in range(0, len(str(row[col])), 20)), styles['Normal']) for col in df.columns]
                # Devolver una lista de objetos Paragraph para cada celda
                return [Paragraph(text, style, encoding='utf-8') for text, style in formatted_columns]

            if export_as_pdf:
                # Llamar a la función para exportar DataFrame a PDF
                pdf_content = export_to_pdf(dfp1, dfp2, dfp3)
                # Descargar el PDF
                st.download_button('Descargar PDF', pdf_content, file_name=f'Reforma_{direc}_{fecha_hora}.pdf', key='download_button')
                # Mensaje de éxito
                st.success('Tabla exportada y PDF descargado exitosamente.') 

            #pdf.output(pdf_output)
            archivo_xlsx = descargar_xlsx(edited_rows, edit_rows, result)
            st.download_button(
                label="Haz clic para descargar",
                data=archivo_xlsx.read(),
                key="archivo_xlsx",
                file_name=f"Reforma_{fecha_hora}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )       
    else:
        #st.markdown('<div style=" margin: 0 auto; background-color:#D9FDFF; padding:10px; text-align: center;"><h4 style="color:#01464A;">Los Movimientos tienen incosistencia revisar para descargar.</h3></div>', unsafe_allow_html=True)
        st.warning('Los Movimientos tienen incosistencia revisar para descargar')    

if paginas == 4:

    st.markdown("<h1 style='text-align:center;background-color: #028d96; color: #ffffff'>🌀 REFORMA AL POA EXTERNA </h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; background-color: #f0efeb; color: #080200'>Entre diferentes Dirección/Unidad </h4>", unsafe_allow_html=True)
    st.markdown("")
    with st.expander("💱  Resumen", expanded=False):
        st.markdown("""
                    
                    Corresponde a la Reforma al POA en que se modifican los valores codificados de las Unidades (disminución/ incremento), por transferencia de valores de una unidad a otra unidad del GADPP,  y afecta:

                    - A la **programación presupuestaria**, por incremento o disminución de los valores codificados de las actividades de los proyectos de las unidades involucradas, y/o
                    - A la **programación física**, por modificación o no de las metas de los proyectos de las unidades involucradas.

                    Una vez realizada la reforma al POA, presupuestaria y/o de metas, se procede a guardar la información; automáticamente se generará un archivo pdf codificado, con la información de las modificaciones realizadas ya sea solo presupuestaria y/o de metas.

                    """
                    )
        st.info(""" **NOTA**: Las modificaciones se harán sobre los saldos disponibles no comprometidos de las asignaciones. """)

    #st.markdown("<h2 style='text-align: left;  color: #ccccc'>Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; background-color: #26469C; color: #ffffff'> Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E'>1.- Escoga la unidad donde se realiza la disminución </h5>", unsafe_allow_html=True)

    df_odoo = pd.DataFrame(odoo)
    df_mt = pd.DataFrame(metas)
    df_mt2 = pd.DataFrame(metas)
    #filtramos solo PAI
    df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
    odf = df_odoo
    #agrupamos las Unidades
    direc = st.selectbox('', options=df_odoo['Unidad'].unique())
    #filtrar columnas 
    df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()##.agg({'Codificado':'sum'},{'Saldo_Disponible':'sum'}) #
    df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mt= df_mt.loc[df_mt.Unidad == direc]
    df_mtfil = ['Proyecto','Metas','Nueva Meta','Observación']
    df_mt = df_mt[df_mtfil]  
    #Creamos los Datos ha editar
    data1=df_od
    df = pd.DataFrame(data1)
    df = agregar_columnas(df)
    #Espacio
    st.title('')
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>2.- Escoja el Proyecto y ajuste los valores en la casilla de “Incremento / Disminución”. </h5>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; background-color: #f1f6f7; color: #080200'> Tabla de Presupuestos de {direc} </h3>", unsafe_allow_html=True)
    
        
    gb = GridOptionsBuilder.from_dataframe(df) 
    gb.configure_column('Unidad', hide=True)
    gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
    gb.configure_column('Estructura')
    gb.configure_column(field ='Codificado', maxWidth=150, aggFunc="sum", valueFormatter="data.Codificado.toLocaleString('en-US');")
    gb.configure_column('Saldo_Disponible',header_name='Saldo', maxWidth=130, valueFormatter="data.Saldo_Disponible.toLocaleString('en-US');", aggFunc='sum' )
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
        
    gb.configure_column('Movimiento', header_name='Increm/Dismi' , editable= True ,type=['numericColumn'], cellStyle=cellsytle_jscode, maxWidth=150, valueFormatter="data.Movimiento.toLocaleString('en-US');")
    gb.configure_column('TOTAL2',header_name='Nuev cod', valueGetter='Number(data.Codificado) + Number(data.Movimiento)',maxWidth=150, cellRenderer='agAnimateShowChangeCellRenderer',
                         editable=True, type=['numericColumn'], aggFunc='sum')
    gb.configure_column('TOTAL', hide=True)
    go = gb.build()
        
    edited_df = AgGrid(
        df,
        editable= True,
        gridOptions=go,
        width=1000, 
        height=350, 
        fit_columns_on_grid_load=True,
        theme="streamlit",
        allow_unsafe_jscode=True, 
        #reload_data=reload_data,
        
    )
    
        # Si se detectan cambios, actualiza el DataFrame
    if edited_df is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edited_df = pd.DataFrame(edited_df['data'])
        edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
        

    total_cod = int(edited_df['Codificado'].sum())
    total_mov = int(edited_df['Movimiento'].sum())
    total_tot = int(edited_df['TOTAL'].sum())

    total_row = {
            'PROYECTO': 'Total', 
            'Total_Codificado': df['Codificado'].sum(),
            'Total_Saldo': df['Saldo_Disponible'].sum(),
            'Tot_Increm/Dismi': edited_df['Movimiento'].sum(),
            'Total_Nuev_Codif':  edited_df['TOTAL'].sum() 
        }
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>3.- Verifica que los valores totales estén correctos. </h5>", unsafe_allow_html=True)
    
    total_df = pd.DataFrame([total_row])
    gbt = GridOptionsBuilder.from_dataframe(total_df) 
    gbt.configure_column('PROYECTO', minWidth =810 )
    gbt.configure_column('Total_Saldo', header_name='Total Saldo', maxWidth =120, valueFormatter="data.Total_Saldo.toLocaleString('en-US');" )
    gbt.configure_column('Tot_Increm/Dismi', header_name='Tot Incr/Dismi', maxWidth =140, valueFormatter="data.Tot_Increm/Dismi.toLocaleString('en-US');" )
    gbt.configure_column('Total_Nuev_Codif', header_name='Tot Nuev Cod', maxWidth =135, valueFormatter="data.Total_Nuev_Codif.toLocaleString('en-US');" )
    gbt.configure_column('Total_Codificado', header_name='Tot Codificado',  maxWidth =145, valueFormatter="data.Total_Codificado.toLocaleString('en-US');" )
    AgGrid(total_df,
           gridOptions=gbt.build(),
           theme='alpine',
           height=120)

    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>4.- Escoja el proyecto, revise las metas y de ser necesario modifíquelas de ser necesario. </h5>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; background-color: #f1f6f7; color: #080200'> Tabla de Metas de {direc}</h3>", unsafe_allow_html=True)
    # Mostrar la tabla con la extensión st_aggrid
    with st.expander("💱  Realizar modificaciones a las metas", expanded=False):
        gbmt = GridOptionsBuilder.from_dataframe(df_mt)
        gbmt.configure_column('Proyecto', minWidth =408, editable=False )
        gbmt.configure_column('Metas', minWidth =350, editable=False )
        gbmt.configure_column('Nueva Meta', minWidth =300, editable=True )
        gbmt.configure_column('Observación', minWidth =300, editable=True )

        edit_df = AgGrid(df_mt,
                         gridOptions=gbmt.build(),
                         height=350)
                            #reload_data=reload_data,)
        edit_df = pd.DataFrame(edit_df['data'])

    #CERRAMOS LA SECCIÓN
    st.markdown("---")


    if total_mov < 0:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">Se a disminuido el valor de:  ${total_mov:} en {direc}</h3></div>', unsafe_allow_html=True)
    elif total_mov == 0:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">No se a realizado ninguna disminución del codificado en {direc}</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">No se puede incrementar un valor solo disminuir  </h3></div>', unsafe_allow_html=True)
        
    ##################################################
    #################     TABLA 2  ###################
    ################################################## 
    #st.markdown("---")
    #st.markdown("<h2 style='text-align: center; background-color: #26469C; color: #ffffff'> Paso 2: Unidad a que se le asigna el incremento  </h2>", unsafe_allow_html=True)
    #st.markdown("<h2 style='text-align: left;  color: #ccccc'>Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E'>5.- Escoja la Unidad donde se agregara el incremento </h5>", unsafe_allow_html=True)
    
    # Obtener las opciones para el segundo selectbox excluyendo la opción seleccionada en el primero    
    opci = odf['Unidad'][odf['Unidad'] != direc]
    selec= st.selectbox('', options= opci.unique())
    #filtrar columnas 
    dfff= odf.loc[odf.Unidad == selec].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()
    dfff.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mtt= df_mt2.loc[df_mt2.Unidad == selec]
    df_mtfil = ['Proyecto','Metas','Nueva Meta','Observación']
    df_mtt = df_mtt[df_mtfil]
    #Creamos los Datos ha editar
    data2=dfff
    dfd2 = pd.DataFrame(data2)
    dfd2 = agregar_column(dfd2)
    
    st.title('')
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>6.- Escoja el Proyecto y ajuste los valores en la casilla de “Incremento / Disminución”. </h5>", unsafe_allow_html=True)

    st.markdown(f"<h3 style='text-align: center; background-color: #f1f6f7; color: #080200'> Tabla de Presupuesto de {selec} </h3>", unsafe_allow_html=True)
   
    edi = AgGrid(
        dfd2,
        editable= True,
        gridOptions=go,
        width=1000, 
        height=350, 
        fit_columns_on_grid_load=True,
        theme="streamlit",
        allow_unsafe_jscode=True, 
        
    )

    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>7.- Crear nuevas actividades presupuestarias: Seleccione el proyecto, coloque el nombre de la partida y el valor a incrementar.</h5>", unsafe_allow_html=True)
    
    with st.expander("🆕  Crear una partida nueva", expanded=False): 
            st.markdown("<p style='text-align: center; background-color: #B5E6FC;'> Agregar nueva partida </p>", unsafe_allow_html=True)
            dfnuevop = pd.DataFrame(columns=['Proyecto','Estructura','Incremento','Parroquia'])
            config = {
                'Proyecto' : st.column_config.SelectboxColumn('Proyecto',width='large', options=df_od['PROYECTO'].unique()),
                'Estructura' : st.column_config.TextColumn('Estructura', width='large', required=True),
                'Incremento' : st.column_config.NumberColumn('Incremento', min_value=0,width='medium', required=True),
                'Parroquia' : st.column_config.TextColumn('Parroquia', width='medium', required=True)
            }

            result = st.data_editor(dfnuevop, column_config = config, num_rows='dynamic')

            if st.button('Crear partida:'):
                st.write(result)
        
    # Si se detectan cambios, actualiza el DataFrame
    if edi is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edi = pd.DataFrame(edi['data'])
        edi['TOTAL'] = edi['Codificado'] + edi['Movimiento']
       
    total_cod2 = int(edi['Codificado'].sum())
    total_mov2 = int(edi['Movimiento'].sum())
    total_tot2 = int(edi['TOTAL'].sum())
    nuevo_p = int(result['Incremento'].sum())

    total_row = {
            'PROYECTO': 'Total',  # No se calcula el total para la columna de texto
            'Total_Codificado': edi['Codificado'].sum(),
            'Total_Saldo': edi['Saldo_Disponible'].sum(),
            'Tot_Increm/Dismi': nuevo_p + edi['Movimiento'].sum(),
            'Total_Nuev_Codif':  edi['TOTAL'].sum() + nuevo_p  
        }
    
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>8.- Verifica que los valores totales estén correctos. </h5>", unsafe_allow_html=True)
    
    total_df = pd.DataFrame([total_row])
    gbt = GridOptionsBuilder.from_dataframe(total_df) 
    gbt.configure_column('PROYECTO', minWidth =810 )
    gbt.configure_column('Total_Saldo', header_name='Total Saldo', maxWidth =120, valueFormatter="data.Total_Saldo.toLocaleString('en-US');" )
    gbt.configure_column('Tot_Increm/Dismi', header_name='Tot Incr/Dismi', maxWidth =140, valueFormatter="data.Tot_Increm/Dismi.toLocaleString('en-US');" )
    gbt.configure_column('Total_Nuev_Codif', header_name='Tot Nuev Cod', maxWidth =135, valueFormatter="data.Total_Nuev_Codif.toLocaleString('en-US');" )
    gbt.configure_column('Total_Codificado', header_name='Tot Codificado',  maxWidth =145, valueFormatter="data.Total_Codificado.toLocaleString('en-US');" )
   
    AgGrid(total_df,
           gridOptions=gbt.build(),
           theme='alpine',
           height=120)
    

    if total_mov2+nuevo_p < 0:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">Se a restado el valor de:  ${total_mov2+nuevo_p:} en {selec}</h3></div>', unsafe_allow_html=True)
    elif total_mov2+nuevo_p == 0:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">No se a realizado un incremento en el codificado de {selec}</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="max-width: 600px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">Se a incrementado el valor de:  ${total_mov2+nuevo_p:} en {selec}</h3></div>', unsafe_allow_html=True)

    #st.markdown("---")
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>9.- Escoja el proyecto, revise las metas y de ser necesario modifíquelas de ser necesario. </h5>", unsafe_allow_html=True)

    st.markdown(f"<h3 style='text-align: center; background-color: #f1f6f7; color: #080200'> Tabla de Metas de {selec} </h3>", unsafe_allow_html=True)
    # Mostrar la tabla con la extensión st_aggrid
    with st.expander("💱  Realizar modificaciones a las metas", expanded=False):
        gbmtd = GridOptionsBuilder.from_dataframe(df_mtt)
        gbmtd.configure_column('Proyecto', minWidth =408, editable=False )
        gbmtd.configure_column('Metas', minWidth =350, editable=False )
        gbmtd.configure_column('Nueva Meta', minWidth =300, editable=True )
        gbmtd.configure_column('Observación', minWidth =300, editable=True )
        

        edit_dfd = AgGrid(df_mtt,
                         gridOptions=gbmtd.build(),
                         height=350)
                            #reload_data=reload_data,)
        edit_dfd = pd.DataFrame(edit_dfd['data'])
        

    st.markdown('<hr style="border: 1.5px solid red; margin-bottom: 5px;">', unsafe_allow_html=True)

    if total_mov2+nuevo_p != -total_mov:
        st.markdown(f'<div style="max-width: auto; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor que se disminuyo en {direc} e incremento en {selec} es diferente, revizar información. </h3></div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div style="max-width: auto; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor que se disminuyo en {direc} e incremento en {selec} es el mismo.</h3></div>', unsafe_allow_html=True)

    st.markdown('<hr style="border: 1.5px solid red; margin-bottom: 5px;">', unsafe_allow_html=True)

    st.info("""
            Descargue adjunto:
                
            En el caso de que no haya anteriores revisiones, guarde la información y descargue el documento PDF para adjuntarlo al memorando de solicitud de reforma.
            * Recuerde que el memorando de solicitud debe contener las justificaciones técnicas y demás documentación de respaldo.

            """)
    
   

    try:
        edited_rows = edited_df[edited_df['Movimiento'] != 0]
        edited_rows2 = edi[edi['Movimiento'] != 0]
        edit_rows = edit_df[edit_df['Nueva Meta'] != '-']
        edit_rows2 = edit_dfd[edit_dfd['Nueva Meta'] != '-']
    except:
        st.write('No se realizaron cambios en la información')

    def descargar_xlsx(edited_rows,edited_rows2, edit_rows, edit_rows2, result):
              # Guardar los DataFrames en dos hojas de un archivo XLSX en memoria
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        edited_rows.to_excel(writer, sheet_name='Presupuesto (DE)', index=False)
                        edited_rows2.to_excel(writer, sheet_name='Presupuesto(PARA)', index=False)
                        edit_rows.to_excel(writer, sheet_name='Metas (DE)', index=False)
                        edit_rows2.to_excel(writer, sheet_name='Metas (PARA)' , index=False)
                        result.to_excel(writer, sheet_name='Nueva_partida', index=False)
                    output.seek(0)
                    return output

    export_as_pdf = st.columns(3)[1].button("Guardar información")
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
    
    dfp1 = pd.concat([nuevo_df, total_row])
    dfp4 = pd.concat([nuevo_df2, total_row2])
    
    # NUEVA PARTIDA
    result_filtro1 = ['Proyecto','Estructura','Incremento']
    result2 = result[result_filtro1]
    result2['Proyecto, Estructura']=result2['Proyecto']+ ' | ' + result2['Estructura']
    resul_filtro=['Proyecto, Estructura','Incremento']
    result2=result2[resul_filtro]
    sum_res=result2[['Incremento']].sum()
    total_res = pd.DataFrame([['Total', sum_res['Incremento']]], 
                            columns=['Proyecto, Estructura','Incremento'], index=['Total'])
    result2 = pd.concat([result2, total_res])
    dfp2=result2

    #Creamos una nueva tabla para las metas
    meta_filtro = ['Proyecto','Metas','Nueva Meta','Observación']
    dfp3 = edit_rows[meta_filtro]
    dfp5 = edit_rows2[meta_filtro]

    if total_mov2+nuevo_p == -(total_mov):
            if export_as_pdf:
                now = datetime.now()
                fecha_hora = now.strftime("%Y%m%d%H%M")    
                st.write('Descargando... ¡Espere un momento!')
                
                
                def export_to_pdf(dfp1, dfp2, dfp3,dfp4,dfp5):
                        # Crear un objeto BytesIO para almacenar el PDF
                    pdf_buffer = BytesIO()
    
                    doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
                            # Obtener estilos de texto predefinidos
                    styles = getSampleStyleSheet()
                    # Agregar título a la página
                    title = f"Reforma Externa-{fecha_hora}"
                    title_style = styles['Title']
                    title_style.spaceAfter = 12
                    title_style.spaceBefore = 22
                    title_paragraph = Paragraph(title, title_style)
                    
                    subtitle_text = f"De: {direc}-Presupuesto"
                    subtitle_paragraph = Paragraph(subtitle_text, title_style)

                    subtitle_text2 = f"Para: {selec}-Presupuesto"
                    subtitle2_paragraph = Paragraph(subtitle_text2, title_style)

                    title2 = f"Nueva Partida - {selec}"
                    title2_paragraph = Paragraph(title2, title_style)

                    title3 = f"{direc} - Reforma metas"
                    title3_paragraph = Paragraph(title3, title_style)         
                    
                    title4 = f"{selec} - Reforma metas"
                    title4_paragraph = Paragraph(title4, title_style)         

                    # Agregar imagen a la página
                    img_path = "logo GadPP.png"  # Reemplaza con la ruta de tu imagen
                    image = Image(img_path, width=100, height=100) 

                            # Convertir DataFrame a lista de listas para la tabla
                    data1 = [dfp1.columns.tolist()] + [configure_cell(dfp1, row) for _, row in dfp1.iterrows()]  # Agregar una fila con los nombres de las variables
                    # Crear la tabla con los datos del DataFrame
                    table1 = Table(data1, repeatRows=1, colWidths=[400] + [None] * (len(dfp1.columns) - 1))

                            # Establecer estilos para la tabla
                    table1.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinear el contenido al centro
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  # Fuente en negrita para la primera fila (encabezado)
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Agregar espacio inferior a la primera fila
                        ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  # Agregar bordes a la tabla
                        ('SPACEAFTER', (0, 0), (-1, -1), 6)  # Espacio después de cada fila
                    ]))

                    data4 = [dfp4.columns.tolist()] + [configure_cell(dfp4, row) for _, row in dfp4.iterrows()]  # Agregar una fila con los nombres de las variables
        # Crear la tabla con los datos del DataFrame
                    table4 = Table(data4, repeatRows=1, colWidths=[400] + [None] * (len(dfp4.columns) - 1))

                            # Establecer estilos para la tabla
                    table4.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinear el contenido al centro
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  # Fuente en negrita para la primera fila (encabezado)
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Agregar espacio inferior a la primera fila
                        ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  # Agregar bordes a la tabla
                        ('SPACEAFTER', (0, 0), (-1, -1), 6)  # Espacio después de cada fila
                    ]))

                    data2 = [dfp2.columns.tolist()] + [configure_cell(dfp2, row) for _, row in dfp2.iterrows()]
                    # Crear la segunda tabla con los datos del DataFrame 2
                    table2 = Table(data2, repeatRows=1, colWidths=[400] + [None] * (len(dfp2.columns) - 1))

                    # Establecer estilos para la segunda tabla
                    table2.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                        ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                        ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                    ]))

                    data3 = [dfp3.columns.tolist()] + [configure_cellmetas(dfp3, row) for _, row in dfp3.iterrows()]
                    # Crear la segunda tabla con los datos del DataFrame 2
                    table3 = Table(data3, repeatRows=1, colWidths=[200] + [None] * (len(dfp3.columns) - 1))

                    # Establecer estilos para la segunda tabla
                    table3.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                        ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                        ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                    ]))
                    data5 = [dfp5.columns.tolist()] + [configure_cellmetas(dfp5, row) for _, row in dfp5.iterrows()]
                    # Crear la segunda tabla con los datos del DataFrame 2
                    table5 = Table(data5, repeatRows=1, colWidths=[200] + [None] * (len(dfp5.columns) - 1))

                    # Establecer estilos para la segunda tabla
                    table5.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                        ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                        ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                    ]))
                # Co
                # Construir el PDF con la tabla
                    frames = [Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')]
                    template = PageTemplate(id='encabezado_izquierdo', frames=frames, onPage=lambda canvas, doc, **kwargs: canvas.drawImage(img_path, doc.leftMargin-40, doc.height+55, width=120, height=90, preserveAspectRatio=True))
                    doc.addPageTemplates([template])
                    doc.build([title_paragraph, subtitle_paragraph, table1,title3_paragraph, table3, subtitle2_paragraph, table4,title2_paragraph, table2,title4_paragraph,table5])
                    # Obtener el contenido del BytesIO
                    pdf_content = pdf_buffer.getvalue()
                            # Cerrar el BytesIO
                    pdf_buffer.close()

                    return pdf_content

                def configure_cell(df, row):
                    styles = getSampleStyleSheet()
                    styles['Normal'].fontSize = 8
                    first_col_text = str(row[df.columns[0]])
                    first_col_text = '\n'.join([first_col_text[j:j+20] for j in range(0, len(first_col_text), 20)])
                    return [Paragraph(first_col_text, styles['Normal'], encoding='utf-8')] + [str(row[col]) for col in df.columns[1:]]
            
                def configure_cellmetas(df, row):
                    styles = getSampleStyleSheet()
                    styles['Normal'].fontSize = 8
                    # Aplicar el formato con '\n' a todas las columnas
                    formatted_columns = [('\n'.join(str(row[col])[j:j+20] for j in range(0, len(str(row[col])), 20)), styles['Normal']) for col in df.columns]
                    # Devolver una lista de objetos Paragraph para cada celda
                    return [Paragraph(text, style, encoding='utf-8') for text, style in formatted_columns]

        

                if export_as_pdf:
                            # Llamar a la función para exportar DataFrame a PDF
                    pdf_content = export_to_pdf(dfp1, dfp2, dfp3,dfp4,dfp5)
                                    # Descargar el PDF
                    st.download_button('Descargar PDF', pdf_content, file_name='tabla_exportada.pdf', key='download_button')
                                    # Mensaje de éxito
                    st.success('Tabla exportada y PDF descargado exitosamente.') 
                
                archivo_xlsx = descargar_xlsx(edited_rows,edited_rows2, edit_rows,edit_rows2, result)
                st.download_button(
                        label="Haz clic para descargar",
                        data=archivo_xlsx.read(),
                        key="archivo_xlsx",
                        file_name=f"Reforma_{fecha_hora}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                    

    else:
            st.warning('La disminución e incremento realizadas tienen incosistencias revisar para descargar.')
            #st.markdown('<div style=" margin: 0 auto; background-color:#D9FDFF; padding:10px; text-align: center;"><h4 style="color:#01464A;">Los Movimientos tienen incosistencia revisar para descargar.</h3></div>', unsafe_allow_html=True)

if paginas == 5:
     
    df_odoo = pd.DataFrame(odoo)
    df_mt = pd.DataFrame(metas)
    #ENCABEZADO
    st.markdown("<h1 style='text-align:center;background-color: #000045; color: #ffffff'>🔄 REFORMA LIBERACIÓN </h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; background-color: #f0efeb; color: #080200'> Se libera el presupuesto de la Dirección/Unidad </h4>", unsafe_allow_html=True)
    st.markdown("")        
    with st.expander("💱  Resumen", expanded=False):
        st.markdown("""

                    Corresponde a la Reforma al POA en que el valor codificado de la Unidad no se modifica, las reformas/modificaciones, se realizan únicamente entre actividades de proyectos de la misma Unidad, y afecta:
                    
                    - A la **programación presupuestaria**, por incremento o disminución de los valores codificados de las actividades de los proyectos; y/o
                    - A la **programación física**, por modificación o no de las metas de los proyectos.

                    Una vez realizada la reforma al POA, presupuestaria y/o de metas, se procede a guardar la información; automáticamente se generará un archivo pdf codificado con la información de las modificaciones realizadas ya sea solo presupuestaria y/o de metas.
                    
                    """
                )
        st.info(""" **NOTA**: Las modificaciones se harán sobre los saldos disponibles no comprometidos de las asignaciones. """)
    
    st.markdown("<h2 style='text-align: center; background-color: #26469C; color: #ffffff'> Pasos para la reforma </h2>", unsafe_allow_html=True)
    #st.markdown("<h2 style='text-align: left;  color: #ccccc'>Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E'>1.- Escoga la unidad en la cual desea realizar la reforma </h5>", unsafe_allow_html=True)
    
    #FILTRAMOS SOLO PAI
    df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
    df_odoo['Codificado'] = df_odoo['Codificado'].round(2)
    #AGRUPAMOS LAS UNIDADES
    direc = st.selectbox('', options=df_odoo['Unidad'].unique())
    #FILTRAMOS COLUMNAS 
    df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()
    df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mt= df_mt.loc[df_mt.Unidad == direc]
    df_mtfil = ['Proyecto','Metas','Nueva Meta','Observación']
    df_mt = df_mt[df_mtfil]        
    df = pd.DataFrame(df_od)
    df = agregar_columnas(df)
    #SUBTITULOS
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>2.- Escoja el Proyecto y ajuste los valores en la casilla de “Incremento / Disminución”. </h5>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Presupuesto de {direc} </h3>", unsafe_allow_html=True, help="Se realiza cambios en la columna Movimientos, tomar encuenta el saldo disponible para restar un valor caso contrario no se tomara encuenta la reforma.")
    #FORMATO DE COLUMNAS
    gb = GridOptionsBuilder.from_dataframe(df) 
    gb.configure_column('Unidad', hide=True)#, rowGroup=True, cellRenderer= "agGroupCellRenderer", )
    gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
    gb.configure_column('Estructura', header_name="Actividad")
    gb.configure_column(field ='Codificado', maxWidth=150, aggFunc="sum", valueFormatter="data.Codificado.toLocaleString('en-US');")
    gb.configure_column('Saldo_Disponible', header_name="Saldo", maxWidth=120, valueFormatter="data.Saldo_Disponible.toLocaleString('en-US');", aggFunc='sum' )
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
      
    gb.configure_column('Movimiento', header_name='Increm/Dismi' , editable= True ,type=['numericColumn'], cellStyle=cellsytle_jscode, maxWidth=120, valueFormatter="data.Movimiento.toLocaleString('en-US');")
    
    gb.configure_column('Nuevo Codificado',header_name='Nuevo Cod' , valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                        type=['numericColumn'],maxWidth=150, valueFormatter="data.Nuevo Codificado.toLocaleString('en-US');", aggFunc='sum', enableValue=True)
    gb.configure_column('TOTAL', hide=True)

    go = gb.build()
    
    go['alwaysShowHorizontalScroll'] = True
    go['scrollbarWidth'] = 1
    reload_data = False

    edited_df = AgGrid(
        df,
        editable= True,
        gridOptions=go,
        width=1000, 
        height=350, 
        fit_columns_on_grid_load=True,
        theme='streamlit',
        allow_unsafe_jscode=True,
        reload_data=reload_data
    )
    
    # Si se detectan cambios, actualiza el DataFrame
    if edited_df is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edited_df = pd.DataFrame(edited_df['data'])
        edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
    #AGREGAR NUEVA PARTIDA
    st.markdown(
        '''
        <style>
        .streamlit-expanderHeader {
            background-color: blue;
            color: black; # Adjust this for expander header color
        }
        .streamlit-expanderContent {
            background-color: blue;
            color: black; # Expander content color
        }
        </style>
        ''',
        unsafe_allow_html=True
    )

    #TOTALES
    total_cod = int(edited_df['Codificado'].sum())
    total_mov = int(edited_df['Movimiento'].sum())
    total_tot = int(edited_df['TOTAL'].sum())
    

    total_row = {
        'PROYECTO': 'Total', 
        'Total_Codificado': df['Codificado'].sum(),
        'Total_Saldo': df['Saldo_Disponible'].sum(),
        'Tot_Increm/Dismi': edited_df['Movimiento'].sum() ,
        'Total_Nuev_Codif': edited_df['TOTAL'].sum() 
    }
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>3.- Verifica que los valores totales estén correctos. </h5>", unsafe_allow_html=True)
    
    total_df = pd.DataFrame([total_row])
    gbt = GridOptionsBuilder.from_dataframe(total_df)
    gbt.configure_column('PROYECTO', minWidth =810 )
    gbt.configure_column('Total_Saldo', header_name='Total Saldo', maxWidth =120, valueFormatter="data.Total_Saldo.toLocaleString('en-US');" )
    gbt.configure_column('Tot_Increm/Dismi', header_name='Tot Incr/Dismi', maxWidth =140, valueFormatter="data.Tot_Increm/Dismi.toLocaleString('en-US');" )
    gbt.configure_column('Total_Nuev_Codif', header_name='Tot Nuev Cod', maxWidth =135, valueFormatter="data.Total_Nuev_Codif.toLocaleString('en-US');" )
    gbt.configure_column('Total_Codificado', header_name='Tot Codificado',  maxWidth =145, valueFormatter="data.Total_Codificado.toLocaleString('en-US');" )
   
    
    
    AgGrid(total_df,
            gridOptions=gbt.build(),
            theme='alpine',
            height=120)
    #st.markdown("---")
    if total_cod > total_tot:
        st.markdown('<div style="max-width: 900px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor total del codificado es mayor al nuevo codificado</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="max-width: 900px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor total del codificado es menor o igual al nuevo codificado por tanto no se esta liberando el presupuesto</h3></div>', unsafe_allow_html=True)


    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>4.- Escoja el proyecto, revise las metas y de ser necesario modifíquelas de ser necesario. </h5>", unsafe_allow_html=True)

    st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Metas de {direc} </h3>", unsafe_allow_html=True, help="En la columna Nueva Meta se puede agregar la modificación a la meta actual")
    # Mostrar la tabla con la extensión st_aggrid
    #with st.expander(f"🆕  Modificar metas de los proyectos de {direc}", expanded=False): 
    gbmt = GridOptionsBuilder.from_dataframe(df_mt)
    gbmt.configure_column('Proyecto', minWidth =408, editable=False )
    gbmt.configure_column('Metas', minWidth =350, editable=False )
    gbmt.configure_column('Nueva Meta', minWidth =300, editable=True )
    gbmt.configure_column('Observación', minWidth =300, editable=True )
    

    edit_df = AgGrid(df_mt,
                        gridOptions=gbmt.build(),
                        height=350)
    edit_df = pd.DataFrame(edit_df['data'])


    try:
        edited_rows = edited_df[edited_df['Movimiento'] != 0]
        edit_rows = edit_df[edit_df['Nueva Meta'] != '-']
    except:
        st.write('No se realizaron cambios en de la información')

    st.markdown('<hr style="border: 1.5px solid red; margin-bottom: 5px;">', unsafe_allow_html=True)

    st.info("""
            Descargue adjunto:
                
            En el caso de que no haya anteriores revisiones, guarde la información y descargue el documento PDF para adjuntarlo al memorando de solicitud de reforma.
            * Recuerde que el memorando de solicitud debe contener las justificaciones técnicas y demás documentación de respaldo.

            """)
    

    st.markdown("---")
    def descargar_xlsx(edited_rows, edit_rows):
            # Guardar los DataFrames en dos hojas de un archivo XLSX en memoria
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    edited_rows.to_excel(writer, sheet_name='Presupuesto', index=False)
                    edit_rows.to_excel(writer, sheet_name='Metas', index=False)
                output.seek(0)
                return output
    export_as_pdf = st.columns(3)[1].button("Guardar información")
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
    dfp1 = nuevo_df
    #Creamos una nueva tabla para las metas
    meta_filtro = ['Proyecto','Metas','Nueva Meta','Observación']
    dfp3 = edit_rows[meta_filtro]
    
    if total_cod > total_tot:
        if export_as_pdf:
            now = datetime.now()
            fecha_hora = now.strftime("%Y%m%d%H%M")    
            st.write('Descargando... ¡Espere un momento!')
            
            
            def export_to_pdf(dfp1, dfp3):
                # Crear un objeto BytesIO para almacenar el PDF
                pdf_buffer = BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
                # Obtener estilos de texto predefinidos
                styles = getSampleStyleSheet()
                # Agregar título a la página
                title = f"Reforma Liberación - {fecha_hora}"
                title_style = styles['Title']
                title_style.spaceAfter = 12
                title_style.spaceBefore = 22
                title_paragraph = Paragraph(title, title_style)
                                                    
                subtitle_text = f"{direc} - Presupuesto"
                subtitle_paragraph = Paragraph(subtitle_text, title_style)

                title3 = f"{direc} - Metas"
                title3_paragraph = Paragraph(title3, title_style) 
                
                # Agregar imagen a la página
                img_path = "logo GadPP.png"  # Reemplaza con la ruta de tu imagen
                image = Image(img_path, width=100, height=100) 

                # Convertir DataFrame a lista de listas para la tabla
                data1 = [dfp1.columns.tolist()] + [configure_cell(dfp1, row) for _, row in dfp1.iterrows()]  # Agregar una fila con los nombres de las variables
                # Crear la tabla con los datos del DataFrame
                table1 = Table(data1, repeatRows=1, colWidths=[400] + [None] * (len(dfp1.columns) - 1))

                # Establecer estilos para la tabla
                table1.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinear el contenido al centro
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  # Fuente en negrita para la primera fila (encabezado)
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Agregar espacio inferior a la primera fila
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  # Agregar bordes a la tabla
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  # Espacio después de cada fila
                ]))

                data3 = [dfp3.columns.tolist()] + [configure_cellmetas(dfp3, row) for _, row in dfp3.iterrows()]
                # Crear la segunda tabla con los datos del DataFrame 2
                table3 = Table(data3, repeatRows=1, colWidths=[200] + [None] * (len(dfp3.columns) - 1))

                # Establecer estilos para la segunda tabla
                table3.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                ]))
                # Construir el PDF con la tabla
                frames = [Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')]
                template = PageTemplate(id='encabezado_izquierdo', frames=frames, onPage=lambda canvas, doc, **kwargs: canvas.drawImage(img_path, doc.leftMargin-40, doc.height+55, width=120, height=90, preserveAspectRatio=True))
                doc.addPageTemplates([template])
                doc.build([title_paragraph, subtitle_paragraph, table1,title3_paragraph, table3])
                # Obtener el contenido del BytesIO
                pdf_content = pdf_buffer.getvalue()
                # Cerrar el BytesIO
                pdf_buffer.close()

                return pdf_content

            def configure_cell(df, row):
                styles = getSampleStyleSheet()
                styles['Normal'].fontSize = 8
                first_col_text = str(row[df.columns[0]])
                first_col_text = '\n'.join([first_col_text[j:j+20] for j in range(0, len(first_col_text), 20)])
                return [Paragraph(first_col_text, styles['Normal'], encoding='utf-8')] + [str(row[col]) for col in df.columns[1:]]
        
            def configure_cellmetas(df, row):
                styles = getSampleStyleSheet()
                styles['Normal'].fontSize = 8
                # Aplicar el formato con '\n' a todas las columnas
                formatted_columns = [('\n'.join(str(row[col])[j:j+20] for j in range(0, len(str(row[col])), 20)), styles['Normal']) for col in df.columns]
                # Devolver una lista de objetos Paragraph para cada celda
                return [Paragraph(text, style, encoding='utf-8') for text, style in formatted_columns]

    

            if export_as_pdf:
                # Llamar a la función para exportar DataFrame a PDF
                pdf_content = export_to_pdf(dfp1, dfp3)
                # Descargar el PDF
                st.download_button('Descargar PDF', pdf_content, file_name='tabla_exportada.pdf', key='download_button')
                # Mensaje de éxito
                st.success('Tabla exportada y PDF descargado exitosamente.') 

            archivo_xlsx = descargar_xlsx(edited_rows, edit_rows)
            st.download_button(
                label="Haz clic para descargar",
                data=archivo_xlsx.read(),
                key="archivo_xlsx",
                file_name=f"Reforma_{fecha_hora}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            
            
    else:
        st.warning('No se puede descargar, porque no se esta liberando el presupuesto')    

if paginas == 6:
    #CARGAMOS LAS BASES
    df_odoo = pd.DataFrame(odoo)
    df_mt = pd.DataFrame(metas)
    #ENCABEZADO
    st.markdown("<h1 style='text-align:center;background-color: #000045; color: #ffffff'>➕ REFORMA AL POA POR INCREMENTO DE PRESUPUESTO</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; background-color: #f0efeb; color: #080200'> Se solicita presupuesto a la Institución</h4>", unsafe_allow_html=True)
    st.markdown("")
    with st.expander("💱  Resumen", expanded=False):
        st.markdown("""

                    Corresponde a la Reforma al POA en que el valor codificado de la Unidad no se modifica, las reformas/modificaciones, se realizan únicamente entre actividades de proyectos de la misma Unidad, y afecta:
                    
                    - A la **programación presupuestaria**, por incremento o disminución de los valores codificados de las actividades de los proyectos; y/o
                    - A la **programación física**, por modificación o no de las metas de los proyectos.

                    Una vez realizada la reforma al POA, presupuestaria y/o de metas, se procede a guardar la información; automáticamente se generará un archivo pdf codificado con la información de las modificaciones realizadas ya sea solo presupuestaria y/o de metas.
                    
                    """
                )
        st.info(""" **NOTA**: Las modificaciones se harán sobre los saldos disponibles no comprometidos de las asignaciones. """)
    
    st.markdown("<h2 style='text-align: center; background-color: #26469C; color: #ffffff'> Pasos para la reforma </h2>", unsafe_allow_html=True)
    #st.markdown("<h2 style='text-align: left;  color: #ccccc'>Pasos para la reforma </h2>", unsafe_allow_html=True)
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E'>1.- Escoga la unidad en la cual desea realizar la reforma </h5>", unsafe_allow_html=True)
    
        #reload---
    reload_data = False
    #FILTRAMOS SOLO PAI
    df_odoo= df_odoo.loc[df_odoo['PAI/NO PAI'] == 'PAI']
    df_odoo['Codificado'] = df_odoo['Codificado'].round(2)
    #AGRUPAMOS LAS UNIDADES
    direc = st.selectbox('', options=df_odoo['Unidad'].unique())
    #FILTRAMOS COLUMNAS 
    df_od= df_odoo.loc[df_odoo.Unidad == direc].groupby(['Unidad','PROYECTO','Código','Estructura'], as_index= False)[['Codificado', 'Saldo Disponible']].sum()
    df_od.rename(columns = {'Saldo Disponible': 'Saldo_Disponible'}, inplace= True)
    df_mt= df_mt.loc[df_mt.Unidad == direc]
    df_mtfil = ['Proyecto','Metas','Nueva Meta','Observación']
    df_mt = df_mt[df_mtfil]        
    df = pd.DataFrame(df_od)
    df = agregar_columnas(df)
    
    
    #SUBTITULOS
    #st.markdown(f"<h2 style='text-align:center;'> {direc} </h2>", unsafe_allow_html=True)
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>2.- Escoja el Proyecto y ajuste los valores en la casilla de “Incremento / Disminución”. </h5>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Presupuesto de {direc} </h3>", unsafe_allow_html=True, help="Se realiza cambios en la columna Movimientos, tomar encuenta el saldo disponible para restar un valor caso contrario no se tomara encuenta la reforma.")
    #FORMATO DE COLUMNAS
    gb = GridOptionsBuilder.from_dataframe(df) 
    gb.configure_column('Unidad', hide=True)#, rowGroup=True, cellRenderer= "agGroupCellRenderer", )
    gb.configure_column('PROYECTO',header_name="PROYECTO", hide=True, rowGroup=True)
    gb.configure_column('Estructura', header_name="Actividad")
    gb.configure_column(field ='Codificado', maxWidth=150, aggFunc="sum", valueFormatter="data.Codificado.toLocaleString('en-US');")
    gb.configure_column('Saldo_Disponible', header_name="Saldo", maxWidth=120, valueFormatter="data.Saldo_Disponible.toLocaleString('en-US');", aggFunc='sum' )
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
                return null
            }
            else{
                alert("NO SE PUEDE AGREGAR UN VALOR NEGATIVO");

            }
            
        }
    """)    
    gb.configure_column('Movimiento', header_name='Increm/Dismi' , editable= True ,type=['numericColumn'], cellStyle=cellsytle_jscode, maxWidth=120, valueFormatter="data.Movimiento.toLocaleString('en-US');")
    
    gb.configure_column('Nuevo Codificado',header_name='Nuevo Cod' , valueGetter='Number(data.Codificado) + Number(data.Movimiento)', cellRenderer='agAnimateShowChangeCellRenderer',
                        type=['numericColumn'],maxWidth=150, valueFormatter="data.Nuevo Codificado.toLocaleString('en-US');", aggFunc='sum', enableValue=True)
    gb.configure_column('TOTAL', hide=True)

    go = gb.build()
    
    go['alwaysShowHorizontalScroll'] = True
    go['scrollbarWidth'] = 1
    reload_data = False


    edited_df = AgGrid(
        df,
        editable= True,
        gridOptions=go,
        width=1000, 
        height=350, 
        fit_columns_on_grid_load=True,
        theme='streamlit',
        allow_unsafe_jscode=True,
        reload_data=reload_data
    )
    
    # Si se detectan cambios, actualiza el DataFrame
    if edited_df is not None:
        # Convierte el objeto AgGridReturn a DataFrame
        edited_df = pd.DataFrame(edited_df['data'])
        edited_df['TOTAL'] = edited_df['Codificado'] + edited_df['Movimiento']
        #AGREGAR NUEVA PARTIDA
    st.markdown(
        '''
        <style>
        .streamlit-expanderHeader {
            background-color: blue;
            color: black; # Adjust this for expander header color
        }
        .streamlit-expanderContent {
            background-color: blue;
            color: black; # Expander content color
        }
        </style>
        ''',
        unsafe_allow_html=True
    )
    
                        #reload_data=reload_data,)
    #edit_df = pd.DataFrame(edit_df['data'])
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>3.- Crear nuevas actividades presupuestarias: Seleccione el proyecto, coloque el nombre de la partida y el valor a incrementar.</h5>", unsafe_allow_html=True)
    with st.expander(f"🆕  Crear una partida nueva para {direc} ", expanded=False): 
        st.markdown("<p style='text-align: center; background-color: #B5E6FC;'> Agregar nueva partida </p>", unsafe_allow_html=True)
        dfnuevop = pd.DataFrame(columns=['Proyecto','Estructura','Incremento','Parroquia'])
        #colors = st.selectbox('Escoja la Unidad', options=df_odoo['Unidad'].unique())
        config = {
            'Proyecto' : st.column_config.SelectboxColumn('Proyecto',width='large', options=df_od['PROYECTO'].unique()),
            'Estructura' : st.column_config.TextColumn('Estructura', width='large', required=True),
            'Incremento' : st.column_config.NumberColumn('Incremento', min_value=0,width='medium', required=True),
            'Parroquia' : st.column_config.TextColumn('Parroquia', width='medium', required=True)
        }

        result = st.data_editor(dfnuevop, column_config = config, num_rows='dynamic')

        if st.button('Crear partida:'):
            st.write(result)
    #TOTALES
    total_cod = int(edited_df['Codificado'].sum())
    total_mov = int(edited_df['Movimiento'].sum())
    total_tot = int(edited_df['TOTAL'].sum())
    nuevo_p = int(result['Incremento'].sum())

    total_row = {
        'PROYECTO': 'Total',  # No se calcula el total para la columna de texto
        'Total_Codificado': df['Codificado'].sum(),
        'Total_Saldo': df['Saldo_Disponible'].sum(),
        'Tot_Increm/Dismi': edited_df['Movimiento'].sum() + nuevo_p,
        'Total_Nuev_Codif':  nuevo_p + edited_df['TOTAL'].sum() 
    }
    total_df = pd.DataFrame([total_row])
    gbt = GridOptionsBuilder.from_dataframe(total_df)
    gbt.configure_column('PROYECTO', minWidth =810 )
    gbt.configure_column('Total_Saldo', header_name='Total Saldo', maxWidth =120, valueFormatter="data.Total_Saldo.toLocaleString('en-US');" )
    gbt.configure_column('Tot_Increm/Dismi', header_name='Tot Incr/Dismi', maxWidth =140, valueFormatter="data.Tot_Increm/Dismi.toLocaleString('en-US');" )
    gbt.configure_column('Total_Nuev_Codif', header_name='Tot Nuev Cod', maxWidth =135, valueFormatter="data.Total_Nuev_Codif.toLocaleString('en-US');" )
    gbt.configure_column('Total_Codificado', header_name='Tot Codificado',  maxWidth =145, valueFormatter="data.Total_Codificado.toLocaleString('en-US');" )
    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>4.- Verifica que los valores totales estén correctos. </h5>", unsafe_allow_html=True)
    
    AgGrid(total_df,
            gridOptions=gbt.build(),
            theme='alpine',
            height=120)
    
    if total_cod < total_tot+nuevo_p:
        st.markdown('<div style="max-width: 900px; margin: 0 auto; background-color:#F1FFEF; padding:10px; text-align: center;"><h4 style="color:#008000;">El valor total del codificado es menor al nuevo codificado</h3></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="max-width: 900px; margin: 0 auto; background-color:#ffcccc; padding:10px; text-align: center;"><h4 style="color:#ff0000;">El valor total del codificado es mayor o igual al nuevo codificado</h3></div>', unsafe_allow_html=True)
    
    #st.markdown("---")

    st.markdown('<hr style="border: 1.5px solid blue; margin-bottom: 5px;">', unsafe_allow_html=True)
    st.markdown("<h5 style='text-align: left;  color: #00524E; background-color: #FfFFFF;'>5.- Escoja el proyecto, revise las metas y de ser necesario modifíquelas de ser necesario. </h5>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center; background-color: #DDDDDD;'> 🗄 Tabla de Metas de {direc} </h3>", unsafe_allow_html=True, help="En la columna Nueva Meta se puede agregar la modificación a la meta actual")
    
    gbmt = GridOptionsBuilder.from_dataframe(df_mt)
    gbmt.configure_column('Proyecto', minWidth =408, editable=False )
    gbmt.configure_column('Metas', minWidth =350, editable=False )
    gbmt.configure_column('Nueva Meta', minWidth =300, editable=True )
    gbmt.configure_column('Observación', minWidth =300, editable=True )

    edit_df = AgGrid(df_mt,
                        gridOptions=gbmt.build(),
                        height=350)

    edit_df = pd.DataFrame(edit_df['data'])

        
    try:
        edited_rows = edited_df[edited_df['Movimiento'] != 0]
        edit_rows = edit_df[edit_df['Nueva Meta'] != '-']
    except:
        st.write('No se realizaron cambios en de la información')

    st.markdown('<hr style="border: 1.5px solid red; margin-bottom: 5px;">', unsafe_allow_html=True)

    st.info("""
            Descargue adjunto:
                
            En el caso de que no haya ulteriores revisiones, guarde la información y descargue el documento PDF para adjuntarlo al memorando de solicitud de reforma.
            * Recuerde que el memorando de solicitud debe contener las justificaciones técnicas y demás documentación de respaldo.

            """)

    st.markdown("---")
    #st.markdown(type(Codificado))
    def descargar_xlsx(edited_rows, edit_rows, result):
            # Guardar los DataFrames en dos hojas de un archivo XLSX en memoria
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    edited_rows.to_excel(writer, sheet_name='Presupuesto', index=False)
                    edit_rows.to_excel(writer, sheet_name='Metas', index=False)
                    result.to_excel(writer, sheet_name='Nueva_partida', index=False)
                output.seek(0)
                return output
    #if st.columns(3)[1].button("click me")
    export_as_pdf = st.columns(3)[1].button("Guardar información")
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
    dfp1 = nuevo_df
    #Creamos una nueva tabla para las metas
    meta_filtro = ['Proyecto','Metas','Nueva Meta','Observación']
    dfp3 = edit_rows[meta_filtro]
    result2 = result
    result2['Proyecto, Estructura']=result2['Proyecto']+ ' | ' + result2['Estructura']
    resul_filtro=['Proyecto, Estructura','Incremento']
    result2=result2[resul_filtro]
    sum_res=result2[['Incremento']].sum()
    total_res = pd.DataFrame([['Total', sum_res['Incremento']]], 
                        columns=['Proyecto, Estructura','Incremento'], index=['Total'])
    result2 = pd.concat([result2, total_res])
    dfp2=result2


    if total_cod < total_tot+nuevo_p:
        if export_as_pdf:
            now = datetime.now()
            fecha_hora = now.strftime("%Y%m%d%H%M")    
            st.write('Descargando... ¡Espere un momento!')
            
            
            def export_to_pdf(dfp1, dfp2, dfp3):
                    # Crear un objeto BytesIO para almacenar el PDF
                pdf_buffer = BytesIO()
                        # Crear un objeto SimpleDocTemplate para el PDF
                #doc = SimpleDocTemplate("output.pdf", pagesize=custom_page_size, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)

                doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
                        # Obtener estilos de texto predefinidos
                styles = getSampleStyleSheet()
                # Agregar título a la página
                
                title = f"Reforma Solicitud-{fecha_hora}"
                title_style = styles['Title']
                title_style.spaceAfter = 12
                title_style.spaceBefore = 22
                title_paragraph = Paragraph(title, title_style)
                
                subtitle_text = f"{direc} - Presupuesto"
                subtitle_paragraph = Paragraph(subtitle_text, title_style)

                title2 = f"{direc} - Nueva Partida"
                title2_paragraph = Paragraph(title2, title_style)

                title3 = f"{direc} - Metas"
                title3_paragraph = Paragraph(title3, title_style) 

                
                # Agregar imagen a la página
                img_path = "logo GadPP.png"  # Reemplaza con la ruta de tu imagen
                image = Image(img_path, width=100, height=100) 

                        # Convertir DataFrame a lista de listas para la tabla
                data1 = [dfp1.columns.tolist()] + [configure_cell(dfp1, row) for _, row in dfp1.iterrows()]  # Agregar una fila con los nombres de las variables
    # Crear la tabla con los datos del DataFrame
                table1 = Table(data1, repeatRows=1, colWidths=[400] + [None] * (len(dfp1.columns) - 1))

                        # Establecer estilos para la tabla
                table1.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alinear el contenido al centro
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  # Fuente en negrita para la primera fila (encabezado)
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Agregar espacio inferior a la primera fila
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  # Agregar bordes a la tabla
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  # Espacio después de cada fila
                ]))

                data2 = [dfp2.columns.tolist()] + [configure_cell(dfp2, row) for _, row in dfp2.iterrows()]
                # Crear la segunda tabla con los datos del DataFrame 2
                table2 = Table(data2, repeatRows=1, colWidths=[400] + [None] * (len(dfp2.columns) - 1))

                # Establecer estilos para la segunda tabla
                table2.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                ]))

                data3 = [dfp3.columns.tolist()] + [configure_cellmetas(dfp3, row) for _, row in dfp3.iterrows()]
                # Crear la segunda tabla con los datos del DataFrame 2
                table3 = Table(data3, repeatRows=1, colWidths=[200] + [None] * (len(dfp3.columns) - 1))

                # Establecer estilos para la segunda tabla
                table3.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),  
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  
                    ('GRID', (0, 0), (-1, -1), 0.7, 'black'),  
                    ('SPACEAFTER', (0, 0), (-1, -1), 6)  
                ]))
            # Construir el PDF con la tabla
                frames = [Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')]
                template = PageTemplate(id='encabezado_izquierdo', frames=frames, onPage=lambda canvas, doc, **kwargs: canvas.drawImage(img_path, doc.leftMargin-40, doc.height+55, width=120, height=90, preserveAspectRatio=True))
                doc.addPageTemplates([template])
                doc.build([title_paragraph, subtitle_paragraph, table1,title2_paragraph, table2,title3_paragraph, table3])
                        # Obtener el contenido del BytesIO
                pdf_content = pdf_buffer.getvalue()
                        # Cerrar el BytesIO
                pdf_buffer.close()

                return pdf_content

            def configure_cell(df, row):
                styles = getSampleStyleSheet()
                styles['Normal'].fontSize = 8
                first_col_text = str(row[df.columns[0]])
                first_col_text = '\n'.join([first_col_text[j:j+20] for j in range(0, len(first_col_text), 20)])
                return [Paragraph(first_col_text, styles['Normal'], encoding='utf-8')] + [str(row[col]) for col in df.columns[1:]]
        
            def configure_cellmetas(df, row):
                styles = getSampleStyleSheet()
                styles['Normal'].fontSize = 8
                # Aplicar el formato con '\n' a todas las columnas
                formatted_columns = [('\n'.join(str(row[col])[j:j+20] for j in range(0, len(str(row[col])), 20)), styles['Normal']) for col in df.columns]
                # Devolver una lista de objetos Paragraph para cada celda
                return [Paragraph(text, style, encoding='utf-8') for text, style in formatted_columns]

    

            if export_as_pdf:
                        # Llamar a la función para exportar DataFrame a PDF
                pdf_content = export_to_pdf(dfp1, dfp2, dfp3)
                                # Descargar el PDF
                st.download_button('Descargar PDF', pdf_content, file_name='tabla_exportada.pdf', key='download_button')
                                # Mensaje de éxito
                st.success('Tabla exportada y PDF descargado exitosamente.') 

            #pdf.output(pdf_output)
            archivo_xlsx = descargar_xlsx(edited_rows, edit_rows, result)
            st.download_button(
                label="Haz clic para descargar",
                data=archivo_xlsx.read(),
                key="archivo_xlsx",
                file_name=f"Reforma_{fecha_hora}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            
            
    else:
        #st.markdown('<div style=" margin: 0 auto; background-color:#D9FDFF; padding:10px; text-align: center;"><h4 style="color:#01464A;">Los Movimientos tienen incosistencia revisar para descargar.</h3></div>', unsafe_allow_html=True)
        st.warning('Los Movimientos tienen incosistencia revisar para descargar')    


