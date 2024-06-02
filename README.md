# Data ETL Pipeline: Python Powered Excel Wizard

## Objetivo
Este proyecto responde a una necesidad del grupo de Tribología, en la que eran incapaces de trabajar con archivos que guardaban los datos de presiones parciales dentro de la cámara de síntesis de PVD, donde realizan los recubrimientos de substratos metálicos para proyectos propios o de colaboración con entidades externas. La idea del código es que sea capaz de leer el archivo .csv en el que el programa de medición ha grabado los datos, transformarlos en acuerdo a las necesidades del proyecto, y volcarlos en un archivo .xlsx donde ya estarán dibujadas las gráficas pertinentes.

## Habilidades puestas en juego
- Programación con Python
- Uso de librerías como Pandas y Openpy
- Creación de un pipelane ETL
- Identificación de áreas de mejora
- Automatización de procesos
- Resolución de problemas

## Explicación del código
Paso 1. Comenzamos importando las librerías que nos van a hacer falta

    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    from openpyxl.chart import BarChart, Reference, DoughnutChart
    from openpyxl.utils import get_column_letter
    import string

Paso 2. Cargamos los archivos que queramos manipular

    dfc1=pd.read_csv(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\OPX-S-0210_13-10-2022 10-28-49.csv')
    dfc2=pd.read_csv(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\OPX-S-0210_13-10-2022 10-34-46.csv')
    dfc3=pd.read_csv(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\OPX-S-0210_13-10-2022 11-34-45.csv')
    dfc4=pd.read_csv(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\OPX-S-0210_13-10-2022 12-12-46.csv')

Paso 3. Creamos un DataFrame (DF) que agrupe todos los archivos en uno solo y creamos un 2º DF donde seeccionamos las columnas que sean necesarias según las especificaciones dadas

    df=pd.concat([dfc1, dfc2, dfc3, dfc4], sort=False)
    df2=pd.DataFrame(df, columns=['Time',' N2+ (RGA)',' O (RGA)',' Ar (RGA)',' N2 (RGA)'])

Paso 4. Hacemos un bucle for que nos pida como entrada un par de horas (timestamps) para poder diferenciar entre diferentes momentos del experimento.

    capa = [0, 1, 2, 3, 4, 5, 6, 7]
    for i in capa:
    a=input('Introduce la hora a la que empezó la capa: ')
    b=input('Ahora, introduce la hora a la que terminó: ')
    capa[i] = df2[(df2['Time'] >= a) & (df2['Time'] <= b) & (df2[' Ar (RGA)'] != df2[' N2 (RGA)'])] 
    promedio=capa[i].mean(skipna=True, numeric_only=True)
    dfp=pd.DataFrame(promedio, columns=['Promedio'])
    dfp=dfp.transpose()
    capa[i] = pd.concat([capa[i],dfp], sort=False) 
    print(capa[i])

Paso 5. Volcamos los datos en un fichero .xlsx

    with pd.ExcelWriter(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\Escalera_P.xlsx') as writer:
    capa[0].to_excel(writer, sheet_name='XXXXXX') 
    capa[1].to_excel(writer, sheet_name='XXXXXX')
    capa[2].to_excel(writer, sheet_name='XXXXXX')
    capa[3].to_excel(writer, sheet_name='XXXXXX')
    capa[4].to_excel(writer, sheet_name='XXXXXX')
    capa[5].to_excel(writer, sheet_name='XXXXXX')
    capa[6].to_excel(writer, sheet_name='XXXXXX')
    capa[7].to_excel(writer, sheet_name='XXXXXX')

Paso 6. Creamos los gráficos que creamos pertinentes.

    pd.read_excel(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\Escalera_P.xlsx')
    es = load_workbook(r'C:\Users\Garantia\Desktop\Calculos_Optix\Datos\Escalera_P.xlsx')
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']
    XXXXXX = es['XXXXXX']

    id_sheet = ['XXXXXX', 'XXXXXX', 'XXXXXX', 'XXXXXX', 
            'XXXXXX', 'XXXXXX', 'XXXXXX', 'XXXXXX']

    for i in range(len(id_sheet)):
    sheet = es.worksheets[i]
    min_col = sheet.min_column
    max_col = sheet.max_column
    min_fila = sheet.min_row
    max_fila = sheet.max_row
    print(sheet.title)
    print('Primera columna:', min_col)
    print('Ultima columna:', max_col)
    print('Primera fila:', min_fila)
    print('Ultima fila:', max_fila, '\n')

    abc = list(string.ascii_uppercase)
    abc_Ex = abc[0:max_col]
 
    for j in abc_Ex:
        if j!='A' and j!='B':
            sheet[f'{j}{max_fila+1}'] = f'=({j}{max_fila}/E{max_fila})'
    
    sheet[f'A{max_fila+1}'] = 'Ratios X/Ar'
    sheet[f'A{max_fila+1}'].font = Font('Calibri', bold=True, size=11)
     
    data = Reference(sheet,
                     min_col=min_col+2,  
                     max_col=max_col,
                     min_row=max_fila, 
                     max_row=max_fila)
    
    dataratio = Reference(sheet,
                     min_col=min_col+2, 
                     max_col=max_col,
                     min_row=max_fila+1, 
                     max_row=max_fila+1)
    
    categories = Reference(sheet,
                           min_col=min_col+2,
                           max_col=max_col, 
                           min_row=min_fila,
                           max_row=min_fila) 
                                             
    barchart=BarChart()
    barchart.add_data(data, from_rows=True)
    barchart.set_categories(categories)
     
    sheet.add_chart(barchart, 'H4')
    barchart.title = 'Presión promedio'
    barchart.y_axis.title = 'Presión (mbar)'
    barchart.legend = None
    barchart.x_axis.majorGridlines = None
    barchart.y_axis.majorGridlines = None
    barchart.style = 2
    
    barchartratio=BarChart()
    barchartratio.add_data(dataratio, from_rows=True)
    barchartratio.set_categories(categories)
    
    sheet.add_chart(barchartratio, 'P4')
    barchartratio.title = 'Ratios X/Ar'
    barchartratio.y_axis.title = 'Ratio (tanto x1)'
    barchartratio.legend = None
    barchartratio.x_axis.majorGridlines = None
    barchartratio.y_axis.majorGridlines = None
    barchartratio.style = 1
    
    donut = DoughnutChart(holeSize=50)
    donut.add_data(data, from_rows=True)
    donut.set_categories(categories)
    sheet.add_chart(donut, 'H19')
    donut.title = 'Presión promedio (mbar)'
    donut.style = 2
    
    es.save(r'C:\Users\Garantia\Desktop\Calculos_Optix\Excels\Escalera_P.xlsx')
