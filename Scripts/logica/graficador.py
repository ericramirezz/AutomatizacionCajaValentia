import os
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


#funciones de apoyo que se usan para calcular cosas antes de graficar

def _sem(valores):
    """
    calcula el error estandar de una lista de numeros
    Args: valores co latencias (list)
    Returns: error estandar de las latencias (float)
    """
    n = len(valores)
    if n < 2:
        #si hay menos de 2 valores no tiene sentido calcularlo
        return 0.0
    s = pd.Series(valores)
    return float(s.std() / math.sqrt(n))


def _tercio_split(serie):
    """
    divide los ensayos en 3 bloques lo mas iguales posible, le da prioridad de tamaño al primer grupo,
    por ejemplo para 29, el primer bloque tiene 10, el segundo 10 y el tercero 9
    Args: valores de latencia (Series)
    Returns: lista con 3 Series, una por bloque (list)"""
    n = len(serie)
    if n == 0: #si no hay datos devolvemos tres series vacias
        return [pd.Series(dtype=float)] * 3 
    g1 = math.ceil(n / 3) #sacamos el tamaño del primer grupo
    rem = n - g1 #calculamos lo que queda
    g2 = math.ceil(rem / 2) if rem > 0 else 0 #definimos el segundo grupo con lo que resta
    return [ 
        serie.iloc[:g1].reset_index(drop=True), #cortamos la primera parte
        serie.iloc[g1: g1 + g2].reset_index(drop=True), #cortamos la parte del medio
        serie.iloc[g1 + g2:].reset_index(drop=True), #cortamos el resto al final
    ]


def _x_tercio(dia, tercio):
    """
    dice en que posicion horizontal va cada punto de la grafica
    Args: dia -> numero del dia (entero), tercio -> numero del bloque (1, 2 o 3) (int)
    Returns: posicion x del punto (float)
    """
    offsets = {1: -0.25, 2: 0.0, 3: 0.25} #definimos los ajustes de posicion para cada tercio
    return dia + offsets[tercio] #regresamos el valor base sumado al ajuste del tercio correspondiente


#funcion principal, la que lee los archivos y arma la grafica

def generar_grafica(archivos, log_fn=print):
    """
    lee los archivos excel de cada dia, calcula promedios por bloque
    y genera la grafica con barras de error
    Args: archivos -> lista de rutas a los archivos .xlsx (uno por dia)
        log_fn   -> funcion para mostrar mensajes (por defecto print)
    Returns: (fig, df_resumen, dias_ratas_raw) si hay datos, o (None, None, None) si no
    """
    filas_resumen = []
    xs_seguro, ys_seguro, sems_seguro = [], [], []
    xs_riesgo,  ys_riesgo,  sems_riesgo  = [], [], []

    dias_ratas_raw = {}

    for dia_idx, ruta in enumerate(archivos, start=1): #iniciamos el bucle principal usando enumerate para obtener simultaneamente el numero de dia comenzando en uno y la ruta absoluta de cada archivo excel que contiene los registros de latencia de esa sesion experimental
        try:
            archivo_excel = pd.ExcelFile(ruta)

            ratas_del_dia = {}
            rata_pos = 0

            for hoja in archivo_excel.sheet_names: #iteramos a traves del nombre de cada pestaña dentro del libro de excel actual partiendo de la premisa de que el software de recoleccion genera una pestaña individual con la totalidad de los ensayos para cada rata evaluada
                if hoja == 'Promedios Latencia': #evaluamos si el nombre de la pestaña coincide exactamente con la hoja de resumen automatico para brincarla inmediatamente ya que nuestro algoritmo calcula sus propias metricas de latencia agrupadas por tercio a partir de los datos crudos
                    continue
                df_tmp = archivo_excel.parse(hoja)
                if 'Latencia' not in df_tmp.columns or 'Estim Electrico' not in df_tmp.columns: #comprobamos mediante un operador logico que la pestaña actual contenga estrictamente las dos variables independientes necesarias para el analisis descartando asi cualquier hoja de metadatos o registros corruptos del equipo
                    continue

                rata_pos += 1
                seg = df_tmp.loc[df_tmp['Estim Electrico'] == 0,
                                'Latencia'].reset_index(drop=True)
                rie = df_tmp.loc[df_tmp['Estim Electrico'] == 1,
                                'Latencia'].reset_index(drop=True)

                ratas_del_dia[rata_pos] = {
                    'seguro': _tercio_split(seg),
                    'riesgo': _tercio_split(rie),
                }

            dias_ratas_raw[dia_idx] = ratas_del_dia

            if not ratas_del_dia: #verificamos si la estructura de datos que deberia contener la informacion de todas las ratas del dia quedo completamente vacia despues de la extraccion lo que indicaria que ninguna pestaña fue util y asi poder notificar el error sin detener el script principal
                log_fn(f"Día {dia_idx}: no se encontraron datos válidos.")
                continue

            for t_idx in range(3): #establecemos un ciclo cerrado de tres iteraciones exactas que representan la segmentacion temporal de la sesion en bloques de ensayos iniciales medios y finales para poder observar la evolucion de la latencia y el aprendizaje en las ratas
                tercio_num = t_idx + 1

                medias_seg = []
                medias_rie = []

                for rata_data in ratas_del_dia.values(): #recorremos los diccionarios internos extraidos previamente que contienen las series de latencias ya separadas por condicion de riesgo y seguridad para ir agrupando de manera individual los valores correspondientes al bloque temporal actual
                    grupo_seg = rata_data['seguro'][t_idx]
                    grupo_rie = rata_data['riesgo'][t_idx]
                    if len(grupo_seg) > 0: #verificamos mediante la longitud del arreglo que la rata haya ejecutado al menos un intento seguro en este tercio especifico de la sesion para poder extraer su media aritmetica y evitar un error fatal al intentar promediar un arreglo vacio
                        medias_seg.append(float(grupo_seg.mean()))
                    if len(grupo_rie) > 0: #nos cercioramos de que el sujeto experimental haya realizado cruces bajo la condicion de choque electrico durante este bloque temporal antes de incluir su latencia promedio en la lista de datos a analizar para el total de la poblacion
                        medias_rie.append(float(grupo_rie.mean()))

                if medias_seg: #comprobamos que la lista acumulativa de medias seguras de todas las ratas no este vacia lo que nos permite calcular de forma iterativa el promedio global del grupo experimental y su varianza estadistica a traves del error estandar de la media
                    prom_seg = round(float(pd.Series(medias_seg).mean()), 2)
                    sem_seg  = round(_sem(medias_seg), 3)
                else:
                    prom_seg, sem_seg = None, None

                if medias_rie: #evaluamos si logramos recopilar al menos una media de riesgo entre todos los sujetos para proceder con el calculo del estimador estadistico poblacional de la latencia ante el estimulo aversivo para este tercio especifico de la prueba
                    prom_rie = round(float(pd.Series(medias_rie).mean()), 2)
                    sem_rie  = round(_sem(medias_rie), 3)
                else:
                    prom_rie, sem_rie = None, None

                x = _x_tercio(dia_idx, tercio_num)

                if prom_seg is not None and prom_seg != 0: #aplicamos una doble validacion para confirmar que el calculo matematico arrojo un valor numerico real y diferente de cero asegurando que solo enviaremos coordenadas fidedignas al motor de graficado de la libreria matplotlib
                    xs_seguro.append(x)
                    ys_seguro.append(prom_seg)
                    sems_seguro.append(sem_seg)

                if prom_rie is not None and prom_rie != 0: #validamos estrictamente que el promedio de riesgo tenga un valor cuantitativo util para el analisis visual descartando nulos o ceros absolutos que podrian distorsionar gravemente la percepcion del comportamiento de las ratas en el canvas final
                    xs_riesgo.append(x)
                    ys_riesgo.append(prom_rie)
                    sems_riesgo.append(sem_rie)

                filas_resumen.append({
                    'Día':         dia_idx,
                    'Bloque':      tercio_num,
                    'Prom Seguro': prom_seg,
                    'SEM Seguro':  sem_seg,
                    'Prom Riesgo': prom_rie,
                    'SEM Riesgo':  sem_rie,
                })

            total_ensayos = sum(
                sum(len(rd[cond][t]) for cond in ('seguro', 'riesgo') for t in range(3))
                for rd in ratas_del_dia.values()
            )
            log_fn(
                f"Día {dia_idx}: {os.path.basename(ruta)} — "
                f"{total_ensayos} ensayos procesados ({rata_pos} rata(s))."
            )

        except Exception as e:
            log_fn(f"Error procesando {os.path.basename(ruta)}: {str(e)}")

    if not xs_seguro and not xs_riesgo: #hacemos una ultima comprobacion de seguridad al terminar de procesar todos los archivos para confirmar si las listas de coordenadas finales estan vacias y asi cancelar la creacion del canvas grafico regresando unicamente valores nulos al entorno
        log_fn("No hay datos válidos para graficar.")
        return None, None, None

    fig, ax = plt.subplots(figsize=(11, 6))
    err_kw = dict(capsize=5, capthick=1.5, elinewidth=1.5)

    if xs_seguro: #verificamos si el arreglo de coordenadas del eje x para la condicion segura contiene elementos antes de ordenar a la libreria grafica que dibuje la linea verde de dispersion con sus respectivos marcadores circulares y barras de error
        ax.errorbar(xs_seguro, ys_seguro, yerr=sems_seguro,
                    fmt='o', color='green', markersize=8, alpha=0.85,
                    label='Seguro', **err_kw)
    if xs_riesgo: #verificamos si existen datos mapeables en la lista de riesgo para indicarle a la libreria grafica que proceda a renderizar los puntos negros cuadrados y trazar la linea de tendencia que representa el conflicto aversivo durante la prueba
        ax.errorbar(xs_riesgo, ys_riesgo, yerr=sems_riesgo,
                    fmt='s', color='black', markersize=8, alpha=0.85,
                    label='Riesgo', **err_kw)

    num_dias = len(archivos)
    ax.set_xticks(range(1, num_dias + 1))
    ax.set_xticklabels([str(d) for d in range(1, num_dias + 1)], fontsize=10)

    for d in range(1, num_dias): #iteramos a traves del numero total de dias procesados menos uno para calcular y dibujar dinamicamente las coordenadas de las lineas punteadas verticales que actuaran como separadores visuales del progreso diario a lo largo de la sesion experimental
        ax.axvline(x=d + 0.5, color='gray', linestyle=':', linewidth=0.8, alpha=0.5)

    any_seg = len(xs_seguro) > 0
    any_rie = len(xs_riesgo) > 0
    if any_seg and any_rie: #evaluamos mediante banderas booleanas de longitud si el procesamiento logro extraer exitosamente tanto datos de ensayos con choque como sin choque para asignar dinamicamente el titulo general de discriminacion de conflictos entre ambos estimulos
        titulo = 'Discrimination (Conflict vs No-conflict)'
    elif not any_seg: #si la condicion anterior fallo entonces verificamos si la ausencia de datos corresponde especificamente a los ensayos seguros lo que cambiaria el contexto del grafico a un experimento enfocado unicamente en el cruce de amenazas con choque electrico
        titulo = 'Threat crossings (Noise/Shock + Light/Food)'
    else: #si caemos en esta ultima condicion logica podemos deducir matematicamente que solo existen datos de cruces seguros sin amenaza por lo que el titulo del grafico se ajustara para reflejar una tarea puramente guiada por la recompensa natural
        titulo = 'Reward crossings (Lights + Cross for food)'

    ax.set_title(titulo, fontsize=14, fontweight='bold')
    ax.set_xlabel('Days (block of trial)', fontsize=12)
    ax.set_ylabel('Latencies (s)', fontsize=12)
    ax.legend(fontsize=11)
    ax.yaxis.set_minor_locator(ticker.AutoMinorLocator())
    ax.grid(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    fig.tight_layout()

    df_resumen = pd.DataFrame(filas_resumen)
    log_fn("Gráfica generada correctamente.")
    return fig, df_resumen, dias_ratas_raw

#funcion que escribe la hoja con los datos crudos de cada rata separados por bloque

def _escribir_hoja_raw(ws, dias_ratas_raw, num_dias):
    """
    lee los archivos excel de cada dia, calcula promedios por bloque
    y genera la grafica con barras de error
    Args: archivos -> lista de rutas a los archivos .xlsx (uno por dia)
        log_fn   -> funcion para mostrar mensajes (por defecto print)
    Returns: (fig, df_resumen, dias_ratas_raw) si hay datos, o (None, None, None) si no
    """
    #nombres de los bloques para que se lean bien en el excel
    nombres_tercio = ['Primer bloque', 'Segundo bloque', 'Tercer bloque']

    #se arma la fila de encabezados con una columna de seguro y una de riesgo por dia
    header = ['']
    for d in range(1, num_dias + 1): #iteramos a traves de la secuencia numerica de los dias totales del experimento para generar dinamicamente las cabeceras de columnas emparejando la condicion segura y la de riesgo para cada jornada
        header.append(f'Seguro_d{d}')
        header.append(f'Riesgo_d{d}')
    for col_idx, h in enumerate(header, start=1): #recorremos la lista de cabeceras recien construida utilizando enumerate para obtener simultaneamente el indice de columna exacto en excel y escribir el texto aplicandole un formato de negritas para diferenciarlo visualmente
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(bold=True)

    #se busca cuantas ratas hay en total para saber cuantas filas se necesitan
    max_ratas = 0
    for ratas in dias_ratas_raw.values(): #iteramos sobre los diccionarios internos de cada dia procesado para localizar cual es el numero identificador maximo de rata registrado y asi dimensionar correctamente cuantas secciones verticales necesitamos construir en la hoja de calculo
        if ratas: #evaluamos si el diccionario de ratas del dia actual contiene elementos reales para evitar fallos al intentar calcular el valor maximo sobre una estructura de datos completamente vacia
            max_ratas = max(max_ratas, max(ratas.keys()))

    current_row = 2

    for rata_idx in range(1, max_ratas + 1): #iniciamos el ciclo principal de escritura iterando sobre el numero de ratas totales para generar los bloques de datos de cada sujeto experimental asegurandonos de incluir a todos incluso si les faltan dias
        #se pone el nombre de la rata en negrita
        cell = ws.cell(row=current_row, column=1, value=f'r{rata_idx}')
        cell.font = Font(bold=True)
        current_row += 1

        for t_idx in range(3): #establecemos un ciclo anidado de tres repeticiones para organizar los datos de la rata actual dividiendolos cronologicamente en primer segundo y tercer bloque de la sesion experimental para facilitar el analisis temporal
            segs_por_dia = []
            ries_por_dia = []

            #se juntan los valores de esta rata para cada dia
            for dia_idx in range(1, num_dias + 1): #recorremos cada uno de los dias experimentales buscando en el diccionario maestro los registros especificos de esta rata durante este bloque temporal particular para poder agruparlos horizontalmente en el excel
                ratas     = dias_ratas_raw.get(dia_idx, {})
                rata_data = ratas.get(rata_idx)
                if rata_data: #comprobamos que la rata actual efectivamente tenga un registro valido de ensayos durante este dia especifico extrayendo sus listas de valores y convirtiendolas a formato nativo de python mediante el metodo tolist para manipularlas comodamente
                    segs_por_dia.append(rata_data['seguro'][t_idx].tolist())
                    ries_por_dia.append(rata_data['riesgo'][t_idx].tolist())
                else: #si la condicion anterior arroja falso porque la rata no tiene datos registrados en este dia especifico entonces insertamos listas vacias para mantener sincronizada la alineacion de columnas
                    segs_por_dia.append([])
                    ries_por_dia.append([])

            max_vals = max(
                max((len(s) for s in segs_por_dia), default=0),
                max((len(r) for r in ries_por_dia), default=0),
                1,
            )

            for val_row in range(max_vals): #iteramos sobre el numero maximo de ensayos detectados para este bloque generando suficientes filas en el excel para acomodar la lista mas larga de latencias sin truncar datos importantes
                row = current_row + val_row
                #en la primera fila de cada bloque se pone el nombre del bloque
                if val_row == 0: #verificamos si estamos escribiendo la primera fila de este conjunto de datos para aprovechar y colocar el titulo descriptivo del bloque en la primera columna ayudando a la legibilidad del archivo final
                    ws.cell(row=row, column=1, value=f'  {nombres_tercio[t_idx]}')

                for dia_idx in range(1, num_dias + 1): #volvemos a iterar sobre los dias experimentales para calcular la posicion exacta de las columnas de excel y proceder a volcar los valores numericos individuales de latencia fila por fila
                    seg_col  = 2 + (dia_idx - 1) * 2
                    rie_col  = seg_col + 1
                    seg_vals = segs_por_dia[dia_idx - 1]
                    rie_vals = ries_por_dia[dia_idx - 1]

                    if val_row < len(seg_vals): #validamos que el indice de fila actual siga siendo menor que la longitud total de la lista de ensayos seguros de este dia para no intentar leer un elemento inexistente y asi poder escribir la celda sin provocar un desbordamiento de indice
                        ws.cell(row=row, column=seg_col,
                                value=round(float(seg_vals[val_row]), 4))
                    if val_row < len(rie_vals): #comprobamos cuidadosamente que aun existan valores pendientes en la lista de latencias de riesgo antes de extraer el numero redondearlo a cuatro decimales e insertarlo en su celda respectiva de excel sin interrumpir el proceso
                        ws.cell(row=row, column=rie_col,
                                value=round(float(rie_vals[val_row]), 4))

            current_row += max_vals
            current_row += 1  #fila en blanco para separar bloques

        current_row += 1  #fila extra para separar ratas

    #se ajustan los anchos de las columnas para que se lea bien
    ws.column_dimensions['A'].width = 20
    for d in range(1, num_dias + 1): #iteramos por ultima vez sobre la cantidad de dias del experimento para calcular dinamicamente la letra correspondiente a cada par de columnas y aplicarles un ancho predefinido mejorando asi la presentacion final del documento
        ws.column_dimensions[get_column_letter(2 + (d - 1) * 2)].width = 14
        ws.column_dimensions[get_column_letter(3 + (d - 1) * 2)].width = 14

def _escribir_hoja_promedios(ws, dias_ratas_raw, num_dias):
    """
    llena la hoja 'Promedio Rata' con el promedio de cada rata por dia
    y al final pone el promedio general de todos los dias
    Args: ws -> hoja de excel donde se va a escribir
            dias_ratas_raw -> diccionario con los datos crudos por dia y rata
            num_dias       -> cuantos dias hay en total
    #Returns: no regresa nada, escribe directamente en la hoja
    """

    #encabezado: una columna por condicion y dia, mas dos de promedio general
    header = ['Rata']
    for d in range(1, num_dias + 1): #iteramos a traves de la cantidad total de dias del experimento para construir de manera dinamica la lista de encabezados agregando un par de columnas por cada dia para separar los datos seguros de los de riesgo
        header.append(f'Seguro Día {d}')
        header.append(f'Riesgo Día {d}')
    col_prom_seg = len(header) + 1
    col_prom_rie = col_prom_seg + 1
    header.append('Prom Seguro')
    header.append('Prom Riesgo')

    for col_idx, h in enumerate(header, start=1): #recorremos el arreglo de encabezados recien creado utilizando la funcion enumerate para obtener el indice exacto de columna donde escribiremos cada titulo aplicandole inmediatamente un formato de texto en negrita
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(bold=True)

    #cuantas ratas hay en total
    max_ratas = 0
    for ratas in dias_ratas_raw.values(): #exploramos los registros diarios extraidos del diccionario maestro buscando iterativamente el numero maximo de rata que haya participado en el experimento para asegurar que asignamos suficientes filas en el documento de salida
        if ratas: #evaluamos que el registro diario no este completamente vacio antes de intentar extraer las llaves de identificacion para prevenir errores de ejecucion al calcular el valor maximo sobre secuencias inexistentes
            max_ratas = max(max_ratas, max(ratas.keys()))

    for rata_idx in range(1, max_ratas + 1): #iniciamos el ciclo principal de escritura iterando secuencialmente sobre el numero total de ratas identificadas para procesar de forma individual el historial completo de cada sujeto experimental a lo largo de todos los dias
        row = rata_idx + 1
        cell = ws.cell(row=row, column=1, value=f'r{rata_idx}')
        cell.font = Font(bold=True)

        promedios_seg_dias = []
        promedios_rie_dias = []

        for dia_idx in range(1, num_dias + 1): #recorremos uno a uno los dias de prueba buscando en la base de datos la informacion correspondiente a la rata actual para asi poder calcular sus promedios diarios antes de insertarlos en las celdas respectivas
            ratas     = dias_ratas_raw.get(dia_idx, {})
            rata_data = ratas.get(rata_idx)

            seg_col = 2 + (dia_idx - 1) * 2
            rie_col = seg_col + 1

            if rata_data: #verificamos cuidadosamente si existen datos experimentales de esta rata especifica durante la sesion actual para poder proceder a unificar sus bloques temporales y calcular una media diaria consolidada
                #se juntan los tres bloques para tener todos los ensayos del dia
                all_seg = pd.concat(rata_data['seguro']).dropna()
                all_rie = pd.concat(rata_data['riesgo']).dropna()

                if len(all_seg) > 0: #comprobamos mediante la longitud del dataframe concatenado que el sujeto efectivamente haya realizado intentos seguros en este dia evitando asi enviar celdas nulas o provocar errores matematicos al intentar promediar
                    prom_dia_seg = round(float(all_seg.mean()), 4)
                    ws.cell(row=row, column=seg_col, value=prom_dia_seg)
                    promedios_seg_dias.append(prom_dia_seg)

                if len(all_rie) > 0: #nos aseguramos de que existan ensayos bajo amenaza validos despues de la limpieza de nulos antes de calcular la media aritmetica redondearla a cuatro decimales y sumarla al historial global del sujeto
                    prom_dia_rie = round(float(all_rie.mean()), 4)
                    ws.cell(row=row, column=rie_col, value=prom_dia_rie)
                    promedios_rie_dias.append(prom_dia_rie)

        #promedio general de todos los dias para esta rata, va en negrita
        if promedios_seg_dias: #evaluamos si despues de analizar todos los dias la lista acumulativa de medias seguras contiene al menos un valor valido para finalmente calcular el promedio global del experimento y resaltarlo visualmente en la tabla
            cell = ws.cell(row=row, column=col_prom_seg,
                        value=round(float(pd.Series(promedios_seg_dias).mean()), 4))
            cell.font = Font(bold=True)
        if promedios_rie_dias: #verificamos que el historial de promedios de riesgo diarios no este vacio confirmando que el sujeto participo en estas condiciones para calcular su media general historica e insertarla en la ultima columna del reporte
            cell = ws.cell(row=row, column=col_prom_rie,
                        value=round(float(pd.Series(promedios_rie_dias).mean()), 4))
            cell.font = Font(bold=True)

    #anchos de columna para que no quede cortado el texto
    ws.column_dimensions['A'].width = 10
    for d in range(1, num_dias + 1): #realizamos una ultima pasada sobre el rango de dias para calcular matematicamente la designacion de las letras de cada columna en excel ajustando asi sus dimensiones horizontales y garantizando la correcta lectura del documento
        ws.column_dimensions[get_column_letter(2 + (d - 1) * 2)].width = 15
        ws.column_dimensions[get_column_letter(3 + (d - 1) * 2)].width = 15
    ws.column_dimensions[get_column_letter(col_prom_seg)].width = 14
    ws.column_dimensions[get_column_letter(col_prom_rie)].width = 14

def guardar_excel_resumen(df_resumen, ruta_xlsx, dias_ratas_raw=None, num_dias=0, log_fn=print):
    """
    guarda el excel final con tres hojas: resumen, desglose y promedios por rata
    Args: df_resumen    -> DataFrame con los promedios y SEM por dia y bloque
            ruta_xlsx     -> ruta donde se va a guardar el archivo
            dias_ratas_raw -> diccionario con los datos crudos (opcional)
            num_dias       -> numero total de dias procesados
            log_fn         -> funcion para mostrar mensajes (por defecto print)
    #Returns: True si se guardo bien, False si hubo algun error
    """
    try:
        with pd.ExcelWriter(ruta_xlsx, engine='openpyxl') as writer:
            #primera hoja: resumen con promedios y errores por dia y bloque
            df_resumen.to_excel(writer, sheet_name='Resumen Bloques', index=False)
            ws_res = writer.sheets['Resumen Bloques']
            for cell in ws_res[1]:
                cell.font = Font(bold=True)
            #se ajusta el ancho de cada columna segun el contenido
            for col_idx, col in enumerate(df_resumen.columns, start=1):
                col_letter = get_column_letter(col_idx)
                max_len = max(
                    len(str(col)),
                    *[len(str(c.value)) for c in ws_res[col_letter]
                    if c.value is not None]
                )
                ws_res.column_dimensions[col_letter].width = max_len + 3

            #segunda hoja: los datos crudos de cada rata separados por bloque
            if dias_ratas_raw and num_dias > 0:
                ws_raw = writer.book.create_sheet(title='Desglose Bloques')
                _escribir_hoja_raw(ws_raw, dias_ratas_raw, num_dias)

            #tercera hoja: el promedio de cada rata por dia
            if dias_ratas_raw and num_dias > 0:
                ws_prom = writer.book.create_sheet(title='Promedio Rata')
                _escribir_hoja_promedios(ws_prom, dias_ratas_raw, num_dias)

        log_fn(f"Excel de resumen guardado en: {ruta_xlsx}")
        return True
    except Exception as e:
        log_fn(f"Error guardando Excel: {str(e)}")
        return False
