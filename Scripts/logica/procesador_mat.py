import pandas as pd
import scipy.io
import os
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def leer_mat_a_df(ruta_archivo, log_fn=print):
    """Lee un archivo .mat y lo convierte en un DataFrame de pandas."""
    try:
        columnas = [
            'Ensayo', 'Lado', 'Estim Electrico', 'Latencia',
            'Tiempo Absoluto', 'Palancas Izq',
            'Palancas Der', 'Desplazamiento'
        ]
        mat_data = scipy.io.loadmat(ruta_archivo)
        matriz_datos = None
        for key, value in mat_data.items():
            if not key.startswith('__'):
                matriz_datos = value
                break

        if matriz_datos is None:
            log_fn(f"Error: No se encontraron datos válidos en {os.path.basename(ruta_archivo)}")
            return None

        if matriz_datos.shape[1] == len(columnas):
            df = pd.DataFrame(matriz_datos, columns=columnas)
            df['archivo_origen'] = os.path.basename(ruta_archivo)
            return df
        else:
            log_fn(
                f"Error en dimensiones: {os.path.basename(ruta_archivo)} "
                f"tiene {matriz_datos.shape[1]} columnas, "
                f"pero esperamos {len(columnas)}."
            )
            return None

    except Exception as e:
        log_fn(f"Error al procesar {os.path.basename(ruta_archivo)}: {str(e)}")
        return None


def eliminar_lados_consecutivos(df):
    """Elimina filas con el mismo 'Lado' consecutivo, dejando solo la primera de cada bloque."""
    if 'Lado' not in df.columns:
        return df
    cambios = df['Lado'].ne(df['Lado'].shift()).cumsum()
    return df.groupby(cambios).first().reset_index(drop=True)


def modificar_dataframe(df):
    """Limpia y transforma el DataFrame: elimina columnas, filtra y renumera ensayos."""
    df = df.drop(['Tiempo Absoluto', 'Palancas Izq', 'Palancas Der', 'archivo_origen'], axis=1)
    df = df[df['Desplazamiento'] > 1]
    df = eliminar_lados_consecutivos(df)
    df['Ensayo'] = range(1, len(df) + 1)
    return df


def calcular_promedios_latencia(dfs_dict, log_fn=print):
    """Calcula promedios de latencia con y sin estimulación eléctrica por archivo."""
    filas = []
    for nombre_archivo, df in dfs_dict.items():
        if 'Latencia' not in df.columns or 'Estim Electrico' not in df.columns:
            log_fn(f"Advertencia: {nombre_archivo} no tiene las columnas requeridas.")
            continue
        promedio_con = df.loc[df['Estim Electrico'] == 1, 'Latencia'].mean()
        promedio_sin = df.loc[df['Estim Electrico'] == 0, 'Latencia'].mean()
        filas.append({
            'Nombre de archivo': nombre_archivo,
            'Promedio Seguro': round(promedio_sin, 1) if pd.notna(promedio_sin) else 0,
            'Promedio Riesgo': round(promedio_con, 1) if pd.notna(promedio_con) else 0,
        })
    return pd.DataFrame(filas)


def guardar_excel(ruta_guardado, dfs_mat, log_fn=print):
    """Escribe todos los DataFrames en un archivo Excel con formato."""
    try:
        with pd.ExcelWriter(ruta_guardado, engine='openpyxl') as writer:
            df_promedios = calcular_promedios_latencia(dfs_mat, log_fn)
            promedio_riesgo = round(df_promedios['Promedio Riesgo'].mean(), 1)
            promedio_seguro = round(df_promedios['Promedio Seguro'].mean(), 1)
            indice_nueva_fila = len(df_promedios) + 2
            df_promedios.loc[indice_nueva_fila, 'Promedio Riesgo'] = promedio_riesgo
            df_promedios.loc[indice_nueva_fila, 'Promedio Seguro'] = promedio_seguro

            dfs_mat['Promedios Latencia'] = df_promedios

            for nombre_archivo, df_final in dfs_mat.items():
                nombre_hoja = nombre_archivo.replace('.mat', '')[:31]
                df_final.to_excel(writer, sheet_name=nombre_hoja, index=False)

                worksheet = writer.sheets[nombre_hoja]
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)

                for col_idx, column in enumerate(df_final.columns, start=1):
                    col_letter = get_column_letter(col_idx)
                    max_length = len(str(column))
                    for cell in worksheet[col_letter]:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except Exception:
                            pass
                    worksheet.column_dimensions[col_letter].width = max_length + 2

        log_fn(f"Archivo guardado en {ruta_guardado}")
        return True

    except Exception as e:
        log_fn(f"Error crítico al guardar: {str(e)}")
        return False
