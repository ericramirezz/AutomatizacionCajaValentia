import os
import pandas as pd
import matplotlib.pyplot as plt


def generar_grafica(archivos, log_fn=print):
    """
    Lee archivos .xlsx y genera una gráfica de dispersión comparando
    promedios de latencia (seguro vs riesgo) a lo largo de los días.

    Args:
        archivos: lista de rutas a archivos .xlsx
        log_fn:   función para registrar mensajes (por defecto print)

    Returns:
        ruta_png (str) si la gráfica se guardó, None si hubo error.
    """
    dias_seguros, valores_seguros = [], []
    dias_riesgo, valores_riesgo = [], []

    for dia_idx, ruta in enumerate(archivos, start=1):
        try:
            archivo_excel = pd.ExcelFile(ruta)
            ultima_hoja = archivo_excel.sheet_names[-1]
            df = archivo_excel.parse(ultima_hoja)

            columna_seguros = 'Promedio Seguro'
            columna_riesgo  = 'Promedio Riesgo'

            if columna_seguros in df.columns and columna_riesgo in df.columns:
                val_seguro = float(df[columna_seguros].iloc[-1])
                val_riesgo = float(df[columna_riesgo].iloc[-1])

                if val_seguro != 0 and val_riesgo == 0:
                    dias_seguros.append(dia_idx)
                    valores_seguros.append(val_seguro)
                elif val_riesgo != 0 and val_seguro == 0:
                    dias_riesgo.append(dia_idx)
                    valores_riesgo.append(val_riesgo)
                elif val_seguro != 0 and val_riesgo != 0:
                    dias_seguros.append(dia_idx)
                    valores_seguros.append(val_seguro)
                    dias_riesgo.append(dia_idx)
                    valores_riesgo.append(val_riesgo)
                else:
                    log_fn(f"Día {dia_idx}: ambos valores son 0, se omite.")
            else:
                log_fn(f"Error en {os.path.basename(ruta)}: faltan columnas.")

        except Exception as e:
            log_fn(f"Error procesando {os.path.basename(ruta)}: {str(e)}")

    if not dias_seguros and not dias_riesgo:
        log_fn("No hay datos válidos para graficar.")
        return None

    plt.figure(figsize=(9, 6))
    if dias_seguros:
        plt.scatter(dias_seguros, valores_seguros, color='blue',
                    label='Promedio Seguros', s=120, alpha=0.8)
    if dias_riesgo:
        plt.scatter(dias_riesgo, valores_riesgo, color='red',
                    label='Promedio Riesgo', s=120, alpha=0.8)

    plt.title('Evolución de Promedios de Latencia', fontsize=15, fontweight='bold')
    plt.xlabel('Día de Prueba', fontsize=12)
    plt.ylabel('Promedio de Latencia', fontsize=12)
    plt.xticks(range(1, len(archivos) + 1))
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.legend()

    ruta_png = 'grafica.png'
    plt.savefig(ruta_png, dpi=300, bbox_inches='tight')
    plt.close()

    log_fn(f"¡Éxito! Gráfica guardada como: {ruta_png}")
    return ruta_png
