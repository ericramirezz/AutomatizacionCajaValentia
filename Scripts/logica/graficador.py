import os
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sem(serie):
    n = serie.count()
    return (serie.std() / math.sqrt(n)) if n > 1 else 0


def _tercio_stats(serie):
    """Divide en 3 grupos por posición y devuelve [{promedio, sem}, ...]."""
    n = len(serie)
    size = n // 3
    grupos = []
    for i in range(3):
        inicio = i * size
        fin = inicio + size if i < 2 else n
        grupo = serie.iloc[inicio:fin]
        grupos.append({
            'promedio': round(grupo.mean(), 2) if len(grupo) > 0 else None,
            'sem':      round(_sem(grupo), 3)  if len(grupo) > 0 else None,
        })
    return grupos


def _x_tercio(dia, tercio):
    """Coordenada X: cada día ocupa 1 unidad, tercios en -0.25, 0, +0.25."""
    offsets = {1: -0.25, 2: 0.0, 3: 0.25}
    return dia + offsets[tercio]


# ── Función principal ─────────────────────────────────────────────────────────

def generar_grafica(archivos, log_fn=print):
    """
    Lee archivos .xlsx (uno por día), divide latencias en 3 tercios,
    calcula promedio y SEM, grafica con barras de error.

    Returns:
        (fig, df_resumen)  si hay datos válidos
        (None, None)       si no hay datos
    """
    filas_resumen = []
    xs_seguro, ys_seguro, sems_seguro = [], [], []
    xs_riesgo,  ys_riesgo,  sems_riesgo  = [], [], []

    for dia_idx, ruta in enumerate(archivos, start=1):
        try:
            archivo_excel = pd.ExcelFile(ruta)
            df_dia = None
            for hoja in archivo_excel.sheet_names:
                df_tmp = archivo_excel.parse(hoja)
                if 'Latencia' in df_tmp.columns and 'Estim Electrico' in df_tmp.columns:
                    df_dia = df_tmp
                    break

            if df_dia is None:
                log_fn(f"Día {dia_idx}: no se encontró hoja con datos de ensayos.")
                continue

            seguro_serie = df_dia.loc[df_dia['Estim Electrico'] == 0, 'Latencia'].reset_index(drop=True)
            riesgo_serie = df_dia.loc[df_dia['Estim Electrico'] == 1, 'Latencia'].reset_index(drop=True)

            for t_idx, (seg, rie) in enumerate(
                zip(_tercio_stats(seguro_serie), _tercio_stats(riesgo_serie)), start=1
            ):
                x = _x_tercio(dia_idx, t_idx)

                if seg['promedio'] is not None and seg['promedio'] != 0:
                    xs_seguro.append(x)
                    ys_seguro.append(seg['promedio'])
                    sems_seguro.append(seg['sem'])

                if rie['promedio'] is not None and rie['promedio'] != 0:
                    xs_riesgo.append(x)
                    ys_riesgo.append(rie['promedio'])
                    sems_riesgo.append(rie['sem'])

                filas_resumen.append({
                    'Día':         dia_idx,
                    'Tercio':      t_idx,
                    'Prom Seguro': seg['promedio'],
                    'SEM Seguro':  seg['sem'],
                    'Prom Riesgo': rie['promedio'],
                    'SEM Riesgo':  rie['sem'],
                })

            log_fn(f"Día {dia_idx}: {os.path.basename(ruta)} — {len(df_dia)} ensayos procesados.")

        except Exception as e:
            log_fn(f"Error procesando {os.path.basename(ruta)}: {str(e)}")

    if not xs_seguro and not xs_riesgo:
        log_fn("No hay datos válidos para graficar.")
        return None, None

    # ── Figura ────────────────────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(11, 6))
    err_kw = dict(capsize=5, capthick=1.5, elinewidth=1.5)

    if xs_seguro:
        ax.errorbar(xs_seguro, ys_seguro, yerr=sems_seguro,
                    fmt='o', color='blue', markersize=8, alpha=0.85,
                    label='Seguro', **err_kw)
    if xs_riesgo:
        ax.errorbar(xs_riesgo, ys_riesgo, yerr=sems_riesgo,
                    fmt='s', color='red', markersize=8, alpha=0.85,
                    label='Riesgo', **err_kw)

    num_dias = len(archivos)
    ax.set_xticks(range(1, num_dias + 1))
    ax.set_xticklabels([f'{d}' for d in range(1, num_dias + 1)], fontsize=10)

    for d in range(1, num_dias):
        ax.axvline(x=d + 0.5, color='gray', linestyle=':', linewidth=0.8, alpha=0.5)

    ax.set_title('Evolución de Promedios de Latencia por Tercio', fontsize=14, fontweight='bold')
    ax.set_xlabel('Día de Prueba', fontsize=12)
    ax.set_ylabel('Latencia (ms)', fontsize=12)
    ax.legend(fontsize=11)
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    ax.yaxis.set_minor_locator(ticker.AutoMinorLocator())
    fig.tight_layout()

    df_resumen = pd.DataFrame(filas_resumen)
    log_fn("Gráfica generada correctamente.")
    return fig, df_resumen


def guardar_excel_resumen(df_resumen, ruta_xlsx, log_fn=print):
    """Guarda el DataFrame de resumen (promedios y SEMs por tercio) en Excel."""
    try:
        with pd.ExcelWriter(ruta_xlsx, engine='openpyxl') as writer:
            df_resumen.to_excel(writer, sheet_name='Resumen Tercios', index=False)
            ws = writer.sheets['Resumen Tercios']
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for col_idx, col in enumerate(df_resumen.columns, start=1):
                col_letter = get_column_letter(col_idx)
                max_len = max(
                    len(str(col)),
                    *[len(str(c.value)) for c in ws[col_letter] if c.value is not None]
                )
                ws.column_dimensions[col_letter].width = max_len + 3
        log_fn(f"Excel de resumen guardado en: {ruta_xlsx}")
        return True
    except Exception as e:
        log_fn(f"Error guardando Excel: {str(e)}")
        return False