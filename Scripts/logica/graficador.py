import os
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def _sem(serie):
    n = serie.count()
    return (serie.std() / math.sqrt(n)) if n > 1 else 0


def _tercio_split(serie):
    n = len(serie)
    if n == 0:
        return [
            pd.Series(dtype=float),
            pd.Series(dtype=float),
            pd.Series(dtype=float),
        ]
    g1 = math.ceil(n / 3)
    rem = n - g1
    g2 = math.ceil(rem / 2) if rem > 0 else 0
    g3 = rem - g2
    return [
        serie.iloc[:g1].reset_index(drop=True),
        serie.iloc[g1: g1 + g2].reset_index(drop=True),
        serie.iloc[g1 + g2:].reset_index(drop=True),
    ]


def _tercio_stats(serie):
    """Devuelve [{promedio, sem}, ...] para cada uno de los 3 tercios."""
    grupos = _tercio_split(serie)
    result = []
    for grupo in grupos:
        result.append({
            'promedio': round(float(grupo.mean()), 2) if len(grupo) > 0 else None,
            'sem':      round(_sem(grupo), 3)         if len(grupo) > 0 else None,
        })
    return result


def _x_tercio(dia, tercio):
    """Coordenada X: cada día ocupa 1 unidad, tercios en -0.25, 0, +0.25."""
    offsets = {1: -0.25, 2: 0.0, 3: 0.25}
    return dia + offsets[tercio]


# ── Función principal ─────────────────────────────────────────────────────────

def generar_grafica(archivos, log_fn=print):
    """
    Lee archivos .xlsx (uno por día).  Por cada día lee TODAS las hojas de rata,
    las combina para el scatter plot (promedios por tercio) y guarda los valores
    crudos por rata para la segunda hoja del Excel de resumen.

    Returns:
        (fig, df_resumen, dias_ratas_raw)   si hay datos válidos
        (None, None, None)                  si no hay datos
    """
    filas_resumen = []
    xs_seguro, ys_seguro, sems_seguro = [], [], []
    xs_riesgo,  ys_riesgo,  sems_riesgo  = [], [], []

    # dias_ratas_raw[dia_idx][rata_pos] = {
    #     'seguro': [Serie_t1, Serie_t2, Serie_t3],
    #     'riesgo': [Serie_t1, Serie_t2, Serie_t3],
    # }
    dias_ratas_raw = {}

    for dia_idx, ruta in enumerate(archivos, start=1):
        try:
            archivo_excel = pd.ExcelFile(ruta)

            all_seguro   = []
            all_riesgo   = []
            ratas_del_dia = {}
            rata_pos      = 0

            for hoja in archivo_excel.sheet_names:
                # Saltar hoja de promedios
                if hoja == 'Promedios Latencia':
                    continue
                df_tmp = archivo_excel.parse(hoja)
                if 'Latencia' not in df_tmp.columns or 'Estim Electrico' not in df_tmp.columns:
                    continue

                rata_pos += 1
                seg = df_tmp.loc[df_tmp['Estim Electrico'] == 0,
                                'Latencia'].reset_index(drop=True)
                rie = df_tmp.loc[df_tmp['Estim Electrico'] == 1,
                                'Latencia'].reset_index(drop=True)

                # Acumular para el pool del scatter
                all_seguro.extend(seg.tolist())
                all_riesgo.extend(rie.tolist())

                # Guardar tercios raw por rata
                ratas_del_dia[rata_pos] = {
                    'seguro': _tercio_split(seg),
                    'riesgo': _tercio_split(rie),
                }

            dias_ratas_raw[dia_idx] = ratas_del_dia

            if not all_seguro and not all_riesgo:
                log_fn(f"Día {dia_idx}: no se encontraron datos válidos.")
                continue

            seguro_serie = pd.Series(all_seguro).reset_index(drop=True)
            riesgo_serie = pd.Series(all_riesgo).reset_index(drop=True)

            for t_idx, (seg_st, rie_st) in enumerate(
                zip(_tercio_stats(seguro_serie), _tercio_stats(riesgo_serie)), start=1
            ):
                x = _x_tercio(dia_idx, t_idx)

                if seg_st['promedio'] is not None and seg_st['promedio'] != 0:
                    xs_seguro.append(x)
                    ys_seguro.append(seg_st['promedio'])
                    sems_seguro.append(seg_st['sem'])

                if rie_st['promedio'] is not None and rie_st['promedio'] != 0:
                    xs_riesgo.append(x)
                    ys_riesgo.append(rie_st['promedio'])
                    sems_riesgo.append(rie_st['sem'])

                filas_resumen.append({
                    'Día':         dia_idx,
                    'Tercio':      t_idx,
                    'Prom Seguro': seg_st['promedio'],
                    'SEM Seguro':  seg_st['sem'],
                    'Prom Riesgo': rie_st['promedio'],
                    'SEM Riesgo':  rie_st['sem'],
                })

            total = len(all_seguro) + len(all_riesgo)
            log_fn(
                f"Día {dia_idx}: {os.path.basename(ruta)} — "
                f"{total} ensayos procesados ({rata_pos} rata(s))."
            )

        except Exception as e:
            log_fn(f"Error procesando {os.path.basename(ruta)}: {str(e)}")

    if not xs_seguro and not xs_riesgo:
        log_fn("No hay datos válidos para graficar.")
        return None, None, None

    # ── Figura ────────────────────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(11, 6))
    err_kw = dict(capsize=5, capthick=1.5, elinewidth=1.5)

    if xs_seguro:
        ax.errorbar(xs_seguro, ys_seguro, yerr=sems_seguro,
                    fmt='o', color='green', markersize=8, alpha=0.85,
                    label='Seguro', **err_kw)
    if xs_riesgo:
        ax.errorbar(xs_riesgo, ys_riesgo, yerr=sems_riesgo,
                    fmt='s', color='black', markersize=8, alpha=0.85,
                    label='Riesgo', **err_kw)

    num_dias = len(archivos)
    ax.set_xticks(range(1, num_dias + 1))
    ax.set_xticklabels([str(d) for d in range(1, num_dias + 1)], fontsize=10)

    for d in range(1, num_dias):
        ax.axvline(x=d + 0.5, color='gray', linestyle=':', linewidth=0.8, alpha=0.5)
    if seg_st['promedio'] and rie_st['promedio'] != 0:   
        ax.set_title('Discrimination (Conflict vs No-conflict)', fontsize=14, fontweight='bold')
    elif seg_st['promedio'] == 0:
        ax.set_title('Threat crossings (Noise/Shock + Light/Food)', fontsize=14, fontweight='bold')
    else:
        ax.set_title('Reward crossings (Lights + Cross for food)', fontsize=14, fontweight='bold')
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


# ── Segunda hoja: datos crudos por rata ───────────────────────────────────────

def _escribir_hoja_raw(ws, dias_ratas_raw, num_dias):
    """
    Escribe la hoja 'Datos por Rata' con el formato:

        etiqueta | Seguro_d1 | Riesgo_d1 | Seguro_d2 | Riesgo_d2 | ...

    Agrupado por rata (r1, r2, …) y dentro de cada rata por tercio
    (Primer tercio, Segundo Tercio, Tercer tercio).
    Cada bloque tercio muestra los valores crudos verticalmente.
    Una fila en blanco separa tercios; una fila extra separa ratas.
    """
    nombres_tercio = ['Primer tercio', 'Segundo Tercio', 'Tercer tercio']

    # ── Encabezado ─────────────────────────────────────────────────────────
    header = ['']
    for d in range(1, num_dias + 1):
        header.append(f'Seguro_d{d}')
        header.append(f'Riesgo_d{d}')

    for col_idx, h in enumerate(header, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(bold=True)

    # ── Número máximo de ratas vistas en cualquier día ─────────────────────
    max_ratas = 0
    for ratas in dias_ratas_raw.values():
        if ratas:
            max_ratas = max(max_ratas, max(ratas.keys()))

    current_row = 2

    for rata_idx in range(1, max_ratas + 1):

        # Etiqueta de rata (negrita)
        cell = ws.cell(row=current_row, column=1, value=f'r{rata_idx}')
        cell.font = Font(bold=True)
        current_row += 1

        for t_idx in range(3):

            # Recopilar los valores de esta rata × tercio para cada día
            segs_por_dia = []
            ries_por_dia = []

            for dia_idx in range(1, num_dias + 1):
                ratas      = dias_ratas_raw.get(dia_idx, {})
                rata_data  = ratas.get(rata_idx)
                if rata_data:
                    segs_por_dia.append(rata_data['seguro'][t_idx].tolist())
                    ries_por_dia.append(rata_data['riesgo'][t_idx].tolist())
                else:
                    segs_por_dia.append([])
                    ries_por_dia.append([])

            max_vals = max(
                max((len(s) for s in segs_por_dia), default=0),
                max((len(r) for r in ries_por_dia), default=0),
                1,          # al menos 1 fila para la etiqueta del tercio
            )

            for val_row in range(max_vals):
                row = current_row + val_row
                # Etiqueta solo en la primera fila del bloque
                if val_row == 0:
                    ws.cell(row=row, column=1, value=f'  {nombres_tercio[t_idx]}')

                for dia_idx in range(1, num_dias + 1):
                    seg_col = 2 + (dia_idx - 1) * 2
                    rie_col = seg_col + 1
                    seg_vals = segs_por_dia[dia_idx - 1]
                    rie_vals = ries_por_dia[dia_idx - 1]

                    if val_row < len(seg_vals):
                        ws.cell(row=row, column=seg_col,
                                value=round(float(seg_vals[val_row]), 4))
                    if val_row < len(rie_vals):
                        ws.cell(row=row, column=rie_col,
                                value=round(float(rie_vals[val_row]), 4))

            current_row += max_vals
            current_row += 1  # fila en blanco entre tercios (y después del último)

        # Fila extra en blanco entre ratas
        current_row += 1

    # ── Ajustar anchos de columna ──────────────────────────────────────────
    ws.column_dimensions['A'].width = 20
    for d in range(1, num_dias + 1):
        ws.column_dimensions[get_column_letter(2 + (d - 1) * 2)].width = 14
        ws.column_dimensions[get_column_letter(3 + (d - 1) * 2)].width = 14


# ── Guardado del Excel de resumen ─────────────────────────────────────────────

def guardar_excel_resumen(df_resumen, ruta_xlsx,
                        dias_ratas_raw=None, num_dias=0, log_fn=print):
    """
    Guarda dos hojas:
    1. 'Resumen Tercios'  — promedios y SEMs por día/tercio (igual que antes).
    2. 'Datos por Rata'   — valores crudos organizados por rata, día y tercio.
    """
    try:
        with pd.ExcelWriter(ruta_xlsx, engine='openpyxl') as writer:
            # ── Hoja 1: Resumen Tercios ────────────────────────────────────
            df_resumen.to_excel(writer, sheet_name='Resumen Tercios', index=False)
            ws_res = writer.sheets['Resumen Tercios']
            for cell in ws_res[1]:
                cell.font = Font(bold=True)
            for col_idx, col in enumerate(df_resumen.columns, start=1):
                col_letter = get_column_letter(col_idx)
                max_len = max(
                    len(str(col)),
                    *[len(str(c.value)) for c in ws_res[col_letter]
                    if c.value is not None]
                )
                ws_res.column_dimensions[col_letter].width = max_len + 3

            # ── Hoja 2: Datos por Rata ─────────────────────────────────────
            if dias_ratas_raw and num_dias > 0:
                wb = writer.book
                ws_raw = wb.create_sheet(title='Desglose tercios')
                _escribir_hoja_raw(ws_raw, dias_ratas_raw, num_dias)

        log_fn(f"Excel de resumen guardado en: {ruta_xlsx}")
        return True
    except Exception as e:
        log_fn(f"Error guardando Excel: {str(e)}")
        return False