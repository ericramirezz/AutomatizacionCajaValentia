import os
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ── Helpers ───────────────────────────────────────────────────────────────────

def _sem(valores):
    """SEM de una lista de valores (desviación estándar / √n)."""
    n = len(valores)
    if n < 2:
        return 0.0
    s = pd.Series(valores)
    return float(s.std() / math.sqrt(n))


def _tercio_split(serie):
    """
    Divide una Serie en 3 grupos dando prioridad de tamaño a los primeros.
    Ejemplo: n=29 → 10, 10, 9  |  n=10 → 4, 3, 3
    """
    n = len(serie)
    if n == 0:
        return [pd.Series(dtype=float)] * 3
    g1 = math.ceil(n / 3)
    rem = n - g1
    g2 = math.ceil(rem / 2) if rem > 0 else 0
    return [
        serie.iloc[:g1].reset_index(drop=True),
        serie.iloc[g1: g1 + g2].reset_index(drop=True),
        serie.iloc[g1 + g2:].reset_index(drop=True),
    ]


def _x_tercio(dia, tercio):
    """Coordenada X: cada día ocupa 1 unidad, tercios en -0.25, 0, +0.25."""
    offsets = {1: -0.25, 2: 0.0, 3: 0.25}
    return dia + offsets[tercio]


# ── Función principal ─────────────────────────────────────────────────────────

def generar_grafica(archivos, log_fn=print):
    """
    Lee archivos .xlsx (uno por día). Por cada día:
      1. Por cada rata: divide sus ensayos en 3 tercios y calcula la media de cada tercio.
      2. El punto de la gráfica (día D, tercio T) = media de las medias de las ratas.
         SEM = SEM inter-rata de esas medias.

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

    # Para la lógica del título dinámico
    ultimo_seg = {'promedio': None}
    ultimo_rie = {'promedio': None}

    for dia_idx, ruta in enumerate(archivos, start=1):
        try:
            archivo_excel = pd.ExcelFile(ruta)

            ratas_del_dia = {}
            rata_pos = 0

            for hoja in archivo_excel.sheet_names:
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

                ratas_del_dia[rata_pos] = {
                    'seguro': _tercio_split(seg),
                    'riesgo': _tercio_split(rie),
                }

            dias_ratas_raw[dia_idx] = ratas_del_dia

            if not ratas_del_dia:
                log_fn(f"Día {dia_idx}: no se encontraron datos válidos.")
                continue

            # ── Para cada tercio: recoger la media de cada rata ─────────────
            for t_idx in range(3):          # 0 → tercio 1, 1 → tercio 2, etc.
                tercio_num = t_idx + 1

                medias_seg = []
                medias_rie = []

                for rata_data in ratas_del_dia.values():
                    grupo_seg = rata_data['seguro'][t_idx]
                    grupo_rie = rata_data['riesgo'][t_idx]
                    if len(grupo_seg) > 0:
                        medias_seg.append(float(grupo_seg.mean()))
                    if len(grupo_rie) > 0:
                        medias_rie.append(float(grupo_rie.mean()))

                # Media-de-medias y SEM inter-rata
                if medias_seg:
                    prom_seg = round(float(pd.Series(medias_seg).mean()), 2)
                    sem_seg  = round(_sem(medias_seg), 3)
                else:
                    prom_seg, sem_seg = None, None

                if medias_rie:
                    prom_rie = round(float(pd.Series(medias_rie).mean()), 2)
                    sem_rie  = round(_sem(medias_rie), 3)
                else:
                    prom_rie, sem_rie = None, None

                ultimo_seg = {'promedio': prom_seg}
                ultimo_rie = {'promedio': prom_rie}

                x = _x_tercio(dia_idx, tercio_num)

                if prom_seg is not None and prom_seg != 0:
                    xs_seguro.append(x)
                    ys_seguro.append(prom_seg)
                    sems_seguro.append(sem_seg)

                if prom_rie is not None and prom_rie != 0:
                    xs_riesgo.append(x)
                    ys_riesgo.append(prom_rie)
                    sems_riesgo.append(sem_rie)

                filas_resumen.append({
                    'Día':         dia_idx,
                    'Tercio':      tercio_num,
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

    # Título dinámico basado en qué datos hay
    tiene_seg = ultimo_seg.get('promedio') is not None and ultimo_seg['promedio'] != 0
    tiene_rie = ultimo_rie.get('promedio') is not None and ultimo_rie['promedio'] != 0
    if tiene_seg and tiene_rie:
        titulo = 'Discrimination (Conflict vs No-conflict)'
    elif not tiene_seg:
        titulo = 'Threat crossings (Noise/Shock + Light/Food)'
    else:
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


# ── Segunda hoja: datos crudos por rata ───────────────────────────────────────

def _escribir_hoja_raw(ws, dias_ratas_raw, num_dias):
    """
    Hoja 'Desglose tercios':
        etiqueta | Seguro_d1 | Riesgo_d1 | Seguro_d2 | Riesgo_d2 | ...

    Agrupado por rata (r1, r2, …), dentro de cada rata por tercio,
    mostrando los valores crudos de cada ensayo verticalmente.
    Una fila en blanco separa tercios; una fila extra separa ratas.
    """
    nombres_tercio = ['Primer tercio', 'Segundo Tercio', 'Tercer tercio']

    # Encabezado
    header = ['']
    for d in range(1, num_dias + 1):
        header.append(f'Seguro_d{d}')
        header.append(f'Riesgo_d{d}')
    for col_idx, h in enumerate(header, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(bold=True)

    max_ratas = 0
    for ratas in dias_ratas_raw.values():
        if ratas:
            max_ratas = max(max_ratas, max(ratas.keys()))

    current_row = 2

    for rata_idx in range(1, max_ratas + 1):
        cell = ws.cell(row=current_row, column=1, value=f'r{rata_idx}')
        cell.font = Font(bold=True)
        current_row += 1

        for t_idx in range(3):
            segs_por_dia = []
            ries_por_dia = []

            for dia_idx in range(1, num_dias + 1):
                ratas     = dias_ratas_raw.get(dia_idx, {})
                rata_data = ratas.get(rata_idx)
                if rata_data:
                    segs_por_dia.append(rata_data['seguro'][t_idx].tolist())
                    ries_por_dia.append(rata_data['riesgo'][t_idx].tolist())
                else:
                    segs_por_dia.append([])
                    ries_por_dia.append([])

            max_vals = max(
                max((len(s) for s in segs_por_dia), default=0),
                max((len(r) for r in ries_por_dia), default=0),
                1,
            )

            for val_row in range(max_vals):
                row = current_row + val_row
                if val_row == 0:
                    ws.cell(row=row, column=1, value=f'  {nombres_tercio[t_idx]}')

                for dia_idx in range(1, num_dias + 1):
                    seg_col  = 2 + (dia_idx - 1) * 2
                    rie_col  = seg_col + 1
                    seg_vals = segs_por_dia[dia_idx - 1]
                    rie_vals = ries_por_dia[dia_idx - 1]

                    if val_row < len(seg_vals):
                        ws.cell(row=row, column=seg_col,
                                value=round(float(seg_vals[val_row]), 4))
                    if val_row < len(rie_vals):
                        ws.cell(row=row, column=rie_col,
                                value=round(float(rie_vals[val_row]), 4))

            current_row += max_vals
            current_row += 1  # fila en blanco entre tercios

        current_row += 1  # fila extra entre ratas

    # Ajustar anchos
    ws.column_dimensions['A'].width = 20
    for d in range(1, num_dias + 1):
        ws.column_dimensions[get_column_letter(2 + (d - 1) * 2)].width = 14
        ws.column_dimensions[get_column_letter(3 + (d - 1) * 2)].width = 14


# ── Guardado del Excel de resumen ─────────────────────────────────────────────

def guardar_excel_resumen(df_resumen, ruta_xlsx, dias_ratas_raw=None, num_dias=0, log_fn=print):
    """
    Guarda dos hojas:
    1. 'Resumen Tercios'  — media-de-medias y SEM inter-rata por día/tercio.
    2. 'Desglose Tercios' — valores crudos por rata, día y tercio.
    """
    try:
        with pd.ExcelWriter(ruta_xlsx, engine='openpyxl') as writer:
            # Hoja 1
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

            # Hoja 2
            if dias_ratas_raw and num_dias > 0:
                ws_raw = writer.book.create_sheet(title='Desglose Tercios')
                _escribir_hoja_raw(ws_raw, dias_ratas_raw, num_dias)

        log_fn(f"Excel de resumen guardado en: {ruta_xlsx}")
        return True
    except Exception as e:
        log_fn(f"Error guardando Excel: {str(e)}")
        return False