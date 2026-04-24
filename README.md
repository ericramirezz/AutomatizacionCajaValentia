# Caja Valentía Analyzer

**Caja Valentía** es una herramienta de automatización desarrollada para optimizar el flujo de trabajo en laboratorios de neurociencias (específicamente para el IFC de la UNAM). El software procesa archivos de datos crudos (`.mat`) provenientes de ensayos conductuales, automatizando la limpieza de datos, el cálculo de latencias y la generación de reportes estadísticos visuales y tabulares.

Esta herramienta fue diseñada específicamente para trabajar con **Pruebas de Conflicto**, permitiendo a los investigadores centrarse en la interpretación de resultados en lugar del procesamiento manual de datos.

## Características Principales

* **Procesamiento Automatizado:** Lectura directa de archivos `.mat` de MATLAB y conversión a estructuras de datos de Pandas.
* **Análisis por Bloques (Tercios):** División automática de sesiones experimentales en tres bloques para el análisis de tendencias intra-sesión (aprendizaje, fatiga, conflicto).
* **Visualización Científica:** Generación de gráficas inter-rata con barras de error (SEM).
* **Exportación Consolidada:** Creación de reportes en Excel (`.xlsx`) con hojas de resumen, promedios por rata y desglose detallado de cada ensayo.
* **Interfaz de Usuario:** UI intuitiva construida con `CustomTkinter`, que incluye una consola de logs en tiempo real y un fondo animado de red neuronal.

## Metodologíaa

La lógica de procesamiento de este software y el diseño del protocolo de análisis se basan estrictamente en la metodología descrita en el siguiente artículo científico:

> **Illescas-Huerta, E., Ramirez-Lugo, L., Sierra, R. O., Quillfeldt, J. A., & Sotres-Bayon, F. (2021).** *Conflict Test Battery for Studying the Act of Facing Threats in Pursuit of Rewards.* Frontiers in Neuroscience, 15, 645769. 
> [DOI: 10.3389/fnins.2021.645769](https://doi.org/10.3389/fnins.2021.645769)

**Si utilizas este software en tu investigación, por favor asegúrate de citar el artículo original mencionado arriba.**

## 🛠️ Requisitos e Instalación

El sistema requiere **Python 3.10 o superior** y las siguientes dependencias:

```bash
pip install pandas matplotlib scipy openpyxl customtkinter Pillow

o se puede descargar el .exe ubicado en el Release del repositorio. Si se trabaja en Windows, se recomienda instalar situar todo el repositorio directamente sobre el disco C: (el .exe)
correrá sin hacerlo, pero de esta forma se tendrá acceso al logo del laboratorio y los scripts que componen el código del programa, así como la documentación.
