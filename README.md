# Script de Autocompletado y Consolidación de Datos Excel
# autocompletar_consolidado_v0.2
## Descripción

Este script de Python automatiza el proceso de consolidar información de múltiples archivos Excel (generados con códigos específicos) en un archivo Excel principal (`CONSOLIDADO.xlsx`). Lee un archivo maestro, busca archivos de datos individuales basados en códigos presentes en el maestro y copia datos específicos de esos archivos individuales a columnas designadas en el archivo consolidado final.

Debido a la naturaleza y el tamaño del archivo `CONSOLIDADO.xlsx`, el script ha sido diseñado para trabajar de manera eficiente con la memoria, leyendo el archivo maestro en modo de solo lectura y generando un **nuevo archivo consolidado** con los datos actualizados, en lugar de modificar el original directamente.

## Características

* **Lectura Eficiente:** Lee el archivo `CONSOLIDADO.xlsx` en modo de solo lectura (`read_only=True`) para manejar eficientemente archivos grandes y evitar problemas de `MemoryError`.
* **Actualización por Código:** Busca archivos individuales (`[CODIGO].xlsx`) basados en los códigos presentes en la Columna D del archivo consolidado.
* **Volcado de Datos Específicos:** Copia valores de celdas predefinidas de los archivos individuales a columnas específicas (`F`, `G`, `I`, `K`, `M`) del archivo consolidado.
* **Generación de Nuevo Archivo:** Produce un nuevo archivo `CONSOLIDADO_COMPLETADO.xlsx` con todas las modificaciones, preservando el archivo original.
* **Manejo de Errores:** Incluye un robusto manejo de errores para archivos no encontrados, problemas de lectura y otros fallos durante el proceso, mostrando mensajes claros en la consola.
* **Copia de Hojas No Procesadas:** Asegura que cualquier hoja del `CONSOLIDADO.xlsx` original que no esté especificada para el procesamiento (`EDU`, `HOSP`, `EMPRESA`) sea copiada al archivo de salida.
* **Ejecutable Portable:** Se puede empaquetar en un ejecutable (`.exe`) usando PyInstaller para facilitar su distribución y uso sin necesidad de tener Python instalado.

## Cómo Funciona (Resumen Técnico)

1.  El script carga `CONSOLIDADO.xlsx` en **modo de solo lectura**. Esto es crucial para archivos grandes, ya que `openpyxl` los maneja como un stream de datos, evitando que todo el archivo se cargue en la RAM de una sola vez para su modificación.
2.  Itera a través de las hojas `EDU`, `HOSP` y `EMPRESA`.
3.  Para cada fila en estas hojas, lee el `código` de la columna D.
4.  Si el archivo `[CODIGO].xlsx` existe en el mismo directorio, lo abre (en modo `data_only=True`, para leer solo valores).
5.  Extrae los valores de las celdas especificadas (`N21`, `N25`, `N49`, `N153`, `N161`) y actualiza la fila correspondiente en una representación en memoria.
6.  A medida que procesa cada fila, esta se escribe directamente en un **nuevo libro de Excel vacío** (`CONSOLIDADO_COMPLETADO.xlsx`). Esto asegura que el archivo final contenga los datos actualizados sin las limitaciones de memoria del archivo original.
7.  Una vez procesadas todas las hojas objetivo, el script copia cualquier otra hoja existente en `CONSOLIDADO.xlsx` que no haya sido procesada a la salida.
8.  Finalmente, el nuevo libro de Excel con todos los datos actualizados y las hojas copiadas se guarda como `CONSOLIDADO_COMPLETADO.xlsx`.

## Requisitos

* Python 3.x
* Librería `openpyxl`

## Instalación

1.  **Clona el repositorio:**
    ```bash
    git clone [https://github.com/SebasRodMag/AutoConsolidado.git]
    cd AutoConsolidado
    ```
2.  **Instala las dependencias:**
    ```bash
    pip install openpyxl
    ```

## Uso

### Archivos Necesarios

Asegúrate de que los siguientes archivos estén en el **mismo directorio** que el script (o el ejecutable):

* `CONSOLIDADO.xlsx` (tu archivo maestro grande)
* Todos los archivos de datos individuales con formato `[CODIGO].xlsx` (ej. `1801381.xlsx`, `180101E.xlsx`, etc.)

### Ejecutar el Script (Python)

```bash
python autocompletar_consolidado.py
```

### Generar Ejecutable (PyInstaller)

Para crear un archivo ejecutable único (Windows, Linux) que no requiera la instalación de Python en el sistema de destino:

* Instala PyInstaller: 
```bash
pip install pyinstaller
```
* Genera el ejecutable: 
```bash
pyinstaller.exe --onefile --console --icon=icono.ico autocompletar_consolidado.py`
```

``--onefile:`` Crea un único archivo ejecutable.

``--console:`` Muestra una ventana de consola para ver el progreso y los errores.

``--icon=icono.ico:`` (Opcional) Asigna un icono personalizado al ejecutable. Asegúrate de que icono.ico esté en el mismo directorio.

El ejecutable se encontrará en la carpeta ``dist/.`` Deberás copiar el ejecutable (autocompletar_consolidado.exe) junto con ``CONSOLIDADO.xlsx`` y los archivos ``[CODIGO].xlsx`` al directorio donde desees ejecutarlo.

### Solución a los problemas
``MemoryError:`` Este error indica que el script se queda sin memoria RAM. La versión actual del script lo maneja leyendo CONSOLIDADO.xlsx en modo de solo lectura y generando un nuevo archivo. Si el error persiste, el problema podría estar en el tamaño o la complejidad de los archivos ``[CODIGO].xlsx`` o en la memoria disponible en el sistema. Asegúrate de que tus archivos Excel estén optimizados (sin filas/columnas vacías excesivas o formato innecesario).

``FileNotFoundError ``o Estado: Archivo no encontrado:

Asegúrate de que CONSOLIDADO.xlsx y todos los archivos ``[CODIGO].xlsx`` estén en el mismo directorio donde se ejecuta el script o el ejecutable.

Verifica que los nombres de los archivos ``[CODIGO].xlsx`` coincidan exactamente con los códigos en la Columna D de ``CONSOLIDADO.xlsx`` (sensible a mayúsculas/minúsculas y espacios extra).

### Limitaciones
El script genera un nuevo archivo (``CONSOLIDADO_COMPLETADO.xlsx``) en lugar de modificar el ``CONSOLIDADO.xlsx ``original. Deberás reemplazar manualmente el archivo original si ese es tu objetivo final.

El script no maneja formatos de celdas complejos, fórmulas o macros al copiar datos; solo copia los valores de las celdas.`

# autocompletar_consolidado_v1.0
## Descripción

Este script de Python automatiza el proceso de consolidar información en un archivo Excel, sin la necesidad de múltiples archivos Excel donde obtener información.

modificamos las celdas especificas para añadir la ruta especifica al dato del archivo especifico que necesitamos obtener.

Dado a la cantidad de rutas especificas diferentes, se utiliza este script para insertar las funciones especificas en cada celda de forma automática, generando un **nuevo archivo consolidado** con los datos actualizados.

### Ejecutar el Script (Python)

```bash
python autocompletar_consolidado_v1.0.py
```

### Autor
* Sebastián Rodríguez

* GitHub: SebasRodMag