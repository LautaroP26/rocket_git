Informe sobre el Flujo de Trabajo del Bot en RocketBot

1. Introducción

Este documento detalla el flujo de trabajo del bot desarrollado en RocketBot, describiendo los módulos y comandos utilizados para su funcionamiento. El bot tiene como objetivo procesar información desde un archivo de configuración .ini y una planilla de facturación en Excel, identificando clientes con un IVA superior a un umbral y creando hojas de trabajo personalizadas para cada uno.

2. Requisitos Previos

Para la correcta ejecución del bot, es necesario contar con:

RocketBot instalado.

Módulos requeridos:

ControlIni (versión 1.3.3)

AdvancedExcel (versión 34.37.21 o superior)

Archivo de configuración: config2.ini

Planilla de facturación: Planilla de facturacion - ACONCAGUA 02-2025.xlsx

3. Flujo de Trabajo del Bot

3.1 Inicialización

Definición de Variables Iniciales:

Se establece la variable cont con un valor inicial de 4.

Se establece la variable fila_procesada con un valor inicial de 0.

3.2 Lectura del Archivo de Configuración (.ini)

Uso del módulo ControlIni para leer el archivo config2.ini.

Extracción de Datos:

Se obtiene la variable flag para verificar si la lectura fue exitosa.

Se extrae la ruta del archivo Excel.

Manejo de Errores:

Si la lectura falla, se genera una alerta indicando el error.

3.3 Procesamiento de la Planilla de Facturación

Apertura del Archivo Excel:

Se abre el archivo Planilla de facturacion - ACONCAGUA 02-2025.xlsx.

Se verifica si el archivo se cargó correctamente.

Extracción de Datos:

Se cuenta el número total de filas en la columna C.

Se obtienen los nombres de los clientes listados en la columna C.

Manejo de Errores:

Si el archivo Excel no se puede abrir, se muestra una alerta de error.

3.4 Evaluación del IVA y Creación de Hojas Personalizadas

Iteración sobre la Lista de Clientes:

Se recorre la lista de clientes obtenida.

Para cada cliente, se obtiene el valor del IVA desde la columna L.

Condicional de Evaluación:

Si el IVA del cliente es mayor o igual a $500,000, se realizan las siguientes acciones:

Se crea una nueva hoja en el archivo Excel con el nombre del cliente.

Se copia el formato de la cabecera desde la hoja principal.

Se copian los valores y formatos correspondientes.

Manejo del Contador:

Se incrementa la variable cont en 1 en cada iteración.

4. Módulos y Comandos Utilizados

A continuación, se detallan los módulos y comandos empleados en el bot:

4.1 Módulo ControlIni

leerIni: Lee el archivo de configuración y almacena los datos en variables.

obtenerDato: Extrae información específica del archivo .ini.

4.2 Módulo AdvancedExcel

countRows: Cuenta el número de filas en una columna específica.

createSheet: Crea una nueva hoja de trabajo con un nombre determinado.

copyPasteFormat: Copia el formato de celdas de una hoja a otra.

copyPaste: Copia valores y formatos de un rango de celdas a otro.

4.3 Comandos de Excel

getcell: Obtiene datos de una celda o rango de celdas.

readxlsx: Abre y lee archivos Excel.

4.4 Comandos de Lógica y Control de Flujo

evaluateIf: Evalúa condiciones lógicas para la toma de decisiones.

for: Itera sobre una lista de valores.

setVar: Modifica el valor de una variable.

alert: Muestra alertas en pantalla en caso de errores.

5. Conclusiones

Este bot de RocketBot automatiza la gestión de facturación al procesar un archivo Excel, evaluando el IVA de cada cliente y generando hojas personalizadas según los criterios definidos. Gracias al uso de los módulos ControlIni y AdvancedExcel, el proceso es eficiente y flexible, permitiendo una adaptación fácil a diferentes escenarios.

6. Recomendaciones

Se recomienda actualizar el módulo AdvancedExcel a la versión más reciente para evitar errores de compatibilidad.

Mantener una copia de seguridad del archivo Excel antes de ejecutar el bot.

Implementar manejo avanzado de errores para mejorar la robustez del sistema.

📌 Autor: Lautaro Pedernera📆 Versión: 12
