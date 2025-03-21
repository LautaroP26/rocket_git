# 🚀 LautaroP_EX2 - Bot en RocketBot

## 📌 Descripción  
Este es un bot desarrollado en RocketBot como parte del segundo ejercicio de BotSolution. Su función principal es leer un archivo `.ini`, obtener datos de una planilla de facturación en Excel y realizar operaciones basadas en el IVA de los clientes.

## 🛠️ Requisitos  
Antes de ejecutar el bot, asegúrate de contar con:  
- **RocketBot** instalado.  
- Los módulos requeridos en RocketBot:  
  - `ControlIni` (versión 1.3.3)  
  - `AdvancedExcel` (mínimo versión 34.37.21, actualizar a 34.39.28 si es posible).  
- Archivo de configuración `config2.ini`.  
- Archivo de Excel `Planilla de facturación - ACONCAGUA 02-2025.xlsx`.

## 🔄 Flujo de Funcionamiento  

1. **Inicialización**  
   - Se establece una variable `cont` con un valor inicial de `4`.  
   - Se establece una variable `fila_procesada` con un valor inicial de `0`.  

2. **Lectura del Archivo `.ini`**  
   - Se obtiene la variable `flag` del archivo de configuración.  
   - Si la lectura es exitosa, se extrae la ruta del archivo Excel.  
   - Si falla, se muestra una alerta de error.  

3. **Procesamiento del Archivo Excel**  
   - Se abre la planilla de facturación.  
   - Se cuenta el total de filas en la columna `C`.  
   - Se obtienen los nombres de los clientes en la columna `C`.  

4. **Evaluación de IVA y Creación de Hojas**  
   - Se recorre la lista de clientes.  
   - Para cada cliente, se obtiene su valor de IVA en la columna `L`.  
   - Si el IVA es mayor o igual a `$500,000`:  
     - Se crea una nueva hoja en Excel con el nombre del cliente.  
     - Se copia el formato de cabecera desde la hoja principal.  
     - Se copian los valores y formatos correspondientes.  
   - Se incrementa el contador `cont`.  

## ⚠️ Errores y Manejo de Excepciones  
- Si el archivo `.ini` no se puede leer, se muestra una alerta.  
- Si la planilla de Excel no se abre correctamente, también se notifica con un mensaje de error.  

## 📝 Notas Adicionales  
- La variable `cont` se usa para iterar y controlar la posición en la planilla.  
- Se recomienda actualizar `AdvancedExcel` a la última versión disponible.  
- El bot maneja clientes con IVA alto de manera automática, creando hojas personalizadas para cada uno.

---

📌 **Autor:** Lautaro Pedernera  
📆 **Versión:** 12  

