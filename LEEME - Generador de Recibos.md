# Generador de Recibos de Sueldo

Esta carpeta contiene una aplicación local para generar recibos de sueldo a partir del Excel completo de liquidación. La solución fue diseñada para Windows 11 sin requerir instalación de Python, Node.js ni otras herramientas externas.

NUEVO ENFOQUE: Ahora también proporcionamos un CLI en Node.js que genera recibos HTML imprimibles y un CSV resumen. Puedes convertir fácilmente cada recibo a PDF usando Puppeteer si lo deseas.

## Cómo iniciar

Abra el archivo **Iniciar Aplicacion Recibos.cmd**. Se abrirá la aplicación en el navegador predeterminado de la computadora.

Luego presione **Cargar liquidación** y seleccione el archivo **LiquidaciondeSueldosAbril12.xlsx** o cualquier Excel con la misma estructura de hojas.

## Datos que toma del Excel

La aplicación lee automáticamente las siguientes hojas y parámetros:

| Hoja | Datos usados |
|---|---|
| **Empleador** | Razón social, CUIT, domicilio legal, localidad, provincia, código postal y domicilio de pago. |
| **Empleados** | Legajo, fecha de ingreso, apellido y nombre, CUIL, sueldo básico, cuenta, calificación, tarea/categoría y sistema previsional. |
| **Remuneración** | Haberes remunerativos, adicionales, acuerdos, horas extras, SAC y vacaciones cuando existan importes. |
| **Ret. y no Rem.** | Descuentos, anticipos, aportes, conceptos no remunerativos, subtotal de descuentos y sueldo neto. |

## Funciones disponibles

La aplicación permite seleccionar un empleado, ver la vista previa del recibo, imprimir o guardar como PDF el recibo seleccionado, imprimir todos los recibos y exportar un resumen CSV de la liquidación.

## Observación importante

El sistema utiliza los valores guardados en el Excel y una interpretación automática de las columnas detectadas. Antes de emitir recibos definitivos, conviene validar los resultados contra la normativa laboral y previsional aplicable y contra la liquidación contable final.
