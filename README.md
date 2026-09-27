# Mercadeo

Addon para SAP Business One que facilita la importación de precios y vigencias desde un archivo Excel (.xlsx) hacia una matriz en un formulario llamado "Mercadeo".

Características principales
- Importación de un archivo .xlsx (se busca la hoja llamada "Hoja1" o la primera hoja disponible).
- Copia de columnas: ItemCode, ItemPrice, FromDate, ToDate hacia un DataTable del formulario (DT_Import) y su visualización en una matriz (mtx_import).
- Manejo robusto de fechas (soporta valores DateTime, números OADate y cadenas parseables).

Requisitos
- SAP Business One con SDK (SAPbouiCOM).
- .NET Framework 4.7.2
- Visual Studio (se probó en Community 2026)
- Paquete NuGet: ExcelDataReader (para leer archivos .xlsx)

Estructura esperada del Excel
- Columna A: ItemCode (código de artículo)
- Columna B: ItemPrice (precio)
- Columna C: FromDate (fecha inicio) — valores aceptados: fecha, cadena parseable o número OADate de Excel
- Columna D: ToDate (fecha fin)

Cómo compilar
1. Abrir la solución `Mercadeo.slnx` en Visual Studio.
2. Restaurar paquetes NuGet si es necesario.
3. Compilar la solución (Target framework: .NET Framework 4.7.2).

Cómo usar (en entorno SAP Business One)
1. Instalar/registrar el addon según el procedimiento estándar del SDK.
2. En SAP Business One abrir el menú del addon y lanzar el formulario "Mercadeo".
3. Pulsar el botón para seleccionar el archivo Excel (ItemUID: btn_xls).
4. Seleccionar el .xlsx y confirmar; los registros se cargarán en la matriz.

Validaciones y mensajes
- Si falta el DataTable (DT_Import) o sus columnas, el addon intentará crearlas automáticamente.
- Se mostrará un MessageBox con el número de registros importados.
- Si alguna fila tiene problemas al asignar valores (p. ej. formato de fecha inválido) se mostrará una advertencia en la barra de estado indicando la fila afectada.

Cambios recientes (fixes importantes aplicados)
- Corregido bucle de importación para no omitir la primera fila y evitar repetir siempre la misma fila (se iteraba mal con un índice fijo).
- Mejorada la función FormatearFechaParaSAP para aceptar DateTime, valores OADate (double/int) y cadenas numéricas/formatos locales.
- Evitado el error "No Fields in Table" asegurando la existencia de columnas en DT_Import antes de añadir filas y protegiendo los Bind y LoadFromDataSource con try/catch.

Limitaciones y recomendaciones
- ItemPrice actualmente se asigna como texto al DataTable de UI; si SAP requiere un tipo numérico puede que sea necesario convertirlo antes de asignar.
- Asegúrate de que el Excel contiene al menos las 4 columnas esperadas o ajusta el archivo/formulario según tus necesidades.

Contribuciones
- Pull requests bienvenidos. Mantener compatibilidad con .NET Framework 4.7.2 y el SDK de SAPbouiCOM.

Contacto
- Repositorio original: https://github.com/enahue/Addon-Mercadeo-b1

