# Analizador de Incidencia Policial — Versión 1

Aplicación web estática preparada para GitHub Pages. El archivo Excel se procesa **en el navegador del usuario** y el nombre del archivo no es relevante: la herramienta identifica la estructura a partir de los encabezados de las columnas.

## Funciones incluidas

- Carga de archivos `.xlsx` / `.xls` con la estructura del cubo suministrado.
- Procesamiento en Web Worker para evitar bloquear la interfaz con archivos grandes.
- Filtros por:
  - `UNIDAD`
  - `UnidadFormal` (delegación)
  - `FechaIngreso` (desde / hasta)
  - `Provincia`
  - `Canton`
  - `Distrito`
  - `Barrio`
  - `Tipo_Incidente`
  - `Origen`
  - `SubtipoIncidente`
- Indicadores de total filtrado, georreferenciación, delegaciones y hora pico.
- Indicador y porcentaje de `NO UBICADO / OTROS` sobre el total filtrado.
- **Data reloj** de 24 horas para visualizar cuándo se concentran los incidentes.
- Gráfico de principales tipos de incidente.
- Mapa con tres estilos:
  - Normal (CARTO Voyager)
  - Calles (OpenStreetMap)
  - Oscuro (CARTO Dark)
- Consulta puntual de incidentes por coordenadas `Latitud` / `Longitud`.
- Al presionar un punto o una fila se abre el detalle con la **DiligenciaPolicial** y datos asociados.
- Generación de informe PDF institucional con encabezado/pie derivados del machote aportado.

## Rendimiento del mapa

Para archivos con cientos de miles de puntos, el mapa consulta únicamente los incidentes que están dentro de la vista actual. Si en una vista existen más de 25.000 puntos, muestra una muestra representativa y lo indica en pantalla. Al acercar el mapa, se consultan nuevamente los puntos y se obtiene el detalle puntual del sector.

## Publicar en GitHub Pages

1. Cree un repositorio nuevo en GitHub.
2. Suba **todo el contenido de esta carpeta**, manteniendo `assets/`.
3. Abra `Settings` → `Pages`.
4. En `Build and deployment`, seleccione `Deploy from a branch`.
5. Seleccione la rama `main` y la carpeta `/ (root)`.
6. Guarde y espere a que GitHub publique la dirección de la aplicación.

No requiere compilación, Node.js ni Visual Studio Code.

## Uso

1. Abra la aplicación publicada.
2. Presione **Seleccionar Excel**.
3. Seleccione el cubo de incidentes (puede tener un nombre diferente).
4. Espere a que finalice la lectura.
5. Aplique filtros.
6. Navegue entre Panel, Mapa, Incidentes e Informes PDF.

## Estructura esperada

La versión 1 reconoce los encabezados principales del archivo analizado, entre ellos:

`idIncidente`, `Origen`, `Tipo_Incidente`, `SubtipoIncidente`, `DiligenciaPolicial`, `Provincia`, `Canton`, `Distrito`, `Barrio`, `Incidente`, `Direccion`, `TipoLugar`, `FechaIngreso`, `Hora`, `FechaFinalizado`, `Indicativo`, `UNIDAD`, `UnidadFormal`, `DiferenciaEnMinutos`, `EstadoIncidente`, `FechaCreado`, `Latitud`, `Longitud`, `HoraIncidente`, `RANGO`, `ANIO`, `MES`, `DIA`, `NUMSemana`, `Fuente`, `CodSUBUnidad`, `Producto`.

## Nota sobre privacidad

La columna `DiligenciaPolicial` puede contener información sensible. La versión 1 la muestra únicamente en la vista de detalle del incidente y no la incluye en el PDF estadístico. La aplicación debe publicarse y utilizarse conforme a los controles de acceso institucionales correspondientes.

## Dependencias externas

Se cargan desde CDN:

- SheetJS — lectura del Excel.
- Leaflet + MarkerCluster — mapa.
- ECharts — gráficos.
- jsPDF + AutoTable — informes PDF.

Los mapas utilizan teselas de OpenStreetMap y CARTO, con sus atribuciones visibles en el mapa.
