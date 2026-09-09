# Excel Lite + AI

Visor y editor de Excel nativo para Windows, liviano y rápido: detecta tus archivos recientes, los abre con vista estilo Excel (colores, fórmulas, imágenes), y trae un asistente de IA integrado para consultar y modificar tus hojas.

<p align="center">
  <img src="https://cdn.jsdelivr.net/gh/NanaBru/ExcelLite-AI@master/assets/app-screenshot.png?v=2" alt="Excel Lite - pantalla principal" width="480"/>
</p>

## Demo

**Abrir un Excel**

<p align="center">
  <img src="https://cdn.jsdelivr.net/gh/NanaBru/ExcelLite-AI@master/assets/demo.gif" alt="Abriendo un Excel en Excel Lite" width="600"/>
</p>

**Asistente de IA: preguntar y aplicar cambios**

<p align="center">
  <img src="https://cdn.jsdelivr.net/gh/NanaBru/ExcelLite-AI@master/assets/demo-ia.gif" alt="Consultando al Asistente IA y aplicando el cambio que sugiere en la hoja" width="600"/>
</p>

## Características

- **Detección automática** de archivos Excel recientes (`.xlsx`, `.xlsm`, `.xls`), con caché deduplicado sin impacto en el rendimiento.
- **Arrastrar y soltar** archivos o carpetas directo a la ventana.
- **Vista estilo Excel real**: colores de celda, negritas, fórmulas calculadas, imágenes embebidas, múltiples hojas con pestañas desplazables.
- **Vista dividida**: mirá dos hojas del mismo libro al mismo tiempo, lado a lado.
- **Edición y guardado**: seleccioná, copiá y pegá rangos de celdas como en Excel; editá directo desde la grilla o la barra de fórmulas; siempre hay filas/columnas en blanco extra para crecer la hoja.
- **Deshacer/Rehacer**, Guardar como, Buscar en todo el libro y Reemplazar, ordenar por columna, favoritos, exportar a CSV.
- **Asistente de IA integrado**: conectá OpenRouter o cualquier proveedor compatible con la API de OpenAI, preguntale sobre tus datos y dejá que proponga (y aplique) cambios — incluso crear hojas nuevas.
- **Ligero de verdad**: GC en modo workstation, vista acotada de filas/columnas, y libera memoria al cerrar cada archivo.
- **Interfaz Windows 11**: tema oscuro con acentos verdes, barra de título personalizada, ventana maximizable sin tapar la barra de tareas.

## Descarga rápida

¿No querés compilar nada? Bajá el `.zip` ya compilado desde la sección [**Releases**](https://github.com/NanaBru/ExcelLite-AI/releases/latest), descomprimilo y ejecutá `ExcelLite.exe`.

## Requisitos

- Windows 10/11
- [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) (o el SDK si vas a compilar desde el código)

## Compilar y ejecutar

```powershell
dotnet build -c Release
```

El ejecutable queda en `bin\Release\net8.0-windows\ExcelLite.exe`.

## Asistente de IA

Desde la vista de un Excel abierto, tocá **Asistente IA → ⚙ Configurar** y cargá:

- **Proveedor**: OpenRouter, o cualquier otro compatible con la API de chat de OpenAI.
- **Modelo** (ej. `openai/gpt-4o-mini`).
- **API Key** (se guarda localmente en tu equipo).

## Stack técnico

- .NET 8 / WPF
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) para lectura/escritura con estilos, fórmulas e imágenes
- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) como respaldo para el formato `.xls` antiguo
