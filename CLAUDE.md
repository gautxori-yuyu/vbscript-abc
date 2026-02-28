# Directrices del Proyecto - VBScript ABC (Gestión de Oportunidades)

## Descripción del proyecto
Sistema de gestión de oportunidades comerciales para la venta de compresores industriales.
Desarrollado en VBScript con interfaz HTML en Internet Explorer / HTA.

## Codificación de ficheros
- **En origen (Windows del programador)**: ANSI (Windows-1252), con fin de línea CRLF.
- **En el repositorio Git**: UTF-8 con LF (conversión automática via `.gitattributes`).
- **En el clon Linux (agente IA)**: UTF-8 (override via `.git/info/attributes`).
- Referencia completa: `docs/RECOMENDACIONES-GIT-ANSI.md`

## Principios de desarrollo
1. **Respetar la programación orientada a objetos** y las buenas prácticas de VBScript.
2. **Cambios mínimos**: solo los necesarios para corregir errores y mejorar estructura.
3. **No modificar comentarios existentes** en el código.
4. **Respetar nombres de variables** (solo cambiarlos si es para mayor claridad).
5. **Respetar la estructura del documento HTML** generado en IE.
6. **Respetar los mensajes al usuario** en su contenido (se pueden elaborar para mayor claridad).
7. **Consultar al usuario** ante decisiones que admitan diferentes alternativas.

## Decisiones tomadas (historial)

### Sesión 2026-02-28 (segunda)
- **Sentencias `Stop`**: Analizar cada una individualmente. Determinar su motivo, añadir `MsgLog` descriptivo que explique la razón de la parada, y eliminar el `Stop`. El objetivo final es que no quede ningún `Stop` en producción.
- **Código legacy de HTMLWindow**: Eliminar las ramas LEGACY inalcanzables (fallback que nunca se ejecuta porque `m_Logger` siempre se crea).
- **MsgBox → HtaConfirm/HtaAccept**: Migrar TODAS las llamadas a `MsgBox` en clases de negocio a las funciones globales `HtaConfirm()` y `HtaAccept()`.

## Arquitectura del código

### Punto de entrada
- `procesar carpeta oportunidad.vbs` — Script principal que orquesta el procesado.

### Clases principales
- `cOportunidad` — Gestión de una oportunidad comercial completa.
- `cOp_CalcsTecn` — Procesado de cálculos técnicos (carpeta 2.CALCULO TECNICO).
- `cOp_ValsEcon` — Procesado de valoraciones económicas (carpeta 3.VALORACION ECONOMICA).
- `cOp_Ofertas` — Procesado de ofertas comerciales (carpeta 4.OFERTA COMERCIAL).
- `cCompressor` — Contenedor de datos de un compresor.
- `cABCGas_XLS` — Datos de cálculo técnico de un modelo de compresor.
- `cOferGas` — Documento de oferta comercial (Excel).
- `cBudget` — Presupuesto para cálculos Excel.

### Infraestructura
- `ExcelManager.vbs` — Gestión de Excel (singleton `cExcelApplication`, file manager `cExcelFMFileManager`, wrapper `cExcelFMFile`).
- `cHTALogger.vbs` — Sistema de logging y UI HTML (clase `HtaLogger`, clase `cOpContainers`).
- `fUtils.vbs` — Utilidades generales (clase `HTMLWindow` wrapper de HtaLogger, clase `cWScript`, funciones globales).
- `template_ie.html` — Plantilla HTML/CSS/JS para la interfaz en IE.
- `constants_globals.vbs` — Constantes y patrones regex globales.
