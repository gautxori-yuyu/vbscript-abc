# Plan de Refactorización - VBScript ABC

## Estado: PENDIENTE DE APROBACIÓN DEL USUARIO

---

## Fase 0: Configuración de Codificación (PREREQUISITO)

### Problema
No existe `.gitattributes` ni `.git/info/attributes`. Los ficheros están en ANSI en el clon Linux.

### Acciones
1. Crear `.gitattributes` en la raíz del repo (se commitea):
   ```
   *.vbs text eol=crlf working-tree-encoding=windows-1252
   *.hta text eol=crlf working-tree-encoding=windows-1252
   *.html text eol=crlf working-tree-encoding=windows-1252
   ```
2. Crear `.git/info/attributes` local (NO se commitea):
   ```
   *.vbs text eol=crlf working-tree-encoding=utf-8
   *.hta text eol=crlf working-tree-encoding=utf-8
   *.html text eol=crlf working-tree-encoding=utf-8
   ```
3. Convertir ficheros existentes de ANSI a UTF-8.
4. Verificar caracteres especiales.

---

## Fase 1: Corrección de Bugs Críticos

| ID | Fichero | Línea | Descripción | Fix |
|----|---------|-------|-------------|-----|
| B1 | `fUtils.vbs` | ~34 | `bFileOpen_FromCmdLine` asigna a `bFileOpen` (variable incorrecta) | Cambiar a `bFileOpen_FromCmdLine = ...` |
| B2 | `fUtils.vbs` | ~47 | `MsgDoc` con parámetros invertidos | Invertir orden: nivel, texto |
| B3 | `cHTALogger.vbs` | ~616-619 | `Set` en asignación de strings en `cOpContainers.folderUpdate` | Quitar `Set` |
| B4 | `cHTALogger.vbs` | ~677 | `m_TreeItem` inexistente en `RemoveContainers` | Cambiar a `m_TreeLi` |
| B5 | `cOp_CalcsTecn.vbs` | ~49-55 | `fich.Close` en File object (no tiene método Close) | Usar TextStream intermedio |
| B6 | `cOp_CalcsTecn.vbs` | ~292 | `regex.Execute().Item(0)` sin validar que haya matches | Verificar `.Count > 0` antes |
| B7 | `template_ie.html` | CSS | Clases `.spoiler-header`/`.spoiler-content` no coinciden con VBS | Actualizar a `.spoilerHead`/`.spoilerBody` |
| B8 | `template_ie.html` | JS | Función `showhide()` no existe, la llaman los spoilers | Añadir función `showhide()` |

---

## Fase 2: Eliminación de sentencias `Stop`

Análisis individual de cada `Stop`. Para cada uno:
1. Determinar el motivo (error, depuración, TODO pendiente).
2. Añadir `MsgLog` descriptivo que explique la razón.
3. Eliminar el `Stop`.

### Ficheros afectados:
- `cOportunidad.vbs`: líneas 208, 278
- `cOp_CalcsTecn.vbs`: líneas 295, 338, 344, 346, 348, 355, 407, 417
- `cOp_ValsEcon.vbs`: líneas 154, 188, 191, 255, 276, 295, 364, 386, 393
- `cABCGas.vbs`: línea 155 y otros
- `fUtils.vbs`: línea 252

---

## Fase 3: Migración MsgBox → HtaConfirm/HtaAccept

Sustituir todas las llamadas directas a `MsgBox` en clases de negocio por las funciones globales `HtaConfirm()` y `HtaAccept()`.

### Ficheros afectados:
- `cOportunidad.vbs`: líneas 21, 118, 128
- `cOp_CalcsTecn.vbs`: líneas 101-102, 106-107
- `cOp_ValsEcon.vbs`: líneas 27, 52-55, 61, 92-94, 115, 210-212, 218-219, 242-243, 294
- `cABCGas.vbs`: líneas 178-182, 191

---

## Fase 4: Simplificación de HTMLWindow (Legacy Code)

Eliminar código LEGACY inalcanzable de `fUtils.vbs`:
1. Ramas `Else` (LEGACY fallback) de todos los métodos delegados.
2. Método `Init_Document()` marcado como `[LEGACY]`.
3. Función `fixHTML()`.
4. Variables legacy: `oDicContainerStack`, `unnamedSPCount`, `m_CurrContainer`, `m_Cuerpo`, `m_Log`.

---

## Fase 5: Mejora de Presentación en IE

1. Unificar CSS de spoilers (ya cubierto en B7/B8).
2. Verificar/añadir estilos para `.bordered-table`, `op-container`, `calc-container`.
3. Verificar zona de botones `header-buttons` (ya documentado en sesión anterior).

---

## Fase 6: Pequeñas mejoras estructurales

1. `cOportunidad.vbs:64` — Almacenar referencia a ExcelApp en la clase.
2. `cOportunidad.vbs:206` — Reformatear If de una línea.
3. Revisar bloques `On Error Resume Next` sin correspondiente `GoTo 0`.

---

## Orden de ejecución
0 → 1 → 5 → 2 → 3 → 4 → 6
