# Recomendaciones de configuración Git para proyectos VBA y VBSCript con ficheros ANSI
(En donde se haga referencia a ficheros VBA, se aplica también a ficheros VBS (VBScript)

> **Contexto de este proyecto:**
> Los ficheros VBA y VBScript (`.cls`, `.bas`, `.frm`, `.vbs`) se editan en **dos entornos distintos**:
> - **Usuario programador (Windows/VBA):** trabaja en ANSI (Windows-1252), requisito del editor VBA de Office.
> - **Agente IA (Linux/Unix):** recibe y edita en UTF-8, requisito de sus herramientas de edición de texto.
> El repositorio Git actúa como capa de conversión entre ambos mundos.

---

## 1. Configuración del repositorio (`.gitattributes`)

El fichero `.gitattributes` en la raíz del repo define la conversión automática para **todos los clones**:

```gitattributes
*.bas text eol=crlf working-tree-encoding=windows-1252
*.cls text eol=crlf working-tree-encoding=windows-1252
*.frm text eol=crlf working-tree-encoding=windows-1252
*.vbs text eol=crlf working-tree-encoding=windows-1252
```

**Qué hace cada directiva:**

| Directiva | Efecto |
|-----------|--------|
| `text` | Git trata el fichero como texto (normaliza EOL, aplica filtros) |
| `eol=crlf` | Working tree siempre en CRLF (necesario para VBA en Windows) |
| `working-tree-encoding=windows-1252` | Working tree en ANSI; repositorio en UTF-8 |

**Flujo resultante:**

```
Commit:   working tree (ANSI/CRLF)  →[Git convierte]→  repo (UTF-8/LF)
Checkout: repo (UTF-8/LF)           →[Git convierte]→  working tree (ANSI/CRLF)
```

Este fichero ya está configurado correctamente. **No modificarlo.**

---

## 2. Carpetas locales del programador (Windows)

### Requisitos
- Trabajar **siempre en ANSI (Windows-1252)** en estas carpetas.
- El editor VBA de Office, y el editor VBSEdit de VBScript, leen y escriben ANSI directamente: ninguna conversión manual es necesaria.
- Al hacer `git add` / `git commit`, Git convierte automáticamente ANSI → UTF-8 gracias al `.gitattributes`.
- Al hacer `git pull` / `git checkout`, Git convierte automáticamente UTF-8 → ANSI.

### Verificación del estado de una carpeta local

Ejecutar desde la carpeta espejo correspondiente:

```cmd
:: 1. Confirmar que el atributo está activo
git check-attr -a clsApplication.cls

:: Salida esperada:
::   clsApplication.cls: text: set
::   clsApplication.cls: eol: crlf
::   clsApplication.cls: working-tree-encoding: windows-1252

:: 2. Confirmar que el fichero local es ANSI
file -i clsApplication.cls
::   Esperado: charset=iso-8859-1  (o windows-1252)
::   Si dice charset=utf-8: el fichero está en UTF-8, ver sección 5.

:: 3. Confirmar que el repo almacena UTF-8
git show HEAD:clsApplication.cls | file -
::   Esperado: ASCII text o Unicode text, UTF-8
```

### Configuración Git local (por carpeta)

No se necesita ninguna configuración adicional en `git config --local` si el `.gitattributes`
está correctamente en el repo. Verificar que no haya overrides que interfieran:

```cmd
git config --local --list
:: No debe aparecer ningún core.autocrlf, working-tree-encoding ni encoding extra
```

Si apareciese `core.autocrlf=true`, eliminarlo:
```cmd
git config --local --unset core.autocrlf
```

---

## 3. Workspace del agente IA (Linux/Unix)

El agente necesita recibir los ficheros en **UTF-8** porque sus herramientas de edición
(`Edit`, `Write`, etc.) trabajan en UTF-8. Para lograrlo **sin alterar el `.gitattributes`**
del repo, se usa un override local en `.git/info/attributes`:

### Fichero `.git/info/attributes` (en el clon Linux)

```gitattributes
*.bas text eol=crlf working-tree-encoding=utf-8
*.cls text eol=crlf working-tree-encoding=utf-8
*.frm text eol=crlf working-tree-encoding=utf-8
*.vbs text eol=crlf working-tree-encoding=utf-8
```

> **Nota:** `working-tree-encoding=utf-8` es un no-op (repo UTF-8 → working tree UTF-8),
> pero anula el `windows-1252` del `.gitattributes` del repo solo para este clon.
> El fichero `.git/info/attributes` es local al clon y **nunca se commitea al repo**.

Este fichero ya está configurado en el clon Linux de este proyecto. Si se crea un clon nuevo
en Linux, hay que recrearlo manualmente (ver sección 4).

### Verificación desde Linux

```bash
# Atributo efectivo (debe mostrar utf-8, no windows-1252)
git check-attr working-tree-encoding -- clsApplication.cls
# Esperado: clsApplication.cls: working-tree-encoding: utf-8

# Codificación real del fichero en working tree
file -i clsApplication.cls
# Esperado: charset=utf-8

# Contenido en el repo (debe ser UTF-8 limpio, sin mojibake)
git cat-file -p HEAD:clsApplication.cls | python3 -c "
import sys; raw=sys.stdin.buffer.read()
highs=[b for b in raw if b>0x7F]
print(f'Bytes altos: {len(highs)}')
if highs:
    i=next(i for i,b in enumerate(raw) if b>0x7F)
    print(raw[max(0,i-10):i+15].decode('utf-8',errors='replace'))
"
```

---

## 4. Creación de un nuevo clon Linux (receta completa)

```bash
git clone <url-repo> <directorio>
cd <directorio>

# Crear override local para recibir UTF-8
cat > .git/info/attributes << 'EOF'
*.bas text eol=crlf working-tree-encoding=utf-8
*.cls text eol=crlf working-tree-encoding=utf-8
*.frm text eol=crlf working-tree-encoding=utf-8
*.vbs text eol=crlf working-tree-encoding=utf-8
EOF

# Re-checkout para que Git aplique la nueva configuración
# (los ficheros recién clonados habrán llegado en ANSI por el .gitattributes del repo)
python3 -c "
import glob
for f in glob.glob('*.cls')+glob.glob('*.bas')+glob.glob('*.frm')+glob.glob('*.vbs'):
    with open(f,'rb') as fh: raw=fh.read()
    try:
        raw.decode('utf-8')
    except UnicodeDecodeError:
        text=raw.decode('cp1252')
        open(f,'wb').write(text.encode('utf-8'))
        print('Converted:', f)
"

echo "Clon Linux listo. Todos los ficheros están en UTF-8."
```

---

## 5. Diagnóstico y reparación de problemas de codificación

### Problema: fichero ANSI en el clon Linux (tras merge/rebase/pull)

**Síntoma:** `file -i fichero.cls` dice `charset=iso-8859-1` en el clon Linux.

**Causa:** Git re-checkout el fichero y aplicó la conversión UTF-8→ANSI
(a veces ocurre tras rebase o reset si el override local se pierde).

**Solución:**
```bash
python3 -c "
import glob
for f in glob.glob('*.cls')+glob.glob('*.bas')+glob.glob('*.frm')+glob.glob('*.vbs'):
    with open(f,'rb') as fh: raw=fh.read()
    try: raw.decode('utf-8')
    except:
        text=raw.decode('cp1252')
        open(f,'wb').write(text.encode('utf-8'))
        print('Converted:', f)
"
```

### Problema: mojibake en el repo (ï¿½ en lugar de ó, é, etc.)

**Síntoma:** el repo almacena `\xc3\xaf\xc2\xbf\xc2\xbd` donde debería haber `\xc3\xb3` (ó).

**Causa:** se commiteó un fichero UTF-8 (con U+FFFD) cuando `working-tree-encoding=windows-1252`
estaba activo, haciendo que Git lo recodificara incorrectamente.

**Cómo detectarlo:**
```bash
git cat-file -p HEAD:clsApplication.cls | python3 -c "
import sys; raw=sys.stdin.buffer.read()
n=raw.count(b'\xc3\xaf\xc2\xbf\xc2\xbd')
print(f'Mojibakes encontrados: {n}')
"
```

**Solución:** editar el fichero afectado, buscar las ocurrencias de `ï¿½` (o U+FFFD `?`)
y sustituirlas por el carácter correcto deducido del contexto. Commitear.

### Problema: fichero UTF-8 en carpeta local Windows

**Síntoma:** `file -i` dice `charset=utf-8` en la carpeta espejo Windows.

**Causa:** el fichero fue guardado en UTF-8 (por VS Code u otro editor) sin pasar por Git.

**Solución:** forzar re-checkout desde el repo (Git convertirá UTF-8→ANSI):
```cmd
git checkout -- nombre-fichero.cls
:: o para todos los ficheros:
git checkout -- "*.cls" "*.bas" "*.frm" "*.vbs"
```

---

## 6. Regla de oro: nunca editar ficheros fuera de Git

Cualquier modificación de ficheros `.cls`, `.bas`, `.frm`, `.vbs`  que **no pase por Git**
(guardar desde un editor externo directamente en la carpeta espejo) puede romper
la cadena de conversión. La única excepción es el editor, que siempre
escribe ANSI y es compatible con el flujo descrito.

---

## Resumen rápido

| Quién | Codificación working tree | Cómo se configura |
|-------|--------------------------|-------------------|
| Programador Windows | ANSI (Windows-1252) | `.gitattributes` del repo (automático) |
| Agente IA Linux | UTF-8 | `.git/info/attributes` local del clon |
| Repositorio (ambos) | UTF-8 | `.gitattributes` del repo (automático) |
