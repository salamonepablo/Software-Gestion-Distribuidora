# SPCSI (VB6)

Proyecto legacy en Visual Basic 6. Este repositorio busca estabilizar y versionar el código fuente, gestionar dependencias y facilitar releases reproducibles.

## Objetivos
- Control de versiones (Git + tags SemVer).
- Documentar dependencias (OCX / DLL) necesarias para build y runtime.
- Script para automatizar cambio de versión en `SPCSI.vbp`.
- Flujo de ramas sencillo para mantenimiento y hotfixes.

## Estructura Principal
- `SPCSI.vbp`: archivo de proyecto VB6 donde se definen formularios y número de versión.
- Archivos `.frm`, `.frx`, `.bas`: código y recursos de la aplicación.
- Base de datos Access `DB_SPC_SI.mdb` (principal) y backups históricos.
- Carpeta `tools/`: utilidades (scripts PowerShell) para versionado.
- Carpeta `docs/`: documentación adicional.

## Versionado (Semantic Versioning)
Se usa formato `MAJOR.MINOR.PATCH` sobre los campos:
- `MajorVer`
- `MinorVer`
- `RevisionVer` (se mapea a PATCH)

Ejemplo: `1.4.12` => `MajorVer=1`, `MinorVer=4`, `RevisionVer=12`.

### Cuándo incrementar
- MAJOR: cambios incompatibles, estructura de DB, gran refactor.
- MINOR: nuevas funcionalidades retro-compatibles.
- PATCH: correcciones, ajustes menores, fixes urgentes.

## Script de versión
Utilidad: `tools/Increment-Version.ps1`.
Permite incrementar o fijar versión directamente en el `.vbp`.

Uso:
```powershell
# Incrementar patch (revision) por defecto
powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1

# Incrementar minor
powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -Increment minor

# Incrementar major
powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -Increment major

# Fijar versión exacta
powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -SetVersion 1.5.0
```
El script muestra la versión previa y la nueva.

## Flujo de Ramas
- `main`: estado estable / últimas releases.
- `develop` (opcional si hay evolución continua): integración de nuevas features antes de preparar release.
- `feature/<nombre>`: ramas para desarrollar funcionalidades específicas.
- `release/<version>`: estabilizar antes del tag (tests, ajustes finales).
- `hotfix/<issue>`: correcciones urgentes partiendo desde `main`.

### Proceso de Release
1. Asegurar cambios integrados en `develop` (o directamente en `main` si el flujo es simple).
2. Crear rama `release/x.y.z` si se usa flujo con `develop`.
3. Ejecutar script de versión para fijar número definitivo.
4. Commit: `git commit -am "Release x.y.z"`.
5. Tag anotado: `git tag -a vX.Y.Z -m "Release vX.Y.Z"`.
6. Push código y tags:
```powershell
git push
git push --tags
```
7. Publicar binario compilado (si aplica) en la sección Releases de GitHub.

## Compilación VB6 por línea de comandos
Si VB6 IDE está instalado (ruta típica `"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"`):
```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make SPCSI.vbp /out build.log
```
Verificar que todas las dependencias (OCX/DLL) estén registradas (`regsvr32 nombre.ocx`).

## Dependencias
Listado detallado en `docs/Dependencies.md`. Incluye componentes Crystal Reports, ADO, DAO, OCXs de grids y reportes.

## Base de Datos
- Mantener sólo la versión activa `DB_SPC_SI.mdb`.
- Backups históricos están ignorados (modificar `.gitignore` si se requiere preservarlos).

## Buenas Prácticas
- Commits pequeños y descriptivos.
- Evitar subir binarios compilados; usar Releases para distribuir.
- Documentar cambios relevantes en `CHANGELOG.md` (crear si se desea).
- Revisar diferencias en `.vbp` antes de hacer tag (para confirmar versión).

## Próximos Pasos
- Completar `docs/Dependencies.md`.
- Crear `CHANGELOG.md` (opcional).
- Automatizar release con GitHub Actions (build + adjuntar artefacto).
- Usar `tools/Verify-RepoState.ps1` antes de un commit grande para asegurar que no haya `.exe` y que `DB_SPC_SI.mdb` esté presente.

---
Cualquier mejora futura: contenedores de build, migración gradual, tests automatizados alrededor de lógica crítica.
