# Git y GitHub: Guía Básica

Esta guía te ayuda a dar los primeros pasos usando Git con tu proyecto VB6.

## 1. Conceptos Clave
- Repositorio: carpeta con tu código + historial.
- Commit: un "snapshot" de cambios con mensaje descriptivo.
- Push: enviar commits locales a GitHub.
- Tag: marca una versión específica (release) p.ej. `v1.0.0`.
- Rama (branch): línea de desarrollo paralela (por ahora puedes trabajar en `main`).

## 2. Flujo mínimo diario
1. Editas archivos (.frm, .bas, .vbp, etc.).
2. Verificas estado:
   ```powershell
   git status
   ```
3. Añades los cambios:
   ```powershell
   git add .
   ```
4. Creas commit:
   ```powershell
   git commit -m "Mensaje claro del cambio"
   ```
5. Subes a GitHub:
   ```powershell
   git push
   ```

## 3. Verificación antes de commit
Ejecuta el script que hicimos:
```powershell
powershell -ExecutionPolicy Bypass -File .\tools\Verify-RepoState.ps1
```
Si muestra ejecutables, elimínalos del control de versiones:
```powershell
git rm --cached ruta\al\archivo.exe
```

## 4. Versionar (cambiar número de versión)
Incrementar (patch por defecto):
```powershell
powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1
```
Luego commit y push:
```powershell
git commit -am "Bump version"
git push
```

## 5. Crear un Tag (Release)
Cuando quieras marcar una versión estable:
```powershell
git tag -a v1.0.1 -m "Release v1.0.1"
git push --tags
```
En GitHub verás el tag y puedes crear un Release (opcional) para adjuntar binario compilado.

## 6. Ver historial
```powershell
git log --oneline
```
Usa `q` para salir si abre paginador.

## 7. Recuperar estado (si algo salió mal)
- Ver qué cambió sin agregar: `git diff`
- Cancelar cambios en un archivo que no has committeado:
  ```powershell
  git restore archivo.frm
  ```
- Revertir último commit (si aún no lo subiste):
  ```powershell
  git reset --soft HEAD~1   # mantiene cambios en staging
  git reset --hard HEAD~1   # descarta cambios definitivos
  ```
  ¡Cuidado con `--hard`!

## 8. Crear una rama (más adelante)
```powershell
git branch feature/mi-cambio
git checkout feature/mi-cambio
# o en una sola línea:
git checkout -b feature/mi-cambio
```
Cuando acabes:
```powershell
git add .
git commit -m "Mi cambio"
git push --set-upstream origin feature/mi-cambio
```
Luego puedes abrir un Pull Request en GitHub.

## 9. Buenas Prácticas de Mensajes
- Usa presente: "Agrega validación de fechas".
- Prefiere commits pequeños y con propósito.
- Evita mensajes genéricos como "Cambios".

## 10. Ejercicio Inicial Recomendado
1. Ejecuta verificación.
2. Incrementa versión a `1.0.1`.
3. Commit: `git commit -am "Bump version to 1.0.1"`.
4. Tag: `git tag -a v1.0.1 -m "Release v1.0.1"`.
5. Push y push tags.
6. Verifica en GitHub que aparece el tag.

## 11. Si aparece error de autenticación
Configura tu usuario y correo (una sola vez):
```powershell
git config --global user.name "Tu Nombre"
git config --global user.email "tuemail@dominio.com"
```
Si pide credenciales, usa PAT (token personal) de GitHub como contraseña (lo generas en Settings > Developer Settings > Personal Access Tokens).

## 12. Próximos Pasos
- Aprender Pull Requests.
- Añadir un CHANGELOG.
- Automatizar build con GitHub Actions.

Mantén esta guía cerca para referencia rápida.
