# Ejercicio: Crear y Subir una Rama Nueva

Este ejercicio te guía para crear una nueva rama, hacer un cambio y subirla a GitHub para luego (opcional) crear un Pull Request.

## 1. Verificar conexión con GitHub
Comprueba que existe el remoto `origin`:
```powershell
git remote -v
```
Debe mostrar URLs de tu repo en GitHub. Si no está:
```powershell
git remote add origin https://github.com/salamonepablo/Software-Gestion-Distribuidora.git
```
(Solo si no se configuró antes.)

## 2. Actualizar tu copia local
Antes de crear la rama, asegúrate de tener lo último:
```powershell
git pull origin main
```
Si la rama principal se llama `master`, usa `master`.

## 3. Crear la rama nueva
Usa un nombre descriptivo. Ejemplo: agregar documentación de cambios.
```powershell
git checkout -b feature/guia-rama
```
Esto crea y cambia a la rama.

## 4. Realizar un cambio
Edita algún archivo (por ejemplo añadir una línea en `README.md`). Guarda.

## 5. Ver estado y preparar commit
```powershell
git status
```
Si ves tus cambios:
```powershell
git add README.md
# o para todo:
git add .
```

## 6. Crear commit
```powershell
git commit -m "Añade sección sobre creación de ramas"
```

## 7. Subir rama a GitHub
```powershell
git push --set-upstream origin feature/guia-rama
```
La primera vez requiere `--set-upstream`. Luego `git push` basta.

## 8. Crear Pull Request (PR)
Ve al repositorio en GitHub. Verás un aviso para comparar y crear un Pull Request. Haz clic y completa:
- Título claro.
- Descripción breve del cambio.

## 9. Revisar y fusionar
Una vez conforme, puedes hacer merge a `main` desde la interfaz web. Después vuelve a tu local y sincroniza:
```powershell
git checkout main
git pull origin main
```

## 10. Limpiar rama local (opcional)
Tras el merge:
```powershell
git branch -d feature/guia-rama
```

## 11. Si aparece error de autenticación
Genera un token personal (PAT) en GitHub y úsalo como contraseña cuando PowerShell te lo pida. Configura tu usuario si no lo hiciste:
```powershell
git config --global user.name "Tu Nombre"
git config --global user.email "tuemail@dominio.com"
```

## 12. Confirmar que todo salió bien
Revisar en GitHub que la rama está y el PR se creó/mergeó.

## Resumen Rápido de Comandos
```powershell
git remote -v
git pull origin main
git checkout -b feature/guia-rama
# (editar archivos)
git add .
git commit -m "Mensaje"
git push --set-upstream origin feature/guia-rama
```

Repite este proceso para futuras mejoras. Mantén los nombres de rama cortos y descriptivos.
