# Hook pre-commit (Ejemplo)

Un hook `pre-commit` permite bloquear la inclusión accidental de ejecutables, librerías o certificados.

## Instalación
1. Copia el archivo ejemplo:
   ```powershell
   Copy-Item .\hooks\pre-commit.sample .\.git\hooks\pre-commit
   ```
2. Asegúrate de que tenga permisos de ejecución (Windows normalmente basta así).
3. Haz un commit de prueba.

## Hook ejemplo
Archivo: `hooks/pre-commit.sample`
Bloquea si detecta alguno de estos patrones en staging: `*.exe, *.dll, *.ocx, *.pfx, *.crt, *.key`.

## Desactivar temporalmente
Puedes saltar hooks usando:
```powershell
git commit -m "Mensaje" --no-verify
```
(Usar solo en casos justificados.)

## Actualizar patrones
Editar el script del hook para ajustar la lista.

Mantén este archivo fuera de `.git/hooks/` para poder versionar el ejemplo.
