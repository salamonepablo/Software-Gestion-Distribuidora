# SPC – Variante Retail

Este documento describe el alcance funcional, decisiones técnicas y proceso de mantenimiento de la variante Retail del sistema SPC (VB6 + Access).

- Rama: `spc-retail`
- Propósito: Orientada a operación de punto de venta / distribuidora, con facturación completa, remitos, cuentas corrientes, reportes comerciales e IVA.
- Estado del PR: PR de revisión (Draft) NO MERGE contra `main` (solo para discusión/revisión).

## 1) Objetivos y alcance

- Incluir pantallas y lógica de:
  - Facturación (A/B/NC/ND e internas cuando corresponda)
  - Remitos, presupuestos, cuentas corrientes
  - Búsquedas y listados avanzados (clientes, productos, ventas)
  - Reportes de IVA ventas (según disponibilidad)
  - Comisiones (si aplica en esta variante)
- Excluir herramientas o módulos que no aporten a la operación de retail (ver sección exclusiones).

## 2) Diferencias clave versus otras variantes

| Tema / Módulo                        | Core (main) | Retail (esta rama) | Minimal |
|-------------------------------------|-------------|--------------------|---------|
| Facturación completa (A/B/NC/ND)    | Sí          | Sí                 | Limitado |
| Remitos y Presupuestos              | Sí          | Sí                 | Limitado |
| Libro IVA Ventas / Reportes fiscales| Sí          | Sí                 | No |
| Cálculo de comisiones               | Sí          | Sí (si aplica)     | No |
| Panel de control extendido          | Sí          | Sí                 | No |
| Herramientas auxiliares (limpieza)  | Sí          | Opcional           | No |
| Cuenta corriente completa           | Sí          | Sí                 | Parcial |
| Integraciones (QRs, códigos barras) | Sí          | Sí                 | Parcial |

Ajustar la tabla según lo que realmente quedó en la rama.

## 3) Formularios y módulos principales (ejemplos comunes en la rama)

- Formularios de operación:
  - `FormFactura.frm`, `FormPagoFactura.frm`, `FormPagoFacturaDesdeFactura.frm`
  - `FormRemito.frm`, `FormPresupuesto.frm`
  - `FormVerFacturas.frm`, `FormVerProductos.frm`, `FormVerPresupuestos.frm`
  - `FormBuscarFactura.frm`, `FormBuscarRemito.frm`, `FormBusquedaFacturaPorCliente.frm`
- Reportes y auxiliares:
  - `FormLibroIvaVentas.frm`, `ListadoProductos.rpt`, `ListadoSaldos.rpt`
- Infraestructura:
  - `VariablesPublicas.bas`, `Sentencias.bas`, `SPCSI.vbp`

Nota: Completar/ajustar con la lista real de formularios que utilice Retail.

## 4) Exclusiones y decisiones (documentar)

- Archivos y módulos que NO se usan en Retail:
  - Ejemplos: herramientas de depuración, formularios “viejos”, utilitarios no productivos, etc.
- Políticas de repositorio:
  - No versionar caches (`*.oca`, `Thumbs.db`), ejecutables o librerías registrables (`*.exe`, `*.dll`, `*.ocx`) salvo excepciones justificadas.
  - Para bases Access con datos reales, preferir:
    - Subir una base “estructura” (sin datos) o scripts de creación/poblado anónimo.

## 5) Base de datos (Access)

- Archivo de trabajo: `DB_SPC_SI.mdb` (revisar si se versiona con datos o estructura vacía).
- Si se requieren datos de prueba, documentar método de generación/anónimo.
- Mantener `docs/db-schema.md` (opcional) con:
  - Tablas principales
  - Campos clave
  - Índices/restricciones
  - Cambios de esquema entre versiones

## 6) Configuración del entorno (VB6)

- Requisitos:
  - Visual Basic 6 (IDE), Service Pack aplicable
  - Referencias/Componentes:
    - MS ActiveX Data Objects (ADO)
    - Controles OCX utilizados por los formularios (registrar en el sistema si corresponde)
- Build:
  - Abrir `SPCSI.vbp` (o el proyecto de Retail que corresponda)
  - Compilar desde el IDE
- Notas:
  - Los archivos `.frx` acompañan a `.frm` (recursos binarios).
  - Mantener `CRLF` en archivos `.frm/.bas/.vbp` (ver `.gitattributes`).

## 7) Estructura y archivos de interés

```
/docs/               Documentación
/QRs/                Imágenes QR para comprobantes
/*.frm, *.frx        Formularios VB6
/*.bas               Módulos
/*.rpt               Reportes
/*.vbp               Proyecto VB6
/*.mdb               Base de datos Access (ver política de datos)
```

## 8) Flujo de trabajo y sincronización

- Los cambios comunes se hacen en `main` y se traen a `spc-retail`:
  ```
  git checkout spc-retail
  git fetch origin
  git merge origin/main
  # Resolver conflictos preservando la lógica de Retail
  git push
  ```
- La rama `spc-retail` NO se mergea a `main` (productos con objetivos distintos).
- Para revisión puntual, se abre PR Draft contra `main` y se cierra sin merge.

## 9) Versionado y tags

- Tag inicial sugerido: `v1.0.0-SPC-retail`
- Incrementos:
  - Patch: fixes menores
  - Minor: nuevas pantallas/funciones sin romper compatibilidad
  - Major: cambios estructurales / incompatibles

Ejemplo:
```
git checkout spc-retail
git tag v1.0.0-SPC-retail
git push origin v1.0.0-SPC-retail
```

## 10) Lista de verificación (checklist)

- [ ] Confirmar que `.gitignore` excluye caches y datos sensibles.
- [ ] Documentar OCX/controles requeridos y cómo registrarlos.
- [ ] Documentar dependencias de reportes (si aplica).
- [ ] Completar tabla de diferencias con `main` y `spc-minimal`.
- [ ] Agregar ejemplos de uso / flujo de facturación.
- [ ] Revisar `docs/VARIANTS.md` para enlazar a este documento.

## 11) Historial y notas

- Esta rama se originó con importación inicial y se creó PR Draft NO MERGE exclusivamente para revisión.
- Para ver el diff general vs `main`, abrir PR Draft o usar la vista Compare.