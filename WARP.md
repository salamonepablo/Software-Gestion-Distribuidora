# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Resumen del proyecto

Este repositorio contiene un ERP de escritorio para distribuidoras, desarrollado en **Visual Basic 6** con base de datos **Access**. El sistema cubre facturación electrónica adaptada a la normativa de AFIP (Argentina), cuentas corrientes de clientes, reportes fiscales (Libro IVA Ventas) y reportes comerciales (comisiones, ranking de ventas, etc.).

El ejecutable principal se genera a partir del proyecto `SPCSI.vbp` y produce un EXE tipo `SPCSI_4.exe`.

## Comandos y scripts útiles

> Todos los ejemplos asumen PowerShell (`pwsh` o `powershell`) ejecutado desde la raíz del repo.

### Hooks de Git / control de binarios

**Instalar / actualizar el hook `pre-commit`** (recomendado tras clonar el repo):

```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\install-hooks.ps1
```

Esto crea wrappers en `.git/hooks/` que invocan `hooks/pre-commit.ps1`. El hook bloquea commits que incluyan archivos como `*.exe`, `*.dll`, `*.ocx`, `*.pfx`, `*.crt`, `*.key`.

**Limpiar binarios ya versionados** (sin borrarlos del disco, solo del índice):

```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\Clean-Binaries.ps1         # limpieza real
pwsh -ExecutionPolicy Bypass -File .\tools\Clean-Binaries.ps1 -DryRun # solo listar
```

Tras ejecutar, revisar con `git status` y hacer commit de los cambios.

### Verificar el estado del repositorio

Script de chequeo rápido de salud del repo: presencia de `.exe`, base de datos, archivos fuente VB6, etc.

```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\Verify-RepoState.ps1
pwsh -ExecutionPolicy Bypass -File .\tools\Verify-RepoState.ps1 -FailOnWarnings # devuelve error en warnings
```

En entornos tipo Unix también existe un wrapper:

```bash
./tools/verify-repo-state.sh
```

### Versionado del ejecutable VB6

El script de versión trabaja sobre el archivo `.vbp` (por defecto `SPCSI.vbp`) y actualiza `MajorVer`, `MinorVer`, `RevisionVer` y opcionalmente el nombre del EXE.

Incrementos típicos:

```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1                 # +1 en PATCH
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -Increment minor
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -Increment major
```

Fijar una versión concreta y actualizar el nombre del EXE:

```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -SetVersion 1.4.3 -UpdateExeName
```

El script muestra la versión anterior/nueva y sugiere el comando `git commit` + `git tag` correspondiente.

### Build y ejecución del sistema

No hay script de build automatizado en CLI. Para compilar o ejecutar el sistema se usa el IDE de Visual Basic 6:

1. Abrir `SPCSI.vbp` en VB6.
2. Asegurarse de que están registradas todas las dependencias COM/OCX indicadas en el proyecto (`MSHFLXGD.OCX`, `TABCTL32.OCX`, `crviewer.dll`, componentes Crystal Reports, ADO, etc.).
3. Verificar que la base `DB_SPC_SI.mdb` está accesible en la ruta esperada.
4. Compilar el proyecto para generar el EXE (`SPCSI_4.exe` por defecto) o ejecutar desde el IDE.

Actualmente **no hay suite de tests automatizados ni scripts de linting** definidos en el repo.

### Flujo de ramas / variantes (Core, Retail, Minimal)

El mismo código base se mantiene en tres variantes de ramas:

- `main` (Core): base general del producto.
- `spc-retail`: variante orientada a punto de venta / distribuidora con funcionalidad completa de facturación, remitos, cuentas corrientes, reportes IVA y comisiones.
- `spc-minimal`: variante reducida para clientes con necesidades básicas (facturación limitada, menos pantallas y reportes).

Las ramas de variantes se mantienen mediante merges **desde** `main` hacia cada variante, usando un flujo como:

```bash
git checkout spc-retail
git fetch origin
git merge origin/main
# Resolver conflictos respetando las personalizaciones de la variante
git push
```

Los PRs desde `spc-retail` / `spc-minimal` hacia `main` se usan solo como **Draft / NO MERGE** para revisión.

## Arquitectura y estructura de alto nivel

### Tecnologías principales

- **Frontend / UI**: Formularios VB6 (`*.frm` + `*.frx`) con MDI (`MDIForm1.frm`) y formularios de menú principal (`MenuPrincipal.frm`).
- **Lógica de negocio**: Código en formularios y módulos `.bas` (por ejemplo `VariablesPublicas.bas`).
- **Persistencia**: Base de datos Access (`DB_SPC_SI.mdb`) accesible vía DAO/ADO y objetos de entorno de datos (Data Environment, Data Report) definidos en el proyecto.
- **Reportes**: Crystal Reports 8.5, VSReport y Data Report de VB6, con archivos `*.rpt` y vistas previas en formularios dedicados.

### Proyecto VB6 principal (`SPCSI.vbp`)

El archivo `SPCSI.vbp` define:

- Tipo de proyecto `Exe`.
- Referencias COM/OCX a ADO, DAO, Crystal Reports, VSReport, controles de rejillas (`MSHFLXGD.OCX`, `MSFLXGRD.OCX`), controles comunes y librerías específicas (por ejemplo `feafip.dll` para integración con AFIP, `BarCode.dll` para códigos de barras/QR).
- Listado de formularios principales, entre ellos:
  - **Maestros**: `Clientes.frm`, `Articulos.frm`, `Empleados.frm`, `Depositos.frm`, formularios de direcciones de clientes, etc.
  - **Operación de venta**: `FormFactura.frm`, `FormRemito.frm`, `FormPresupuesto.frm`, `FormNotaCredito.frm`, `FormNotasdeDebito.frm`, formularios de búsqueda (`FormBuscarFactura.frm`, `FormBuscarRemito.frm`, etc.).
  - **Cuentas corrientes y cobranzas**: `FormMovimentosCuentaCorriente.frm`, `FormMovCCFechas.frm`, `FormVerFacturaCtaCte.frm`, formularios de recibos (`FormRecibo.frm`) y pagos (`FormPagoFactura.frm`, `FormPagoFacturaDesdeFactura.frm`).
  - **Reportes**: `FormLibroIvaVentas.frm`, `FormLiqComisiones.frm`, `FormListadoVentas.frm`, listados de clientes por vendedor, listados de consignaciones, etc.
  - **Utilitarios**: formularios de importación (`FormImportTxt.frm`), panel de control (`frmPanelControl.frm`), menú de búsqueda de formularios (`FormMenuBuscarFormularios.frm`), etc.

El formulario de inicio (`Startup`) y el ícono principal están configurados a `MenuPrincipal` en el `.vbp`, con un formulario MDI (`MDIForm1`) que actúa como contenedor de la mayoría de pantallas.

### Organización lógica por capas (implícita)

No existe una separación formal en capas al estilo moderno; la organización típica es:

- **UI + lógica de presentación**: cada formulario maneja eventos de usuario, validaciones y armado de consultas SQL.
- **Lógica de negocio compartida**: concentrada en módulos `.bas` como `VariablesPublicas.bas` (variables y constantes globales) y otros módulos auxiliares que encapsulan sentencias SQL o funciones comunes.
- **Acceso a datos**: uso directo de ADO/DAO dentro de formularios y módulos, apuntando a la base `DB_SPC_SI.mdb`.
- **Reportes**: cada reporte se vincula a formularios específicos que preparan los datos y llaman a los viewers de Crystal/VSReport.

Esto implica que cambios en la estructura de la base o en reglas de negocio suelen requerir ediciones coordinadas en varios formularios y, a veces, en módulos `.bas`.

### Datos y base Access

- Archivo principal: `DB_SPC_SI.mdb` en la raíz del repo.
- La documentación de variantes sugiere tratar esta base como estructura de referencia (idealmente sin datos reales) y, cuando se necesiten datos de prueba, usar cargas de datos anonimizadas o bases vacías.
- Para cambios estructurales importantes se recomienda mantener documentación de esquema (por ejemplo en `docs/`), aunque puede no estar completamente actualizada.

### Carpetas y archivos relevantes

- `docs/`
  - `Hooks.md`: documentación del hook `pre-commit` y su instalación.
  - `variants/retail.md`: detalles funcionales y técnicos de la variante **Retail** (formularios clave, alcance funcional, requisitos de VB6, estrategia de ramas y tags).
  - `variants/minimal.md`: descripción de diferencias entre **Core**, **Retail** y **Minimal**, más recomendaciones de estructura y versionado.
- `hooks/`
  - Scripts de ejemplo para el hook `pre-commit` basados en PowerShell.
- `tools/`
  - Scripts PowerShell y shell para limpieza de binarios, verificación de estado del repo, instalación de hooks y manejo de versiones.
- Raíz del repo
  - Formularios VB6 (`*.frm` + `*.frx`), módulos (`*.bas`), reportes (`*.rpt`), proyecto `SPCSI.vbp` y la base `DB_SPC_SI.mdb`.

Al trabajar en este repositorio, priorizar siempre mantener compatibilidad con VB6 (encoding ANSI, finales de línea CRLF en archivos de código VB6) y respetar la política de no añadir nuevos binarios/secretos al control de versiones; usar los scripts de `tools/` para ayudar a mantener ese estado.