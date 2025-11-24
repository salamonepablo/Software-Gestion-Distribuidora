# Dependencias y Componentes VB6

Este documento lista los OCX/DLL/TLB referenciados en `SPCSI.vbp` y pasos para instalarlos / registrarlos en una máquina nueva.

> Nota: Todas las rutas en el .vbp apuntan a `SysWOW64` (entorno 32-bit en Windows 64-bit). Asegúrese de usar `regsvr32` de 32 bits (ubicado normalmente en `C:\Windows\SysWOW64\regsvr32.exe`).

## Lista de Referencias (del .vbp)
| Componente | Archivo | Descripción | Registro |
|------------|---------|-------------|----------|
| OLE Automation | stdole2.tlb | Tipos base OLE | Nativo (ya registrado) |
| VSReport 8.0 | vsrpt8.ocx | Reporting (ComponentOne) | `regsvr32 vsrpt8.ocx` |
| ADO 2.5 | msado25.tlb | Data access | Incluido MDAC |
| Data Environment | MSDE.DLL / MSDERUN.DLL | VB6 Data Environment | Normalmente ya registrado |
| Data Formatting | MSSTDFMT.DLL | Formato de datos | `regsvr32 MSSTDFMT.DLL` |
| Data Report Designer | MSDBRPT.DLL | Reports VB6 | `regsvr32 MSDBRPT.DLL` |
| DAO 3.6 | dao360.dll | Acceso a Jet/Access | `regsvr32 dao360.dll` |
| Printer FlexGrid | Printer FlexGrid.dll | Grid impresión | `regsvr32 "Printer FlexGrid.dll"` |
| Crystal Report Export | sviewhlp.dll | Export helper | `regsvr32 sviewhlp.dll` |
| Crystal Report Viewer | crviewer.dll | Visor CR | `regsvr32 crviewer.dll` |
| Crystal Reports Runtime | craxdrt.dll | Runtime CR 8.5 | `regsvr32 craxdrt.dll` |
| CR Standard Wizard | crystalwizard.dll | Asistente CR | `regsvr32 crystalwizard.dll` |
| Access Object Library | MSACC.OLB | Object Library Access | Provisto por Office |
| Jet & Replication | msjro.dll | Replicación Jet | `regsvr32 msjro.dll` |
| POSUtilts | BarCode.dll | Código de barras | `regsvr32 BarCode.dll` |
| AccessibilityCpl | AccessibilityCpl.dll | Funciones accesibilidad | `regsvr32 AccessibilityCpl.dll` |
| FEAFIPLib | feafip.dll | Facturación electrónica AFIP | `regsvr32 feafip.dll` (si expone COM) |
| Tab Control | TABCTL32.OCX | Pestañas estándar VB6 | `regsvr32 TABCTL32.OCX` |
| Hierarchical FlexGrid | MSHFLXGD.OCX | Grid jerárquico | `regsvr32 MSHFLXGD.OCX` |
| FlexGrid | MSFLXGRD.OCX | Grid estándar | `regsvr32 MSFLXGRD.OCX` |
| ADO Data Control | MSADODC.OCX | Data control | `regsvr32 MSADODC.OCX` |
| Common Controls 2 | mscomct2.ocx | DatePicker, etc. | `regsvr32 mscomct2.ocx` |
| Picture Clip | PICCLP32.OCX | Imagenes | `regsvr32 PICCLP32.OCX` |
| ??? cselexpt | cselexpt.ocx | Export util | `regsvr32 cselexpt.ocx` |

## Pasos de Instalación
1. Instalar VB6 IDE (si se compila localmente) con Service Pack 6.
2. Instalar Crystal Reports 8.5 Runtime (o copiar DLLs y registrar). Algunas requieren dependencias de MFC/Visual C++ runtime antiguo.
3. Instalar MDAC/Jet si no está presente (Windows modernos suelen traerlo ya).
4. Colocar todos los OCX/DLL en `C:\Windows\SysWOW64` (o mantener carpeta propia y ajustar `PATH`).
5. Registrar cada componente necesario:
```powershell
# Ejemplo masivo (ajustar lista según disponibles)
$components = @(
  'vsrpt8.ocx','MSSTDFMT.DLL','MSDBRPT.DLL','dao360.dll','"Printer FlexGrid.dll"',
  'sviewhlp.dll','crviewer.dll','craxdrt.dll','crystalwizard.dll','msjro.dll','BarCode.dll',
  'AccessibilityCpl.dll','feafip.dll','TABCTL32.OCX','MSHFLXGD.OCX','MSFLXGRD.OCX',
  'MSADODC.OCX','mscomct2.ocx','PICCLP32.OCX','cselexpt.ocx'
)
foreach ($c in $components) { & "$env:WINDIR\SysWOW64\regsvr32.exe" /s $c }
```
6. Verificar que `SPCSI.vbp` no marca referencias faltantes al abrir en VB6.

## Notas de Compatibilidad
- Ejecución en Windows 10/11 de componentes antiguos puede requerir permisos elevados.
- Crystal Reports 8.5 es muy legacy; evaluar migración futura.
- Si algún OCX no registra, comprobar dependencias con herramientas como `Dependency Walker` (depends.exe).

## Migración Futura (Ideas)
- Reemplazar grids VB6 por soluciones .NET / modernizar UI.
- Migrar acceso a datos hacia OLEDB más reciente o ADO.NET (si se migra a .NET).
- Encapsular lógica crítica en librerías separadas para facilitar portabilidad.

Mantener este documento actualizado ante cada nueva dependencia.
