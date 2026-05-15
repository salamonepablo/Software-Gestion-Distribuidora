# SPC-Core: Contexto para Rama Retail

**Fecha**: 2026-05-15
**Rama actual durante esta sesión**: `main` (producción)
**Próxima rama a trabajar**: `spc-retail` (punto de venta)

---

## Estado Actual del Proyecto Main (Producción)

### Último Trabajo Realizado (Sesión 2026-05-15)

#### 1. FormOrdenPago: Múltiples mejoras y fixes
**Commits realizados (7 total):**
- `b0e8bb0` - Fix EnLetras: formato de entrada y validación
- `b1db646` - Mejoras impresión: detalles + optimizaciones + ocultar $0
- `b2919ba` - Fix EnLetras: evaluación no-cortocircuito VB6
- `49aa30c` - Agregar importes individuales en detalles
- `a637b2e` - Fix SafeCharAt: ByVal para evitar type mismatch
- `10e2707` - Formato europeo dinámico en textbox Efectivo

**Problemas resueltos:**
- ✅ Bug conversión EnLetras: "1000" → "mil" (antes decía "cien")
- ✅ VB6 And non-short-circuit evaluation en EnLetras
- ✅ Impresión detallada: TODOS los ítems de pago (Cheques, Transferencias, Facturas, Otros)
- ✅ Importes individuales visibles en cada detalle impreso
- ✅ Secciones con $0 ocultas para ahorrar espacio
- ✅ Label "CHEQUE/E-CHEQ" actualizado
- ✅ Formato europeo dinámico en tiempo real (mientras escribís)

#### 2. FormVerPagoFactura: Fix bug movimientos cuenta corriente
**Commit:** `e5b0de1`

**Problema reportado por cliente:**
- Al hacer doble clic en pago desde FormMovimientosCuentaCorriente
- Cliente 2197, pago #183 → Mostraba datos de cliente 24 (¡incorrecto!)
- Root cause: buscodatos() solo filtraba por NroPago, sin verificar cliente

**Solución implementada:**
- Agregada variable `idClienteMov` para capturar ID cliente correcto
- Modificado Form_Load para leer TextCodigoCliente del formulario origen
- Query PagoC: `NroPago AND IdCliente` (antes solo NroPago)
- Query PagoD: `NroPago AND IdSucursal` para consistencia
- ✅ **Probado y funcionando perfectamente**

---

## Arquitectura de Ramas (Variant Branches)

### Main (actual - producción)
- **Target**: ERP completo, full-featured
- **Características**: Facturación A/B/C, créditos/débitos, reportes completos, múltiples depósitos
- **Estado**: Producción activa
- **Sync**: Base para las otras ramas

### SPC-Retail (próxima rama a trabajar)
- **Target**: Punto de venta para distribuidoras
- **Características**: Facturación A/B/C, créditos/débitos, reportes IVA/comisiones, múltiples depósitos
- **Estado**: Draft PR (review-only)
- **Sync**: `git merge origin/main` (one-way, NUNCA merge back to main)
- **Diferencias vs Main**: Orientado a ventas directas, UI optimizada para caja

### SPC-Minimal (no tocar por ahora)
- **Target**: Clientes pequeños, features básicos
- **Características**: Tipos de factura limitados, reportes básicos, sin múltiples depósitos
- **Estado**: Draft PR (review-only)
- **Sync**: `git merge origin/main` (one-way)

---

## Archivos Clave Modificados Recientemente

### FormOrdenPago.frm
- **Líneas clave modificadas**:
  - EnLetras function (2385+)
  - cmdImprimir_Click (2011+)
  - RecalcularTodo (2658+)
  - txtEfectivo_KeyUp (nuevo)
  - FormatImporteTextBoxDynamic (nueva función)

### FormVerPagoFactura.frm
- **Líneas clave modificadas**:
  - Variable idClienteMov (794)
  - Form_Load: captura idClienteMov (1374+)
  - buscodatos(): filtros por cliente (1502, 1530, 1547)

---

## Comandos Git para Cambiar a Retail

```bash
# Verificar estado actual
git status

# Asegurar que todo está commiteado en main
git add .
git commit -m "Pre-switch to retail: save current state"

# Cambiar a rama retail
git checkout spc-retail

# Traer últimos cambios de main (one-way merge)
git merge origin/main

# Si hay conflictos, resolverlos y continuar
```

---

## Cuando Regreses: Checklist

1. ✅ Verificar rama actual: `git branch` (debe decir `spc-retail`)
2. ✅ Leer este archivo: `RETAIL-CONTEXT.md`
3. ✅ Verificar diferencias retail vs main: `git diff main..spc-retail`
4. ✅ Identificar archivos específicos de retail
5. ✅ Aplicar fixes de main si es necesario (cherry-pick o merge)

---

## Notas Importantes

### Para la Rama Retail:
- **NUNCA mergear retail → main** (one-way sync only)
- Cambios específicos de retail quedan en retail
- Bug fixes de main se sincronizan a retail via merge
- Retail puede tener archivos/forms adicionales específicos de POS

### Tecnología:
- VB6 con Access database (DB_SPC_SI.mdb)
- Crystal Reports 8.5, VSReport 8.0
- ADO 2.5, DAO 3.6
- Encoding: ANSI, line endings: CRLF (enforced por .gitattributes)

### Testing:
- Manual testing only (no automated tests)
- Compilar en VB6 IDE: File → Make SPCSI_5.exe
- Probar workflows manualmente

---

## Contacto del Cliente (Patricio)

**Feedback reciente:**
- ✅ EnLetras corregido
- ✅ Impresión con detalles completos
- ✅ Importes individuales en detalles
- ✅ Bug cuenta corriente resuelto
- Próximo: Trabajo en versión retail

---

## Estado de la Base de Datos

**Archivos tracked en Git:**
- `DB_SPC_SI.mdb` (~37 MB) - Base principal
- `DB_SPC_SI_limpia.mdb` - Template (NO tocar manualmente)

**Cambios recientes en DB:**
- Ninguno estructural en esta sesión
- Solo datos de prueba

---

## Para Continuar el Trabajo

1. Lee este archivo completo
2. Checkout a `spc-retail`
3. Revisa diferencias con main
4. Identifica qué necesita el cliente en retail
5. Aplica fixes de main si corresponde
6. Continúa desarrollo específico de retail

**Última actualización**: 2026-05-15 19:55
**Branch actual**: main
**Commits pushed**: 8 (FormOrdenPago + FormVerPagoFactura)
**Estado**: ✅ Todo funcionando en producción
