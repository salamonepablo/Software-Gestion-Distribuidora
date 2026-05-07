# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Project Overview

**SPC-Core** is a production-grade **Enterprise Resource Planning (ERP) system** built in **Visual Basic 6** for battery distribution companies (distribuidoras) in Argentina. It manages invoicing (facturas, notas de crédito/débito), delivery notes (remitos), receipts, inventory, customer accounts, sales commissions, and fiscal compliance reporting (AFIP-compliant Libro IVA Ventas).

**Technology Stack**:
- **Language**: Visual Basic 6 (VB6)
- **Database**: Microsoft Access (DB_SPC_SI.mdb, ~37 MB)
- **Reporting**: Crystal Reports 8.5, VSReport 8.0, VB6 Data Report
- **UI**: VB6 Forms (MDI architecture, 100+ forms)
- **Data Access**: ADO 2.5, DAO 3.6
- **Key Libraries**: feafip.dll (AFIP integration), BarCode.dll (QR/barcodes), crviewer.dll (Crystal Reports)

**Project Files**:
- `SPCSI.vbp` - Main VB6 project
- `DB_SPC_SI.mdb` - Main database (tracked in Git)
- `MenuPrincipal.frm` - Main menu entry point
- `MDIForm1.frm` - MDI container for child forms

---

## Architecture & Code Organization

### High-Level Architecture

The codebase follows a **3-tier implicit architecture**:

```
┌─────────────────────────────────────────────┐
│  UI LAYER: 100+ VB6 Forms (MDI-based)       │
│  MenuPrincipal.frm (entry), MDIForm1.frm    │
│  Forms for invoices, payments, reports, etc.│
└──────────────┬──────────────────────────────┘
               ↓
┌──────────────────────────────────────────────┐
│ BUSINESS LOGIC: Form handlers + Modules      │
│ VariablesPublicas.bas (global state)         │
│ Sentencias.bas (shared SQL)                  │
│ Event-driven, procedural style               │
└──────────────┬──────────────────────────────┘
               ↓
┌──────────────────────────────────────────────┐
│ DATA LAYER: ADO/DAO Recordsets + SQL         │
│ Direct SQL queries in forms and modules      │
└──────────────┬──────────────────────────────┘
               ↓
┌──────────────────────────────────────────────┐
│ PERSISTENCE: Microsoft Access DB + Reports  │
└──────────────────────────────────────────────┘
```

**Key characteristics**:
- **Monolithic**: Most logic lives in forms + shared modules
- **Event-driven**: Business logic triggered by UI events (button clicks, form load, etc.)
- **Global state**: Public variables in `VariablesPublicas.bas` (recordsets, vendor codes, warehouse IDs, line counters, etc.)
- **No formal class structure**: Pure procedural VB6 style
- **Direct SQL**: Most forms construct and execute SQL queries directly (not abstracted)

### Module Organization

| Module Type | Purpose | Examples |
|------------|---------|----------|
| **Business Forms** | Invoice, payment, budget, delivery note creation & editing | `FormFactura.frm`, `FormPagoFactura.frm`, `FormRemito.frm` |
| **Search Forms** | Find and display existing documents | `FormBuscarFactura.frm`, `FormBuscarRemito.frm` |
| **Master Data Forms** | Customer, product, warehouse, employee management | `Clientes.frm`, `Articulos.frm`, `Depositos.frm` |
| **Report Forms** | Tax journals, commission reports, sales listings | `FormLibroIvaVentas.frm`, `FormLiqComisiones.frm` |
| **Utility Forms** | Printing, imports, admin tasks | `FormImprimeRemito.frm`, `FormImportTxt.frm` |
| **Global Logic Modules** | Shared variables, SQL, API declarations | `VariablesPublicas.bas`, `Sentencias.bas`, `Declaraciones.bas` |

### Module-Level Architecture

**Core modules**:

- **VariablesPublicas.bas**: Declares `Public Database BaseSPC`, public recordsets (tClientes, tProductos, tFacturaC, etc.), and config variables (codVendedor, codDeposito, linea). This is the shared state hub.
- **Sentencias.bas**: Common SQL queries and database operations reused across forms.
- **Declaraciones.bas**: API and external library declarations (feafip.dll, BarCode.dll, etc.).

**Form patterns**:

Most forms follow this structure:
1. Module-level recordsets and variables
2. `Form_Load()`: Initialize UI, open recordsets, populate dropdowns
3. Event handlers: Button clicks, list selections, text edits
4. Helper subs/functions: Validation, SQL construction, record operations (Add, Update, Delete)
5. Cleanup: `Form_Unload()` closes recordsets

### Current Database Structure

The Access database contains tables for:
- **Customers** (Clientes): Client master data
- **Products** (Articulos): Product/article catalog
- **Invoices** (Facturas): Invoice headers and line items
- **Payments** (Pagos): Payment records
- **Stock**: Warehouse inventory by deposit
- **Accounts Receivable** (Cuentas Corrientes): Customer ledgers
- **Delivery Notes** (Remitos): Shipment tracking
- **Receipts** (Recibos): Receipt records
- **Debit/Credit Notes**: Adjustments and corrections
- **Commissions**: Sales commission calculations

---

## Building, Running, and Development Commands

### Compilation & Execution

**To build from VB6 IDE**:
1. Open `SPCSI.vbp` in Visual Basic 6 IDE
2. Go to **File → Make SPCSI_5.exe** (or use Ctrl+Shift+F2)
3. Output: Executable in the project root directory
4. Run directly or via double-click

**To compile from command line**:
- VB6 has no official CLI build tool; use IDE or third-party tools (e.g., `VB6.exe /make` via Windows COM)

### Version Management

**Increment version**:
```powershell
# Patch version (e.g., 1.0.1 → 1.0.2)
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1

# Minor version (e.g., 1.0.1 → 1.1.0)
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -Increment minor

# Set specific version (e.g., 1.4.3)
pwsh -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -SetVersion 1.4.3 -UpdateExeName
```

Reads/updates version from `SPCSI.vbp`, updates exe name (`SPCSI_5.exe` with version suffix).

### Git Hooks & Repository Management

**Install pre-commit hook** (prevents committing binaries, DLLs, certs):
```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\install-hooks.ps1
```

Hook blocks: `*.exe`, `*.dll`, `*.ocx`, `*.pfx`, `*.crt`, `*.key`

**Verify repository state** (checks for tracked binaries, large files, etc.):
```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\Verify-RepoState.ps1
```

**Clean binaries from Git history** (remove tracked `.exe`, `.dll`, etc.):
```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\Clean-Binaries.ps1
```

### Testing

**No automated test suite** is currently configured. Testing is manual:
1. Compile from IDE
2. Run executable
3. Test workflows manually (invoicing, payments, reports, etc.)
4. Validate database state in Access

---

## Key Workflows & Common Tasks

### Adding a New Form

1. **Create form**: **File → New Form** in VB6 IDE
2. **Design UI**: Add controls (textboxes, buttons, listboxes, etc.)
3. **Declare module-level variables**: Recordsets, SQL strings, validation flags
4. **Implement `Form_Load()`**: Initialize recordsets, populate dropdowns, set defaults
5. **Add event handlers**: Button clicks, list selections, edits
6. **Add helper subs**: SQL construction, validation, record operations
7. **Implement `Form_Unload()`**: Close recordsets gracefully
8. **Add to MDIForm1** (if needed): Register in `frmMDI.Show` or menu
9. **Update `MenuPrincipal.frm`**: Add menu item to launch the form

### Modifying a Report

Reports exist in two forms:

**Crystal Reports** (`.rpt` files):
1. Open in **Crystal Reports 8.5** or newer
2. Modify report design, add/remove fields, adjust formatting
3. Save (`.rpt` will be tracked in Git)
4. Update form that displays the report (`FormLibroIvaVentas.frm`, etc.)

**VB6 Data Report**:
1. Open form in IDE (e.g., `FormImprimeRemito.frm`)
2. Locate the associated DataReport object
3. Edit design, bind fields to recordsets
4. Update form code to pass correct data

### Adding a Database Field

1. **Backup database**: Save `DB_SPC_SI.mdb` before modifying
2. **Modify Access schema**: Open `DB_SPC_SI.mdb` in Access, add column to relevant table
3. **Update VB6 code**: Modify forms/modules that reference the table
   - Update SQL in `Sentencias.bas`
   - Update recordset declarations in `VariablesPublicas.bas`
   - Update form field bindings and event handlers
4. **Test**: Recompile, test data insertion/retrieval
5. **Commit**: Git tracks `.mdb` changes; include .frm/.bas file changes

### Creating an Invoice

**High-level flow** (see `FormFactura.frm` for implementation):
1. User opens **Facturas** menu → `FormFactura.frm` loads
2. Selects customer (opens `Clientes.frm` if needed)
3. Enters line items: product, quantity, price
4. Validates: Stock check, tax calculation, customer credit limits
5. Saves to database (Facturas table)
6. Generates delivery note (`FormRemito.frm`) or receipt
7. Later: Payment recorded in `FormPagoFactura.frm`

---

## Important Conventions

### File Encoding & Line Endings

**VB6 source files MUST use ANSI encoding and CRLF line endings**. This is enforced via `.gitattributes`:

```
*.frm text eol=crlf
*.bas text eol=crlf
*.vbp text eol=crlf
```

If you edit in a text editor, ensure CRLF is preserved. VB6 IDE automatically handles this.

### Naming Conventions

- **Forms**: `Form{Feature}.frm` (e.g., `FormFactura.frm`, `FormPagoFactura.frm`)
- **Modules**: `{Category}.bas` (e.g., `VariablesPublicas.bas`, `Sentencias.bas`)
- **Reports**: `{Name}.rpt` (Crystal Reports) or DataReport in form
- **Controls**: Hungarian notation (e.g., `txtCliente`, `cmdGuardar`, `lstProductos`)

### Global State Management

All public variables live in `VariablesPublicas.bas`:
```vb
Public BaseSPC As Database
Public tClientes, tProductos, tStock As Recordset
Public codVendedor As String
Public codDeposito As String
Public linea As Integer
```

When adding global state, declare here. Avoid scattering public variables across modules.

### Database Connectivity

All database access goes through `BaseSPC` (Access Database object). Connection is established in `MenuPrincipal.frm` or `MDIForm1.frm`:

```vb
Set BaseSPC = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
```

Forms use DAO to open recordsets:
```vb
Set tClientes = BaseSPC.OpenRecordset("SELECT * FROM Clientes WHERE activo=True")
```

Never hardcode file paths; use `App.Path` to reference the database location relative to the executable.

### Shared SQL Statements

Common queries belong in `Sentencias.bas`. Forms should reference them rather than duplicating SQL. Example:
```vb
' In Sentencias.bas
Public Function GetClientesByName(nombre As String) As String
    GetClientesByName = "SELECT * FROM Clientes WHERE nombre LIKE '%" & nombre & "%' ORDER BY nombre"
End Function

' In Form
Set tClientes = BaseSPC.OpenRecordset(GetClientesByName(txtNombre.Text))
```

---

## Variant Branches

The project maintains **three variant branches** for different deployment scenarios:

| Aspect | Core (main) | Retail (spc-retail) | Minimal (spc-minimal) |
|--------|-----------|------------------|-------------------|
| **Target** | Full-featured ERP | Point-of-sale distribuidora | Small clients, basic features |
| **Invoicing** | A/B/C types, credits, debits | A/B/C types, credits, debits | Limited types |
| **Reporting** | Full IVA books, commissions | Full IVA, commissions | Basic exports only |
| **Warehouses** | Multiple deposits | Yes | No |
| **Status** | Production | Draft PR (review-only) | Draft PR (review-only) |

**Synchronization approach**:
- Bug fixes and core features go into `main` first
- Synced to variants via `git merge origin/main`
- Variant-specific changes are never merged back to main
- Retail and Minimal branches are review-only; never merged into main

---

## Configuration Files

| File | Purpose |
|------|---------|
| `SPCSI.vbp` | Project manifest; lists forms, modules, references (ADO, DAO, Crystal Reports, VSReport), version info (1.0.1), entry point |
| `DB_SPC_SI.mdb` | Main Access database (tracked in Git); contains all persistent data |
| `.gitignore` | Ignores: `.exe`, `.dll`, `.ocx`, `.lic`, `.pdf`, `.jpg`, backups; excludes QuilplacVB folder (external) |
| `.gitattributes` | Enforces CRLF line endings for `.frm`, `.bas`, `.vbp` |
| `Schema.ini` | Column definitions for text/CSV imports (used by `FormImportTxt.frm`) |

---

## Git Workflow

**Main branch**: `main` - Production code, stable releases

**Variant branches** (review-only, never merge to main):
- `spc-retail` - Retail variant
- `spc-minimal` - Minimal variant

**Typical workflow**:
1. Create feature/bugfix branch from `main`
2. Make changes (edit `.frm`, `.bas`, modify `DB_SPC_SI.mdb` if schema changes)
3. Test manually in VB6 IDE
4. Commit with descriptive message
5. Create PR, request review
6. Merge to `main` once approved
7. Variants sync via `git merge origin/main` (one-way)

**Important**:
- Never manually change `DB_SPC_SI_limpia.mdb` (clean template)
- Pre-commit hook prevents committing binaries, DLLs, certs
- Line endings must stay CRLF in VB6 files (enforced by `.gitattributes`)

---

## Debugging & Troubleshooting

### Common Issues

**"Database is locked" (`DB_SPC_SI.ldb` present)**:
- VB6 IDE or executable has the database open
- Close all instances, delete `.ldb` file, retry

**Missing OCX controls** (compile error):
- Register controls: `regsvr32 MSHFLXGD.OCX`, `regsvr32 TABCTL32.OCX`, etc.
- Or: Repair VB6 IDE, reinstall missing components

**AFIP integration fails**:
- Ensure `feafip.dll` is registered: `regsvr32 feafip.dll`
- Check certificate paths and permissions (used in fiscal compliance features)

**Barcode/QR generation fails**:
- Ensure `BarCode.dll` is registered: `regsvr32 BarCode.dll`

### Debugging in IDE

1. **Set breakpoints**: Click in gutter next to code line
2. **Step through code**: F8 (step), F7 (step into), Shift+F8 (step out)
3. **Watch variables**: Debug → Add Watch, inspect values
4. **Print to immediate window**: `Debug.Print varName` in code
5. **Locals window**: View all local variables in scope

---

## Recent Development Focus

**Latest work** (from recent commits):
- Printing functionality for budget forms (layout refinements)
- Delivery note printing improvements
- Invoice printing with detailed layout and data retrieval
- Form label captions and layout fixes

**Current working state**:
- Modified: `FormPagoFactura.frm`, `FormPresupuesto.frm`, `FormVerPresupuestos.frm`, `FormVerPagoFactura.frm`, `DB_SPC_SI.mdb`
- Untracked: `OrdenPago_Pato-01.jpeg`, `ReciboNoOficial.jpeg` (sample images for testing)

---

## Quick Start for New Developers

1. **Understand the domain**: Read `README.md` (Spanish project overview), `WARP.md` (technical guidance)
2. **Set up Git hooks**: Run `tools/install-hooks.ps1` to prevent accidental binary commits
3. **Explore the database**: Open `DB_SPC_SI.mdb` in Microsoft Access to understand schema (Clientes, Articulos, Facturas, Pagos, etc.)
4. **Study a key module**: Open `FormFactura.frm` in VB6 IDE to understand UI + business logic pattern
5. **Review global state**: Check `VariablesPublicas.bas` to see shared recordsets and config variables
6. **Open project in VB6 IDE**: Load `SPCSI.vbp` to navigate forms, modules, references
7. **Compile & run**: Build from IDE (File → Make SPCSI_5.exe), run executable, test manually
8. **Consult shared SQL**: Review `Sentencias.bas` for common database operations

---

## Key Files to Know

| File | Purpose |
|------|---------|
| `MenuPrincipal.frm` | Main menu; entry point to all application features |
| `MDIForm1.frm` | MDI container; manages child windows |
| `VariablesPublicas.bas` | Global variables, recordsets, config (shared state hub) |
| `Sentencias.bas` | Shared SQL queries and database operations |
| `FormFactura.frm` | Core invoicing workflow |
| `FormPagoFactura.frm` | Payment registration and tracking |
| `FormLibroIvaVentas.frm` | Tax compliance report (AFIP-mandated) |
| `Clientes.frm` | Customer master data management |
| `Articulos.frm` | Product/article catalog management |
| `DB_SPC_SI.mdb` | Main database (all persistent data) |

---

## Final Notes

- **Monolithic codebase**: Logic lives mostly in forms; expect tight coupling between UI and business logic
- **Manual testing**: No automated test suite; rely on manual testing in IDE and executable
- **Database-first**: Schema changes directly in Access; coordinate with form/module updates
- **Multi-form architecture**: Many forms (100+) for different workflows; navigation is menu-driven
- **Production-ready**: Used for real business operations; changes should be tested thoroughly before commit
- **Fiscal compliance**: AFIP-integrated for Argentine electronic invoicing; respect tax document workflows

This is a mature VB6 application with strong domain-specific focus. Respect existing patterns, maintain CRLF line endings, test manually, and coordinate form + database changes together.
