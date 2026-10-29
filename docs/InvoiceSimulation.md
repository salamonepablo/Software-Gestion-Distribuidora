# Invoice simulation (facturación de prueba sin ARCA)

Use this mode to issue test invoices from the VB6 IDE without calling ARCA/AFIP.
It works like the mode that already exists in SPC-Core.

> **It writes to the local database.** Stock, current account (cuenta corriente),
> the invoice and its items are saved exactly as in a real sale. Only the ARCA
> call is replaced. Use it only with a local copy of `DB_SPC_SI.mdb`.

## Where it works

The mode is available only when **both** of these are true:

| Condition | Check |
|-----------|-------|
| Running from the VB6 IDE | `Debug.Assert` detection in `CanSimulateInvoices` (compiled `.exe` builds drop it) |
| Project folder is exactly `C:\Trabajos Activos\SPC-Retail` | `App.Path`, case-insensitive, no subfolders |

The compiled executable and any other folder (including SPC-Core or an installed
copy) always use ARCA. Pressing the shortcut there shows
*"El modo de prueba requiere ejecutar este proyecto desde el IDE en C:\Trabajos Activos\SPC-Retail."*

## Turn it on

1. Run the project from the IDE in `C:\Trabajos Activos\SPC-Retail`.
2. Open **Facturas**. The button **Activar facturacion de prueba** appears in the bottom
   action bar, right of **Salir**.
   If it is hidden, press **Ctrl+Shift+F12**.
3. Click the button and read the confirmation. **No** is the default; choose **Sí** to continue.
4. The form title changes to **FACTURAS - PRUEBA - SIN VALIDEZ FISCAL** and the button
   is disabled with the same text.

## What happens while it is on

- `FacturaElectronicaSPC` returns before any ARCA code (no login, certificate or web service call).
- Invoices and credit notes A and B (ARCA types 1, 6, 3, 8) get CAE `99999999999999`
  and CAE expiry today + 10 days, written with the existing `ActualizarCAE`.
- Do not issue debit notes in this mode. They are rejected with
  *"El modo de prueba no autoriza notas de debito: solo facturas y notas de credito A y B.
  El comprobante local puede haber quedado grabado sin CAE."*
- Any other unsupported type is rejected with the same message, as is a document
  that is not in `FacturaC`/`NotaCreditoC`. **Facturas** and **Notas de Crédito**
  then revert the document, just as they do when ARCA rejects it.
  **Notas de Débito** ignores the result, so a debit note issued in this mode
  stays saved without a CAE.
- The mode applies to the whole session, including **Notas de Crédito**.
- Printed and PDF invoices and credit notes show **PRUEBA - SIN VALIDEZ FISCAL**
  instead of the CAE line whenever the stored CAE is `99999999999999`.

## Turn it off

Stop the project in the IDE and run it again. Nothing turns it off during a session.

## Tests

```powershell
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\tests\Test-InvoiceSimulation.ps1
```

The test runs the real procedures in 32-bit ScriptControl with mocked `App`,
`Debug.Assert`, `MsgBox`, ARCA and an in-memory database. It does not compile
the project or touch `DB_SPC_SI.mdb`.
