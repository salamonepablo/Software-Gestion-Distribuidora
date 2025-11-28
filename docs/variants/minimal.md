# SPC Variantes (Core, Retail, Minimal)

Este documento describe las diferencias entre las tres variantes del sistema SPC desarrollado en VB6 + Access.

## Objetivo de cada variante

- Core (main): Base general del producto. Contiene módulos compartidos y funcionalidades completas destinadas a entornos estándar.
- Retail (spc-retail): Orientada a punto de venta (local al publico) con todas las pantallas y lógica comercial amplia (facturación completa, remitos, comisiones, reportes IVA detallados).
- Minimal (spc-minimal): Versión reducida para clientes con necesidades básicas (sin facturación oficial, emisión limitada de comprobantes, gestión simplificada de cuentas corrientes, menor cantidad de pantallas y reportes).

## Estado de las ramas

| Rama         | Uso principal                | PR Draft contra main | Merge esperado | Tags iniciales sugeridos           |
|--------------|------------------------------|----------------------|----------------|------------------------------------|
| main         | Core (base)                  | N/A                  | N/A            | `v1.0.0-SPC-core`                  |
| spc-retail   | Variante Retail completa     | Sí (NO MERGE)        | No             | `v1.0.0-SPC-retail`                |
| spc-minimal  | Variante Minimal básica      | Sí (NO MERGE)        | No             | `v1.0.0-SPC-minimal`               |

Las PRs Draft se usan solo para revisión y discusión. No deben mergearse a `main`.

## Diferencias funcionales (ejemplo inicial)

| Módulo / Función                  | Core | Retail | Minimal | Notas |
|----------------------------------|:----:|:------:|:------:|-------|
| Facturación completa (A/B/NC/ND) | Sí   | Sí     | Limitado (solo tipos básicos) | En Minimal se excluyen variantes avanzadas. |
| Gestión Clientes avanzada        | Sí   | Sí     | Reducida | Menos formularios de búsqueda y listados. |
| Reporte Libro IVA Ventas         | Sí   | Sí     | No      | Puede reemplazarse por exportación simple. |
| Cálculo Comisiones               | Sí   | Sí     | No      | Retail lo usa; Minimal no. |
| Manejo Depósitos múltiples       | Sí   | Sí     | No      | Eliminado para simplificar. |
| Panel Control extendido          | Sí   | Sí     | No      | Se elimina UI compleja en Minimal. |
| Limpieza / Herramientas auxilia  | Sí   | Sí     | No      | Puede ofrecerse como herramienta externa. |
| Cuenta Corriente completa        | Sí   | Sí     | Parcial | Minimal conserva lo básico de saldos y movimientos. |
| Generación códigos barras / QRs  | Sí   | Sí     | Parcial | Requiere revisión de qué formatos quedan en Minimal. |

(Ajusta esta tabla según lo que realmente esté en cada rama.)

## Estrategia de sincronización

1. Correcciones generales (bugs comunes):
   - Se aplican en `main`.
   - Se propagan a `spc-retail` y `spc-minimal` mediante:
     ```
     git checkout spc-retail
     git fetch origin
     git merge origin/main
     # Resolver (aceptar changes de main solo en archivos compartidos)
     git push
     ```
     (Igual para `spc-minimal`.)

2. Cambios específicos de una variante NO se llevan automáticamente a `main` para evitar mezclar lógica exclusiva.

3. Si una funcionalidad de variante pasa a ser común:
   - Se extrae el código hacia `main`.
   - Se limpia duplicación en cada variante.

## Recomendación de estructura futura (opcional)

Podrías unificar en un solo árbol con subcarpetas:
```
/src/core
/src/retail
/src/minimal
/docs
```
Y configurar un “modo” por archivo de configuración (ej: `config/variant.ini`) para habilitar/ocultar formularios y módulos. Esto reduciría mantenimiento duplicado.

## Base de datos (Access)

- Si las variantes comparten el esquema principal, mantener un script de migración (documentar campos) en `docs/db-schema.txt`.
- Evitar subir bases con datos reales. Incluir solo:
  - `DB_SPC_SI.mdb` vacía (estructura).
  - Scripts de carga de datos de prueba (anonimizados) opcionales.

## Política de PRs

- PRs Draft con título: `[VARIANTE] SPC Retail – NO MERGE`.
- Etiquetas sugeridas: `variant`, `no-merge`, `review`.
- Descripción debe incluir la razón del NO MERGE.
- Revisión se centra en estándares de estilo, seguridad y portabilidad.

## Tags / Versionado

- Inicial: `v1.0.0-SPC-core`, `v1.0.0-SPC-retail`, `v1.0.0-SPC-minimal`.
- Incrementos:
  - Patch (x.y.z+1): correcciones pequeñas.
  - Minor (x.y+1.0): nuevas pantallas / mejoras internas sin romper compatibilidad.
  - Major (x+1.0.0): cambios estructurales o ruptura de compatibilidad.

Ejemplo de creación de tag:
```
git checkout spc-retail
git tag v1.0.0-SPC-retail
git push origin v1.0.0-SPC-retail
```

## Pendientes próximos (Checklist)

- [ ] Completar tabla de diferencias reales.
- [ ] Agregar etiquetas en GitHub.
- [ ] Revisar si hay datos sensibles en commits iniciales.
- [ ] Documentar script de back-up / restore de Access.
- [ ] Evaluar paso a repositorio monorepo con carpetas por variante (opcional).

## Notas finales

Este documento debe actualizarse cada vez que una función se agrega, mueve o elimina en alguna variante para mantener claridad de mantenimiento.