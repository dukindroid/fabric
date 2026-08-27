# Active Context

## Estado de inicialización

La memoria del proyecto fue inicializada el 14 de agosto de 2026 a partir de la estructura y configuración presentes en el checkout.

## Foco actual

El objetivo inmediato es la implementación la interfaz de usuario necesaria en `client/src/components/pos/payment-dialog.tsx` para permitirnos capturar en cada venta la información pertinente requerida por la gerente al momento del corte de caja. Para esto planeo bosquejar el flujo de opciones de la interfaz de usuario mediante un bosquejo en figma. Mi intención es que sea una interfaz limpia y con pocas opciones, mostradas una a una mediante animaciónes tipo acordión según se vaya avanzando en el proceso de finalización de una venta.

Esta interfaz nos permitirá poblar los datos extras necesarios en `orders`. Considero que es posible personalizar el campo `source` para que las opciones incluidas en este campo sean `["mostrador", "uber", "didi", "rappi", "entrega", "transferencia", "interno"]`. Se requerirá de un campo extra para definir si esta orden se cobra con o sin IVA, dependiendo de la necesidad de facturación y otro campo más que nos especifique si el pago fue mediante efectivo o tarjeta.

### Evaluación del diálogo actual

- `PaymentDialog` actualmente mezcla origen de venta y pago en el estado `method`.
- Los botones “Entrega” y “Transferencia” usan ambos el valor `receipt_only`, por lo que hoy no se pueden distinguir al persistir la orden.
- `cash_with_discount` representa simultáneamente una venta de plataforma y la posibilidad de editar precios; no debe convertirse automáticamente en el contrato final de origen/pago.
- El callback actual solo transmite `method`, `cashReceived`, `change` y `updatedCart`; para el nuevo flujo necesitará un contrato explícito que separe origen, método de pago, fiscalidad y datos adicionales.
- `orders.source` todavía está restringido en `shared/schema.ts` a `pos | whatsapp | web`; las categorías de negocio propuestas requieren migración, compatibilidad y actualización de tipos/API.
- La implementación actual de precios modificados por plataforma ya se propaga desde `payment-dialog.tsx` hacia `pos.tsx` y se guarda dentro de los items serializados del pedido mediante `originalPrice`/`discountedPrice`.
- El componente padre es `client/src/pages/pos.tsx`; no conviene cambiar únicamente la interfaz sin coordinar el callback, la mutación de creación de órdenes, recibos, reportes y esquema.

A la tabla `cashOuts` se le deberán agregar los campos necesarios para almacenar las cantidades de moneda existentes y algún otro dato adicional que de momento no recuerdo. Como referencia tenemos el archivo en `memory-bank/HOJA DE CORTE.XLSX`. También deberémos de diseñar un layout que nos permita mostrar el reporte del corte de caja de la manera más limpia y concisa, con opción a hacer click en algún ícono de `+` para presentar un reporte mas completo de ventas diarias.

## Cambios preexistentes detectados

`git status` mostró cambios que no fueron realizados por esta inicialización:

- `.env.example` eliminado.
- `client/src/components/pos/payment-dialog.tsx` modificado.
- `electron/start-dev-optimized.cjs` modificado.
- `package-lock.json` modificado.
- `package.json` modificado.
- `vite.config.ts` modificado.
- `.clinerules/` sin seguimiento.

No modificar ni revertir esos cambios sin revisar primero su intención.

## Decisiones y patrones a respetar

- Usar alias `@/*` para `client/src/*` y `@shared/*` para `shared/*`.
- Mantener tipos y esquema compartidos en `shared/`.
- Seguir React Query para fetching/cache e invalidación.
- Mantener los límites actuales entre cliente, API Express y proceso Electron.
- Antes de implementar una feature, revisar rutas existentes y componentes relacionados para evitar duplicación.

## Próximos pasos sugeridos

1. Confirmar el estado de los cambios preexistentes antes de trabajar sobre ellos.
2. Ejecutar `npm run check` y builds relevantes para establecer una línea base.
3. Revisar inconsistencias de branding entre Claudette y OpenSauce.
4. Mantener este archivo actualizado después de cada cambio significativo.

## Validación realizada en esta sesión

- Se releídos los seis archivos del banco de memoria.
- Se revisaron `payment-dialog.tsx`, su padre `pos.tsx`, `shared/schema.ts`, el contexto de moneda y el componente base de diálogo.
- Se confirmó que `payment-dialog.tsx` tiene un cambio local preexistente de traducción del título (“Process Payment” → “Proceso de pago”); no se revirtió.
- Se inició `npm run check`; el resultado final debe confirmarse antes de implementar cambios estructurales.

## Tareas a no perder de vista en un futuro

- Revisión de problemas encontrados por el reporte de testeo de la aplicación en `
tests/reports/`
- Impresión de formatos de conteo físico de inventario, e interfaz para su debida entrada al sistema una vez realizado. Posibilidad de dejar notas en el conteo de inventario y modificaciones según lo encontrado
- Interfaz de planeación de producción, en la cual se presente un listado del producto próximo a requerir elaboración, con fecha límite de entrega
- Interfaz de listado de entregas pendientes de pago, así como de transferencias bancarias pendientes de entrega de producto, con su respectivo detalle correspondiente
