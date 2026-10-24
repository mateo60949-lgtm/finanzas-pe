# Activar la lectura automática de correos

Unos 10 minutos, una sola vez. Todo corre en tu cuenta de Google; nadie más ve tus datos.

## 1. Crear la hoja y pegar el script
1. Abre **sheets.new** (crea una hoja nueva) y ponle de nombre `finanzas pe`.
2. Menú **Extensiones → Apps Script**.
3. Borra lo que aparece en `Código.gs` y pega todo el contenido de `Codigo.gs` (este mismo folder).
4. Guarda (ícono del disquete).

## 2. Instalar
1. Arriba, en la lista de funciones, elige **`instalar`** y toca **Ejecutar**.
2. Google pedirá permisos: *Revisar permisos* → elige tu cuenta → *Configuración avanzada* → *Ir a finanzas pe (no seguro)* → **Permitir**.
   Aparece "no seguro" porque el script lo escribiste tú y Google no lo ha revisado; es normal para scripts personales.
3. Al terminar, la hoja tiene las pestañas **Movimientos**, **Revisar** y **Respaldo**, y el script queda leyendo tus correos cada 10 minutos.

## 3. Publicarlo para que la app lo lea
1. **Implementar → Nueva implementación** → engranaje → **Aplicación web**.
2. *Ejecutar como*: **Yo**. *Quién tiene acceso*: **Cualquier usuario**.
3. **Implementar**.
4. Vuelve al editor, elige la función **`verEnlace`** y **Ejecutar**. En el registro aparece el enlace que termina en `/exec?t=…`, o las instrucciones para armarlo con la URL de *Gestionar implementaciones*. Un enlace que termine en `/dev` no sirve: es el de prueba y solo funciona con tu sesión de Google.

"Cualquier usuario" no expone nada: sin el código `t=…` del enlace el script no entrega datos. No compartas ese enlace.

## 4. Conectar la app
En la app: **Config → Sincronización con tus bancos** → pega el enlace → **Conectar**.
Los movimientos aparecen en **Inicio → Por revisar**. La primera vez eliges a qué cuenta o tarjeta de la app corresponde cada cuenta del banco; desde ahí lo recuerda y puedes usar **Confirmar todos**.

## Mantenimiento
- **Si el enlace se compartió por error**: ejecuta `cambiarToken`. El enlace anterior deja de funcionar al instante; copia el nuevo del registro (o arma `…/exec?t=` + el token nuevo) y pégalo en la app: *Config → Sincronización → Conectar*.
- **Precios de inversiones**: el script crea la pestaña **Precios** con fórmulas GOOGLEFINANCE para cada ticker que registres en la app. Si un ticker no aparece (por ejemplo un ETF europeo), en la app toca la posición y pon el precio a mano, o cambia el ticker en la hoja por el código con bolsa (ej: `ETR:VWCE`).
- **Correos que no entendió**: quedan en la pestaña **Revisar** de la hoja, nunca se pierden. Pásame algunos y agrego la regla.
- **Si cambias el script**: después de pegar la versión nueva, ejecuta `releerTodo` para volver a leer con las reglas nuevas (los movimientos ya confirmados en la app no se duplican), y en *Implementar → Gestionar implementaciones* edita la implementación y elige **Nueva versión**.
- **Bank of America**: activa en su app *Alerts → Email* para compras, depósitos y transferencias. Sus reglas aún no se han probado con correos reales; los primeros aparecerán marcados para revisar.
