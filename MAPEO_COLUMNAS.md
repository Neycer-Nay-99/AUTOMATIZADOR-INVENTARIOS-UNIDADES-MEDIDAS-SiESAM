# Mapeo de columnas — qué detecta el sistema y dónde lo coloca

Esta guía explica, campo por campo, cómo el Automatizador de Inventarios lee el Excel/CSV de entrada y en qué hoja/columna del Excel de salida (formato oficial SI ESAM) termina cada dato.

> Para ver esto mismo pero interactivo (con los nombres reales de columna de TU archivo), abrí la app, seleccioná el archivo y mirá el panel **"Columnas detectadas en este archivo"**, o el botón **"ℹ Ver columnas detectables"**.

---

## Cómo detecta las columnas

La app **no depende del orden de las columnas ni de mayúsculas/acentos**. Busca, en el encabezado de cada columna, si contiene alguna de las palabras clave de la tabla de abajo (coincidencia parcial, sin distinguir mayúsculas). Por ejemplo, para "Nombre" da lo mismo que la columna se llame `Nombre`, `NOMBRE DEL PRODUCTO` o `Producto`.

Si dos columnas del origen matchean la misma palabra clave (ej. dos columnas que contienen "precio"), la app las distingue por orden de aparición (ver la sección **Precio vs. Costo vs. Precio alterno** más abajo).

---

## Tabla de mapeo

| Campo detectado | Palabras clave (parcial, sin distinguir mayúsculas) | Va a la hoja... | ...columna | Notas |
|---|---|---|---|---|
| **Código/SKU** | `codigo`, `código`, `cod`, `sku`, `item_code`, `product_code`, `code`, `codebar`, `barcode` | PRODUCTOS | `codigo` | Si no hay columna de código, se genera uno incremental (código inicial configurable en la GUI) |
| **Código/ítem (2da)** | (misma lista que Código, 2da columna que matchea) | PRODUCTOS | `codigo_item` | Solo si el origen trae DOS columnas de código (ej. código interno + código de barras) |
| **Nombre/Descripción** | `nombre`, `producto`, `descripcion`, `item`, `articulo`, `name`, `product` | PRODUCTOS | `descripcion` | Se pasa a MAYÚSCULAS. Si no se detecta, se usa la primera columna de texto del archivo |
| **Descripción corta** | `descripcion corta`, `desc corta`, `descripcion breve`, `short description` | PRODUCTOS | `descripcion_corta` | Se busca antes que "Nombre" para no confundirse con ella |
| **Precio** | `precio`, `price`, `valor`, `pvp`, `tarifa`, `importe`, `monto`, `venta` | PRODUCTO SEDES | `precio_unitario` | Precio de venta del producto base |
| **Precio alterno** | (misma lista que Precio, 2da columna que matchea) | PRODUCTOS_UNIDADES_MEDIDAS | `precio_unitario` | Precio de las presentaciones alternas (caja, docena, etc.) — ver sección abajo |
| **Costo** | `costo`, `cost`, `compra` | PRODUCTOS DETALLES y PRODUCTO SEDES | `costo` (ambas hojas) | También se usa para calcular `utilidad` en PRODUCTO SEDES |
| **Marca** | `marca`, `brand`, `fabricante` | MARCAS (nueva fila si no existe) + PRODUCTOS | `marca_id` | Si no hay columna, todos los productos quedan con marca `GENERAL` (id=1) |
| **Categoría** | `categoria`, `category`, `tipo`, `grupo`, `familia`, `family` | CATEGORIAS (nueva fila) + PRODUCTOS | `categoria_id` | Si no hay columna, categoría `GENERAL` (id=1) |
| **Subcategoría** | `subcategoria`, `sub categoria`, `sub-categoria`, `subcategory` | CATEGORIAS (nueva fila, hija de Categoría) + PRODUCTOS | `categoria_id` | Si existe, el producto queda vinculado a la **subcategoría** (no a la categoría padre) — ver sección abajo |
| **Cantidad** | `cantidad`, `qty`, `stock`, `existencia`, `quantity`, `unidades` | PRODUCTOS DETALLES | `cantidad` | Si no se detecta, queda en `0` (no se asume 1) |
| **Existencia mínima** | `existencia minima`, `stock minimo`, `cantidad minima`, `minimo` | PRODUCTOS | `existencia_minima` | |
| **Existencia máxima** | `existencia maxima`, `stock maximo`, `cantidad maxima`, `maximo` | PRODUCTOS | `existencia_maxima` | |
| **Ubicación** | `ubicacion`, `location`, `sede`, `almacen`, `bodega`, `deposito` | UBICACION (nueva fila) + PRODUCTOS | `ubicacion_id` | Si no hay columna, ubicación `GENERAL` (id=1) |
| **Unidad (presentación)** | `unidad`, `presentacion`, `empaque` | PRODUCTOS (unidad base) y/o UNIDAD_MEDIDA (presentaciones alternas) | `unidad_id` / `descripcion` | Se resuelve por palabra clave al catálogo fijo de 122 unidades del SIN (ver sección abajo) |
| **Factor** | `factor`, `conversion`, `equivalencia`, `multiplicador` | UNIDAD_MEDIDA | `factor` | Solo tiene efecto junto con la columna Unidad (ver sección abajo) |
| **Lote** | `lote`, `lot`, `batch` | PRODUCTOS DETALLES | `lote` | Se copia tal cual viene |
| **Vencimiento** | `vencimiento`, `caducidad`, `expiracion`, `expiry`, `expiration` | PRODUCTOS DETALLES | `vencimiento` | Se normaliza siempre a `AAAA-MM-DD` (ver sección abajo) |
| **Principio Activo** | `principio activo`, `active ingredient` | PRODUCTOS | `principio_activo` | Se copia tal cual viene (uso típico: farmacia) |
| **Registro Sanitario** | `registro sanitario`, `reg sanitario` | PRODUCTOS | `registro_sanitario` | Se copia tal cual viene (uso típico: farmacia) |

---

## Precio vs. Costo vs. Precio alterno

Algunos proveedores (ej. DUNAMIS) mandan **dos columnas que contienen la palabra "precio"**: una para el producto principal y otra, más adelante en el archivo, para las filas de presentación alterna (caja, docena, etc.). La app las distingue por orden:

1. La **primera** columna que matchea "precio" (y no es la de Costo) → precio del producto base → `PRODUCTO SEDES.precio_unitario`.
2. La **segunda** (si existe) → precio de cada presentación alterna → `PRODUCTOS_UNIDADES_MEDIDAS.precio_unitario`.
3. Si solo hay una columna de precio, se usa para ambos casos (sin cambio de comportamiento en archivos simples).

El **Costo**, aunque también contenga la palabra "precio" en algunos casos (ej. "Precio Compra"), se detecta y excluye por separado usando sus propias palabras clave (`costo`, `cost`, `compra`), así que nunca se confunde con el precio de venta.

**Utilidad** (`PRODUCTO SEDES.utilidad`) se calcula automáticamente: `((precio_venta / costo) - 1) * 100`. Si hay precio pero no costo, la utilidad queda en `100`. Si hay costo pero no precio, queda en `0`.

---

## Categoría y Subcategoría (jerarquía de 2 niveles)

Si el archivo trae **ambas** columnas (Categoría y Subcategoría), el sistema arma una jerarquía de 2 niveles en la hoja CATEGORIAS:

1. Se crea (o reutiliza si ya existe) la fila de la **Categoría** (padre), sin `categoria_padre_id`.
2. Se crea (o reutiliza) la fila de la **Subcategoría** (hija), con `categoria_padre_id` apuntando a la categoría padre.
3. El producto queda vinculado (`PRODUCTOS.categoria_id`) a la **subcategoría** (el nivel más específico), no a la categoría padre.

**Ejemplo:** un producto con Categoría=`LUBRICANTES` y Subcategoría=`ACEITES HIDRAULICOS` genera:

| id | descripcion | categoria_padre_id |
|---|---|---|
| 1 | GENERAL | (vacío) |
| 2 | LUBRICANTES | (vacío) — es la categoría padre |
| 3 | ACEITES HIDRAULICOS | 2 — hija de LUBRICANTES |

Y el producto queda con `categoria_id = 3`.

Si el archivo **no trae Subcategoría**, se comporta igual que antes: el producto queda vinculado directamente a la Categoría (que queda sin padre).

La deduplicación de nombres es **global** (por nombre, sin importar bajo qué padre está): si el mismo nombre de subcategoría aparece bajo dos padres distintos, se reutiliza la primera fila creada.

---

## Unidad y Factor (presentaciones alternas: caja, docena, etc.)

Este es el mecanismo más particular del sistema. Se activa cuando el archivo trae **ambas** columnas Unidad y Factor, y sigue el patrón de "fila base + filas de continuación" que usan proveedores como DUNAMIS:

- Una fila **con nombre de producto** = un producto nuevo. Su columna Unidad/texto se resuelve como la **unidad base** del producto (`PRODUCTOS.unidad_id`).
- Una fila **sin nombre de producto**, pero con datos en Unidad y/o Factor = una **presentación alterna** del producto de la fila anterior (ej. "CAJA", "DOCENA", "MEDIA CAJA").

Cada presentación alterna genera:
1. Una fila en **UNIDAD_MEDIDA**, con `descripcion` = texto de origen + factor (ej. `"CAJA X20"`, `"DOCENA X12"`) y su propio `factor`.
2. Una fila en **PRODUCTOS_UNIDADES_MEDIDAS** (tabla puente producto↔presentación), con el precio de esa presentación (columna "Precio alterno", ver arriba) y el `orden` en que aparece.

**Ejemplo real** (proveedor DUNAMIS, producto "ACUARELA BOLSA" con 5 presentaciones):

| Unidad (origen) | Factor | Precio | → UNIDAD_MEDIDA.descripcion |
|---|---|---|---|
| CUARTA | 3 | 20 | CUARTA X3 |
| MEDIA DOCENA | 6 | 40 | MEDIA DOCENA X6 |
| DOCENA | 12 | 80 | DOCENA X12 |
| MEDIA CAJA | 144 | 900 | MEDIA CAJA X144 |
| CAJA (24DOC) | 288 | 1800 | CAJA (24DOC) X288 |

**Resolución de la unidad a un ID fijo:** el texto libre de Unidad (ej. "CAJA (24DOC)", "MEDIA DOCENA") se traduce al catálogo fijo de 122 unidades del SIN Bolivia por palabra clave:

| Si el texto contiene... | Se resuelve a la unidad (ID) |
|---|---|
| CUARTA, DOCENA (o MEDIA DOCENA) | DOCENA (14) |
| CAJA (o MEDIA CAJA) | CAJA (6) |
| GRUESA | GRUESA (18) |
| PAQ / PAQUETE | PAQUETE (42) |
| FARDO | FARDO (15) |
| BOLSA | BOLSA (4) |
| MILLAR(ES) | MILLAR (38) |
| BULTO | BULTO (66) |
| DISPLAY | DISPLAY (65) |
| CARTON(ES) | CARTON (7) |
| SET / JUEGO | SET (21) |
| UNIDAD | UNIDAD BIENES (57) |
| (ninguno de los anteriores) | OTRO (62) — respaldo |

Si el archivo **no trae** columnas de Unidad/Factor, el sistema se comporta como antes: no se generan filas en UNIDAD_MEDIDA/PRODUCTOS_UNIDADES_MEDIDAS, y `PRODUCTOS.unidad_id` usa el valor "Unidad ID" configurado en la GUI (57 = UNIDAD BIENES por defecto) para todos los productos.

---

## Vencimiento: normalización de fecha

Sin importar cómo venga la fecha en el origen, siempre se exporta como **`AAAA-MM-DD`** (ej. `2028-09-12`). Formatos de entrada aceptados:

- Texto `DD-MM-AAAA`, `DD/MM/AAAA`, `DD.MM.AAAA` (ej. `12/9/2028`)
- Texto ya en `AAAA-MM-DD` o `AAAA/MM/DD`
- Fecha real de Excel (la celda tiene formato de fecha) — se lee automáticamente
- **Número de serie de Excel** (ej. `46609`), que ocurre cuando Excel/pandas no reconoce el formato de la celda como fecha — se convierte igual, calculando la fecha real a partir de ese número
- Celda vacía → queda vacío (sin error)

Si el valor no matchea ningún formato conocido (texto no reconocible), se deja tal cual vino y se registra un aviso en el log de la app — no se pierde el dato, pero conviene revisarlo.

---

## Campos que SIEMPRE van vacíos o fijos (no se leen del origen)

Estos campos existen en el Excel de salida porque el ERP los espera, pero el sistema no los completa desde ningún dato de entrada — quedan vacíos o con un valor fijo a propósito:

- `slug`, `created_at`, `updated_at` (en PRODUCTOS, CATEGORIAS, MARCAS, PRODUCTOS DETALLES, UNIDAD_MEDIDA)
- `imagen`, `codigo_actividad`, `codigo_linea`, `descripcion_larga`, `codigo_sin`, `tipo`, `codigo_volvo`, `config_adicionales`, `medidas` (en PRODUCTOS)
- `estado = 1`, `es_servicio = 0`, `novedoso = 0`, `es_bonificado = 0`, `venta_controlada = 0` (en PRODUCTOS)
- `utilidad2`, `utilidad3`, `utilidad4`, `precio_factura`, `utilidad_factura`, `aplicar_lista_precios`, `calcular_cantidad_por_precio`, `mostrar_en_ecommerce`, `es_alquilable`, `tarifa_alquiler` (en PRODUCTO SEDES)
- `UNIDADES` es siempre la lista fija de 122 unidades del SIN Bolivia — no se genera desde el archivo de entrada, ya viene completa en el sistema.
- Hoja `CLIENTE PROVEEDORES` — siempre vacía (plantilla).

---

Para instrucciones de instalación y uso general de la app, ver **`INSTALACION.md`** y **`MANUAL_USUARIO.md`**.
