# Manual de Usuario — Automatizador de Inventarios SI ESAM

Guía rápida para convertir cualquier inventario (Excel, CSV o TXT) al formato oficial de SI ESAM, sin necesidad de tocar código.

---

## 1. Abrir la aplicación

Doble clic en **`lanzar.bat`**.

Se abre la ventana **"CONVERTIDOR DE INVENTARIOS PARA SI ESAM"**, dividida en dos paneles:

- **Izquierda — "Entrada de datos"**: aquí cargas el inventario.
- **Derecha — "Opciones de conversión"**: aquí ajustas opciones, conviertes y ves el registro de actividad.

---

## 2. Cargar el inventario

1. Clic en **"Examinar…"**.
2. Selecciona tu archivo: Excel (`.xlsx`/`.xls`), CSV o TXT.
3. El nombre del archivo aparece en el cuadro junto al botón. Si te equivocaste, usa el botón **"✗"** para quitarlo.

Si tu archivo es TXT, cada línea representa un producto, en alguno de estos formatos:
```
LAVADORA 10KG PANASONIC - Bs 6500
REFRIGERADOR 2P SHARP: 5200
MICROONDAS TOSHIBA 30L; 850
```
> Nota: con un archivo TXT, la app no puede detectar cantidad ni costo por separado — solo nombre y precio.

---

## 3. Revisar las opciones de conversión

En el panel derecho, arriba de todo:

| Campo | Qué significa | Valor por defecto |
|-------|----------------|---------------------|
| **Código inicial** | Desde qué número empieza a generar códigos automáticos para productos sin código propio | 10001 |
| **Sede ID** | El ID de la sede/sucursal en el sistema SI ESAM | 1 |
| **Unidad ID** | La unidad de medida por defecto (57 = UNIDAD BIENES) | 57 |

Normalmente **no hace falta cambiar nada** — solo ajusta el "Código inicial" si ya tienes productos cargados en el sistema y no quieres que se choquen los códigos.

La casilla **"Previsualizar productos antes de guardar"** debe quedar marcada (viene así por defecto): te deja revisar los productos antes de generar el Excel final.

### ¿Cómo sé qué columnas detecta la app en mi Excel?
Clic en **"ℹ Ver columnas detectables y Unidades"**. Se abre una ventana con:
- Pestaña **"Detección de columnas"**: qué palabras debe tener el encabezado de cada columna para que la app la reconozca (ej. una columna "Precio Venta" se detecta como precio).
- Pestaña **"Unidades"**: la lista fija de 124 unidades del SIN Bolivia.

---

## 4. Convertir

1. Clic en el botón grande **"⚡ CONVERTIR INVENTARIO"**.
2. Revisa el **"Registro"** (recuadro oscuro inferior derecho): muestra qué columnas detectó y cuántos productos encontró.
3. Si la previsualización está activada, aparece una ventana con la lista de productos detectados (código, descripción, precio, marca, categoría, cantidad). Revisa que todo esté correcto y confirma.
4. Se abrirá un cuadro para **elegir dónde guardar** el archivo Excel final — elige la carpeta y nombre que prefieras.

Al terminar, el registro muestra un resumen: cuántas categorías, marcas y ubicaciones se generaron, y confirma el guardado.

---

## 5. Qué contiene el Excel de salida

El archivo generado trae 8 hojas, listas para importar al sistema SI ESAM:

| Hoja | Contenido |
|------|-----------|
| CLIENTE PROVEEDORES | Vacía (plantilla) |
| PRODUCTOS | Un producto por fila |
| PRODUCTOS DETALLES | Cantidad y costo por sede |
| PRODUCTO SEDES | Precio, costo y utilidad calculada por sede |
| UNIDADES | Lista fija del SIN Bolivia |
| CATEGORIA | Categorías detectadas + "GENERAL" |
| MARCA | Marcas detectadas + "GENERAL" |
| UBICACIONES | Ubicaciones detectadas + "GENERAL" |

**Sobre Categoría, Marca y Ubicación:** si tu Excel de entrada no trae esas columnas, la app le asigna "GENERAL" a todos los productos (no falla, simplemente no las separa).

**Sobre la utilidad:** se calcula sola como `((precio de venta / costo) - 1) × 100`. Si el producto no tiene costo cargado, la utilidad queda en 100% por defecto.

**Sobre la cantidad:** si un producto no trae cantidad en el archivo de origen, queda en **0** (no se asume 1).

---

## 6. Problemas comunes

| Problema | Qué revisar |
|----------|-------------|
| "No se encontraron productos" | Verifica que el archivo tenga una columna con nombre reconocible (ver "ℹ Ver columnas detectables"), o que el archivo TXT tenga el formato correcto |
| El precio o costo salió mal / la utilidad da 0 | Es probable que dos columnas de tu Excel compartan una palabra clave (ej. "Precio Compra" y "Precio Venta" ambas contienen "precio"). Revisa el Registro: muestra exactamente qué columna detectó como `precio=` y `costo=` |
| La app no abre / error de dependencias | Sigue las instrucciones de instalación en `README.md` |

---

Para detalles técnicos (instalación, personalización de categorías/marcas conocidas, estructura del código), ver `README.md`.
