# Automatizador de Inventarios — SI ESAM v2.0

Convierte cualquier inventario (Excel o CSV) al formato oficial de SI ESAM.  
Las hojas **CATEGORIA**, **MARCA** y **UBICACIONES** se construyen automáticamente desde los datos de entrada.  
La hoja **UNIDADES** es siempre la lista fija oficial del SIN Bolivia (124 unidades).

> 📖 **[Guía de instalación paso a paso](INSTALACION.md)** — para instalar en una computadora nueva, desde cero.
> 📘 **[Manual de usuario](MANUAL_USUARIO.md)** — cómo usar la aplicación una vez instalada.

---

## Instalación rápida

### 1. Crear entorno virtual e instalar dependencias

```powershell
# Abrir PowerShell en la carpeta del proyecto
python -m venv venv
.\venv\Scripts\activate
pip install pandas openpyxl
```

### 2. Ejecutar la aplicación

**Doble clic** en `lanzar.bat`  
_(o desde PowerShell)_:

```powershell
python main.py
```

---

## Cómo usar

| Paso | Acción |
|------|--------|
| 1 | Seleccionar un archivo Excel / CSV |
| 2 | Ajustar opciones: código inicial, sede ID, unidad ID |
| 3 | Clic en **CONVERTIR INVENTARIO** |
| 4 | Revisar la previsualización y confirmar |
| 5 | Elegir dónde guardar el Excel resultante |

---

## Detección automática de columnas en Excel/CSV

La aplicación busca columnas por nombre usando estas palabras clave:

| Campo | Palabras clave detectadas |
|-------|--------------------------|
| Nombre del producto | nombre, producto, descripcion, item, articulo, name |
| Precio de venta | precio, price, valor, pvp, tarifa, importe, monto, venta |
| Costo | costo, cost, compra |
| Marca | marca, brand, fabricante |
| Categoría | categoria, category, tipo, grupo, familia |
| Cantidad | cantidad, qty, stock, existencia, quantity, unidades |
| Existencia mínima | existencia minima, stock minimo, cantidad minima, minimo |
| Existencia máxima | existencia maxima, stock maximo, cantidad maxima, maximo |
| Ubicación | ubicacion, location, sede, almacen, bodega |
| Código | codigo, código, cod, sku, item_code, code, barcode |

**Notas importantes:**
- Precio y costo se detectan por columnas separadas — si el Excel trae ambas ("PRECIO VENTA" y "PRECIO COMPRA", por ejemplo), no se confunden entre sí aunque compartan palabras.
- Si el Excel trae **dos** columnas que matchean las palabras clave de código, la primera llena `PRODUCTOS.codigo` y la segunda `PRODUCTOS.codigo_item`.
- Si un producto no trae cantidad (columna vacía o no detectada), la cantidad queda en **0** (no se asume 1).
- La columna `utilidad` de PRODUCTO SEDES se calcula automáticamente: `((precio_venta / costo) - 1) * 100`.

---

## Tablas dinámicas

Las hojas **CATEGORIA**, **MARCA** y **UBICACIONES** siempre comienzan con `id=1, descripcion="GENERAL"` y se amplían automáticamente:

- Si el Excel de entrada tiene columna de **marca** → se extrae y agrega a la tabla MARCA.
- Si el Excel de entrada tiene columna de **categoría** → se extrae y agrega a la tabla CATEGORIA.
- Si no hay columna explícita, se **infiere** desde la descripción del producto usando palabras clave.

---

## Personalizar categorías con palabras clave

Edite el diccionario `CATEGORIA_KEYWORDS` al inicio de `main.py`:

```python
CATEGORIA_KEYWORDS: dict[str, list[str]] = {
    "FRONTAL":     ["frontal", "front load"],
    "2P":          ["refrigerador", "nevera", "2p"],
    # Agregar nueva categoría:
    "SPLIT":       ["split", "minisplit", "aire acondicionado"],
    "HORNO":       ["horno", "cocina", "estufa"],
}
```

## Personalizar marcas conocidas

Edite la lista `KNOWN_BRANDS` al inicio de `main.py`:

```python
KNOWN_BRANDS: list[str] = [
    "PANASONIC", "SHARP", "HITACHI", "TOSHIBA",
    # Agregar nuevas marcas:
    "MIRAY", "ELECTROLUX", "INDURAMA",
]
```

---

## Hojas del Excel de salida

| Hoja | Contenido |
|------|-----------|
| CLIENTE PROVEEDORES | Vacía (plantilla) |
| PRODUCTOS | Un producto por fila con todos los campos requeridos |
| PRODUCTOS DETALLES | Stock inicial (cantidad, costo) por sede |
| PRODUCTO SEDES | Precio y stock por sede |
| UNIDADES | Lista fija de 124 unidades del SIN Bolivia |
| CATEGORIA | Categorías detectadas dinámicamente + GENERAL |
| MARCA | Marcas detectadas dinámicamente + GENERAL |
| UBICACIONES | Ubicaciones detectadas dinámicamente + GENERAL |

---

## Solución de problemas

| Problema | Solución |
|----------|----------|
| `ModuleNotFoundError: pandas` | `pip install pandas openpyxl` |
| Columna de precio no detectada | Renombrar la columna a "precio" o "costo" en el Excel de entrada |
