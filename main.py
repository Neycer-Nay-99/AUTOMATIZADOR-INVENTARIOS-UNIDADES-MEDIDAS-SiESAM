#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automatizador de Inventarios — SI ESAM  v1.2
Convierte cualquier inventario (Excel, CSV, TXT o imagen) al formato oficial.
CATEGORIAS, MARCAS y UBICACION se construyen dinámicamente desde los datos de entrada.
UNIDADES es una lista fija completa (SIN, Bolivia).
Si el origen trae columnas Unidad/Factor, se generan además UNIDAD_MEDIDA y
PRODUCTOS_UNIDADES_MEDIDAS para presentaciones alternas (caja, docena, etc.).
Código de producto: se extrae del origen si existe la columna, si no se genera incremental.
"""
from __future__ import annotations

import os
import re
import traceback
import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# ─── Dependencias opcionales ──────────────────────────────────────────────────
try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False
    pd = None  # type: ignore

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False
    Font = PatternFill = Alignment = None  # type: ignore

try:
    from PIL import Image
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    PIL_OK = True
except ImportError:
    PIL_OK = False

try:
    import cv2
    import numpy as np
    CV2_OK = True
except ImportError:
    CV2_OK = False
    np = None  # type: ignore

# ─── UNIDADES — lista fija completa (SIN Bolivia) ────────────────────────────
UNIDADES_REF: list[dict] = [
    {"ID": 1,   "DESCRIPCION": "BOBINAS",                                  "ESTADO": 1},
    {"ID": 2,   "DESCRIPCION": "BALDE",                                    "ESTADO": 1},
    {"ID": 3,   "DESCRIPCION": "BARRILES",                                 "ESTADO": 1},
    {"ID": 4,   "DESCRIPCION": "BOLSA",                                    "ESTADO": 1},
    {"ID": 5,   "DESCRIPCION": "BOTELLAS",                                 "ESTADO": 1},
    {"ID": 6,   "DESCRIPCION": "CAJA",                                     "ESTADO": 1},
    {"ID": 7,   "DESCRIPCION": "CARTONES",                                 "ESTADO": 1},
    {"ID": 8,   "DESCRIPCION": "CENTIMETRO CUADRADO",                      "ESTADO": 1},
    {"ID": 9,   "DESCRIPCION": "CENTIMETRO CUBICO",                        "ESTADO": 1},
    {"ID": 10,  "DESCRIPCION": "CENTIMETRO LINEAL",                        "ESTADO": 1},
    {"ID": 11,  "DESCRIPCION": "CIENTO DE UNIDADES",                       "ESTADO": 1},
    {"ID": 12,  "DESCRIPCION": "CILINDRO",                                 "ESTADO": 1},
    {"ID": 13,  "DESCRIPCION": "CONOS",                                    "ESTADO": 1},
    {"ID": 14,  "DESCRIPCION": "DOCENA",                                   "ESTADO": 1},
    {"ID": 15,  "DESCRIPCION": "FARDO",                                    "ESTADO": 1},
    {"ID": 16,  "DESCRIPCION": "GALON INGLES",                             "ESTADO": 1},
    {"ID": 17,  "DESCRIPCION": "GRAMO",                                    "ESTADO": 1},
    {"ID": 18,  "DESCRIPCION": "GRUESA",                                   "ESTADO": 1},
    {"ID": 19,  "DESCRIPCION": "HECTOLITRO",                               "ESTADO": 1},
    {"ID": 20,  "DESCRIPCION": "HOJA",                                     "ESTADO": 1},
    {"ID": 21,  "DESCRIPCION": "JUEGO",                                    "ESTADO": 1},
    {"ID": 22,  "DESCRIPCION": "KILOGRAMO",                                "ESTADO": 1},
    {"ID": 23,  "DESCRIPCION": "KILOMETRO",                                "ESTADO": 1},
    {"ID": 24,  "DESCRIPCION": "KILOVATIO HORA",                           "ESTADO": 1},
    {"ID": 25,  "DESCRIPCION": "KIT",                                      "ESTADO": 1},
    {"ID": 26,  "DESCRIPCION": "LATAS",                                    "ESTADO": 1},
    {"ID": 27,  "DESCRIPCION": "LIBRAS",                                   "ESTADO": 1},
    {"ID": 28,  "DESCRIPCION": "LITRO",                                    "ESTADO": 1},
    {"ID": 29,  "DESCRIPCION": "MEGAWATT HORA",                            "ESTADO": 1},
    {"ID": 30,  "DESCRIPCION": "METRO",                                    "ESTADO": 1},
    {"ID": 31,  "DESCRIPCION": "METRO CUADRADO",                           "ESTADO": 1},
    {"ID": 32,  "DESCRIPCION": "METRO CUBICO",                             "ESTADO": 1},
    {"ID": 33,  "DESCRIPCION": "MILIGRAMOS",                               "ESTADO": 1},
    {"ID": 34,  "DESCRIPCION": "MILILITRO",                                "ESTADO": 1},
    {"ID": 35,  "DESCRIPCION": "MILIMETRO",                                "ESTADO": 1},
    {"ID": 36,  "DESCRIPCION": "MILIMETRO CUADRADO",                       "ESTADO": 1},
    {"ID": 37,  "DESCRIPCION": "MILIMETRO CUBICO",                         "ESTADO": 1},
    {"ID": 38,  "DESCRIPCION": "MILLARES",                                 "ESTADO": 1},
    {"ID": 39,  "DESCRIPCION": "MILLON DE UNIDADES",                       "ESTADO": 1},
    {"ID": 40,  "DESCRIPCION": "ONZAS",                                    "ESTADO": 1},
    {"ID": 41,  "DESCRIPCION": "PALETAS",                                  "ESTADO": 1},
    {"ID": 42,  "DESCRIPCION": "PAQUETE",                                  "ESTADO": 1},
    {"ID": 43,  "DESCRIPCION": "PAR",                                      "ESTADO": 1},
    {"ID": 44,  "DESCRIPCION": "PIES",                                     "ESTADO": 1},
    {"ID": 45,  "DESCRIPCION": "PIES CUADRADOS",                           "ESTADO": 1},
    {"ID": 46,  "DESCRIPCION": "PIES CUBICOS",                             "ESTADO": 1},
    {"ID": 47,  "DESCRIPCION": "PIEZAS",                                   "ESTADO": 1},
    {"ID": 48,  "DESCRIPCION": "PLACAS",                                   "ESTADO": 1},
    {"ID": 49,  "DESCRIPCION": "PLIEGO",                                   "ESTADO": 1},
    {"ID": 50,  "DESCRIPCION": "PULGADAS",                                 "ESTADO": 1},
    {"ID": 51,  "DESCRIPCION": "RESMA",                                    "ESTADO": 1},
    {"ID": 52,  "DESCRIPCION": "TAMBOR",                                   "ESTADO": 1},
    {"ID": 53,  "DESCRIPCION": "TONELADA CORTA",                           "ESTADO": 1},
    {"ID": 54,  "DESCRIPCION": "TONELADA LARGA",                           "ESTADO": 1},
    {"ID": 55,  "DESCRIPCION": "TONELADAS",                                "ESTADO": 1},
    {"ID": 56,  "DESCRIPCION": "TUBOS",                                    "ESTADO": 1},
    {"ID": 57,  "DESCRIPCION": "UNIDAD (BIENES)",                          "ESTADO": 1},
    {"ID": 58,  "DESCRIPCION": "UNIDAD (SERVICIOS)",                       "ESTADO": 1},
    {"ID": 59,  "DESCRIPCION": "US GALON (3.7843 L)",                      "ESTADO": 1},
    {"ID": 60,  "DESCRIPCION": "YARDA",                                    "ESTADO": 1},
    {"ID": 61,  "DESCRIPCION": "YARDA CUADRADA",                           "ESTADO": 1},
    {"ID": 62,  "DESCRIPCION": "OTRO",                                     "ESTADO": 1},
    {"ID": 63,  "DESCRIPCION": "ONZA TROY",                                "ESTADO": 1},
    {"ID": 64,  "DESCRIPCION": "LIBRA FINA",                               "ESTADO": 1},
    {"ID": 65,  "DESCRIPCION": "DISPLAY",                                  "ESTADO": 1},
    {"ID": 66,  "DESCRIPCION": "BULTO",                                    "ESTADO": 1},
    {"ID": 67,  "DESCRIPCION": "DIAS",                                     "ESTADO": 1},
    {"ID": 68,  "DESCRIPCION": "MESES",                                    "ESTADO": 1},
    {"ID": 69,  "DESCRIPCION": "QUINTAL",                                  "ESTADO": 1},
    {"ID": 70,  "DESCRIPCION": "ROLLO",                                    "ESTADO": 1},
    {"ID": 71,  "DESCRIPCION": "HORAS",                                    "ESTADO": 1},
    {"ID": 72,  "DESCRIPCION": "AGUJA",                                    "ESTADO": 1},
    {"ID": 73,  "DESCRIPCION": "AMPOLLA",                                  "ESTADO": 1},
    {"ID": 74,  "DESCRIPCION": "BIDÓN",                               "ESTADO": 1},
    {"ID": 76,  "DESCRIPCION": "CAPSULA",                                  "ESTADO": 1},
    {"ID": 77,  "DESCRIPCION": "CARTUCHO",                                 "ESTADO": 1},
    {"ID": 78,  "DESCRIPCION": "COMPRIMIDO",                               "ESTADO": 1},
    {"ID": 79,  "DESCRIPCION": "ESTUCHE",                                  "ESTADO": 1},
    {"ID": 80,  "DESCRIPCION": "FRASCO",                                   "ESTADO": 1},
    {"ID": 81,  "DESCRIPCION": "JERINGA",                                  "ESTADO": 1},
    {"ID": 82,  "DESCRIPCION": "MINI BOTELLA",                             "ESTADO": 1},
    {"ID": 83,  "DESCRIPCION": "SACHET",                                   "ESTADO": 1},
    {"ID": 84,  "DESCRIPCION": "TABLETA",                                  "ESTADO": 1},
    {"ID": 85,  "DESCRIPCION": "TERMO",                                    "ESTADO": 1},
    {"ID": 86,  "DESCRIPCION": "TUBO",                                     "ESTADO": 1},
    {"ID": 87,  "DESCRIPCION": "BARRIL (EEUU) 60 F",                       "ESTADO": 1},
    {"ID": 88,  "DESCRIPCION": "BARRIL [42 GALONES(EEUU)]",                "ESTADO": 1},
    {"ID": 89,  "DESCRIPCION": "METRO CUBICO 68F VOL",                     "ESTADO": 1},
    {"ID": 90,  "DESCRIPCION": "MIL PIES CUBICOS 14696 PSI",               "ESTADO": 1},
    {"ID": 91,  "DESCRIPCION": "MIL PIES CUBICOS 14696 PSI 68FAH",         "ESTADO": 1},
    {"ID": 92,  "DESCRIPCION": "MILLAR DE PIES CUBICOS (1000 PC)",          "ESTADO": 1},
    {"ID": 93,  "DESCRIPCION": "MILLONES DE PIES CUBICOS (1000000 PC)",     "ESTADO": 1},
    {"ID": 94,  "DESCRIPCION": "MILLONES DE BTU (1000000 BTU)",             "ESTADO": 1},
    {"ID": 95,  "DESCRIPCION": "UNIDAD TERMICA BRITANICA (TI)",             "ESTADO": 1},
    {"ID": 96,  "DESCRIPCION": "POMO",                                     "ESTADO": 1},
    {"ID": 97,  "DESCRIPCION": "VASO",                                     "ESTADO": 1},
    {"ID": 98,  "DESCRIPCION": "TETRAPACK",                                "ESTADO": 1},
    {"ID": 99,  "DESCRIPCION": "CARTOLA",                                  "ESTADO": 1},
    {"ID": 100, "DESCRIPCION": "JABA",                                     "ESTADO": 1},
    {"ID": 102, "DESCRIPCION": "BANDEJA",                                  "ESTADO": 1},
    {"ID": 103, "DESCRIPCION": "TURRIL",                                   "ESTADO": 1},
    {"ID": 104, "DESCRIPCION": "BLISTER",                                  "ESTADO": 1},
    {"ID": 105, "DESCRIPCION": "TIRA",                                     "ESTADO": 1},
    {"ID": 106, "DESCRIPCION": "MEGAWATT",                                 "ESTADO": 1},
    {"ID": 107, "DESCRIPCION": "KILOWATT",                                 "ESTADO": 1},
    {"ID": 108, "DESCRIPCION": "AMORTIZACION",                             "ESTADO": 1},
    {"ID": 109, "DESCRIPCION": "OVULOS",                                   "ESTADO": 1},
    {"ID": 110, "DESCRIPCION": "SUPOSITORIOS",                             "ESTADO": 1},
    {"ID": 111, "DESCRIPCION": "SOBRES",                                   "ESTADO": 1},
    {"ID": 112, "DESCRIPCION": "VIAL",                                     "ESTADO": 1},
    {"ID": 113, "DESCRIPCION": "HECTAREAS",                                "ESTADO": 1},
    {"ID": 114, "DESCRIPCION": "ARROBA",                                   "ESTADO": 1},
    {"ID": 115, "DESCRIPCION": "AEROSOL",                                  "ESTADO": 1},
    {"ID": 116, "DESCRIPCION": "BARRA",                                    "ESTADO": 1},
    {"ID": 117, "DESCRIPCION": "CONJUNTO",                                 "ESTADO": 1},
    {"ID": 118, "DESCRIPCION": "FANEGA",                                   "ESTADO": 1},
    {"ID": 119, "DESCRIPCION": "PACK",                                     "ESTADO": 1},
    {"ID": 120, "DESCRIPCION": "PIPETA",                                   "ESTADO": 1},
    {"ID": 121, "DESCRIPCION": "POTE",                                     "ESTADO": 1},
    {"ID": 122, "DESCRIPCION": "PASTILLA",                                 "ESTADO": 1},
    {"ID": 123, "DESCRIPCION": "TONELADA METRICA",                         "ESTADO": 1},
    {"ID": 124, "DESCRIPCION": "EQUIPOS",                                  "ESTADO": 1},
]

# ─── Constantes ───────────────────────────────────────────────────────────────
CLIENTE_PROV_COLS = [
    "id", "codigo", "correo_electronico", "tipo_documento_id", "numero_documento",
    "complemento", "nro_documento_complemento", "razon_social", "tipo", "direccion",
    "ciudad", "zona", "phonecode", "celular", "telefono", "es_verificado", "puntos",
    "usuario_id", "nombre_cliente", "referencia", "municipio_id", "tipo_entidad",
    "personal_id", "limite_credito",
]

SHEET_ORDER = [
    "CLIENTE PROVEEDORES", "PRODUCTOS", "PRODUCTOS DETALLES",
    "PRODUCTO SEDES", "UNIDADES", "UNIDAD_MEDIDA", "PRODUCTOS_UNIDADES_MEDIDAS",
    "CATEGORIAS", "MARCAS", "UBICACION",
]

# ─── Orden de columnas de exportación (debe igualar la plantilla oficial) ────
PRODUCTOS_COLS = [
    "id", "descripcion", "codigo", "codigo_actividad", "unidad_id", "imagen", "estado",
    "es_servicio", "categoria_id", "marca_id", "ubicacion_id", "descripcion_larga", "codigo_sin",
    "created_at", "updated_at", "slug", "descripcion_corta", "principio_activo", "registro_sanitario",
    "venta_controlada", "existencia_minima", "novedoso", "tipo", "codigo_item", "codigo_volvo",
    "codigo_linea", "es_bonificado", "existencia_maxima", "config_adicionales", "medidas",
]
PRODUCTO_SEDES_COLS = [
    "id", "producto_id", "costo", "sede_id", "precio_unitario", "esta_activo",
    "cantidad", "utilidad", "factor", "precio_unitario2", "cantidad2", "precio_unitario3", "cantidad3",
    "precio_unitario4", "cantidad4", "utilidad2", "utilidad3", "utilidad4", "precio_factura",
    "utilidad_factura", "aplicar_lista_precios", "calcular_cantidad_por_precio",
    "mostrar_en_ecommerce", "es_alquilable", "tarifa_alquiler",
]
CATEGORIAS_COLS = [
    "id", "descripcion", "categoria_padre_id", "slug", "descripcion_larga",
    "descripcion_corta", "es_para_menu",
]
MARCAS_COLS = ["id", "descripcion", "slug"]
UBICACION_COLS = ["id", "descripcion", "color"]
UNIDADES_EXPORT_COLS = ["ID", "DESCRIPCION"]
UNIDAD_MEDIDA_COLS = ["id", "unidad_id", "descripcion", "factor", "estado", "created_at", "updated_at"]
PRODUCTOS_UM_COLS = ["id", "producto_id", "unidad_medida_id", "precio_unitario", "orden", "calcular_precio"]

PLACEHOLDER = (
    "Ejemplo (uno por línea):\n"
    "LAVADORA 10KG PANASONIC NA-VG1000L - Bs 6500\n"
    "REFRIGERADOR 2P SHARP SJ-PT547NS - 5200\n"
    "MICROONDAS TOSHIBA 30L: 850\n"
    "TELEVISOR HITACHI 55 PULGADAS 4K; 4800"
)

# ─── Tabla dinámica ───────────────────────────────────────────────────────────

class DynamicTable:
    """
    Tabla que comienza con GENERAL (id=1) y agrega entradas nuevas
    de forma incremental. Deduplicación case-insensitive.
    """

    def __init__(self):
        self._rows: list[dict] = [{"id": 1, "descripcion": "GENERAL"}]
        self._idx:  dict[str, int] = {"GENERAL": 1}

    def get_or_add(self, text: str) -> int:
        """Devuelve el id existente, o agrega la entrada y devuelve el nuevo id."""
        if not text or not text.strip():
            return 1
        key = text.strip().upper()
        if key in self._idx:
            return self._idx[key]
        new_id = self._rows[-1]["id"] + 1
        self._rows.append({"id": new_id, "descripcion": key})
        self._idx[key] = new_id
        return new_id

    def items(self) -> list[dict]:
        return list(self._rows)

    def count(self) -> int:
        return len(self._rows)

# ─── Utilidades ───────────────────────────────────────────────────────────────

def now_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def make_slug(text: str) -> str:
    s = text.lower()
    for chars, repl in [
        ("áàäâ", "a"), ("éèëê", "e"),
        ("íìïî", "i"), ("óòöô", "o"),
        ("úùüû", "u"), ("ñ", "n"),
    ]:
        for c in chars:
            s = s.replace(c, repl)
    s = re.sub(r"[^a-z0-9\s-]", "", s)
    return re.sub(r"\s+", "-", s.strip())


def _resolve_brand_cat(p: dict) -> tuple[str, str]:
    """
    Devuelve (brand_name, cat_name) desde campos explícitos del producto.
    Si el campo está vacío o ausente, devuelve "GENERAL".
    No se realiza ninguna inferencia desde la descripción.
    """
    brand_name = (p.get("marca_text")     or "").strip().upper() or "GENERAL"
    cat_name   = (p.get("categoria_text") or "").strip().upper() or "GENERAL"
    return brand_name, cat_name


# ─── Resolución de unidad de medida por palabra clave ────────────────────────

_UNIT_KEYWORD_RULES: list[tuple[re.Pattern, int]] = [
    (re.compile(r"\bCUARTA\b"), 14),   # DOCENA — 1/4 de docena
    (re.compile(r"\bDOCENA\b"), 14),   # también matchea "MEDIA DOCENA"
    (re.compile(r"\bCAJA\b"), 6),      # también matchea "MEDIA CAJA"
    (re.compile(r"\bGRUESA\b"), 18),
    (re.compile(r"\bPAQ"), 42),        # PAQUETE / PAQ / PAQ.
    (re.compile(r"\bFARDO\b"), 15),
    (re.compile(r"\bBOLSA\b"), 4),
    (re.compile(r"\bMILLAR"), 38),     # MILLAR / MILLARES
    (re.compile(r"\bBULTO\b"), 66),
    (re.compile(r"\bDISPLAY\b"), 65),
    (re.compile(r"\bCARTON"), 7),      # CARTON / CARTONES
    (re.compile(r"\b(SET|JUEGO)\b"), 21),
    (re.compile(r"\bUNIDAD"), 57),     # respaldo genérico antes de OTRO
]
_UNIT_FALLBACK_ID = 62  # OTRO


def _normalize_unit_text(text: str) -> str:
    s = (text or "").strip().lower()
    for chars, repl in [
        ("áàäâ", "a"), ("éèëê", "e"),
        ("íìïî", "i"), ("óòöô", "o"),
        ("úùüû", "u"), ("ñ", "n"),
    ]:
        for c in chars:
            s = s.replace(c, repl)
    return s.upper()


def resolve_unidad_id(text: str, default: int = _UNIT_FALLBACK_ID) -> int:
    """
    Resuelve el id de UNIDADES_REF que corresponde a un texto libre de unidad
    (p.ej. "CAJA (24DOC)", "MEDIA DOCENA", "CUARTA") por palabra clave.
    Si no hay match, devuelve `default` (OTRO por defecto).
    """
    t = _normalize_unit_text(text)
    if not t:
        return default
    for pattern, uid in _UNIT_KEYWORD_RULES:
        if pattern.search(t):
            return uid
    return default


def _format_factor(factor: float) -> str:
    """Formatea el factor sin decimales sobrantes: 20.0 -> '20', 2.5 -> '2.5'."""
    f = float(factor or 0)
    if f == int(f):
        return str(int(f))
    return str(f).rstrip("0").rstrip(".")


class UnidadMedidaTable:
    """
    Catálogo deduplicado de presentaciones alternas (hoja UNIDAD_MEDIDA):
    una fila por combinación única (descripcion, factor). A diferencia de
    DynamicTable, empieza vacío (sin fila semilla) y cada fila lleva además
    su propio unidad_id resuelto por palabra clave.
    """

    def __init__(self):
        self._rows: list[dict] = []
        self._idx: dict[tuple[str, float], int] = {}

    def get_or_add(self, descripcion: str, factor: float) -> int:
        desc = (descripcion or "").strip().upper()
        key = (desc, round(float(factor or 0), 6))
        if key in self._idx:
            return self._idx[key]
        new_id = len(self._rows) + 1
        self._rows.append({
            "id": new_id,
            "unidad_id": resolve_unidad_id(desc),
            "descripcion": f"{desc} X{_format_factor(factor)}",
            "factor": factor,
        })
        self._idx[key] = new_id
        return new_id

    def items(self) -> list[dict]:
        return list(self._rows)


# ─── Extracción de datos ──────────────────────────────────────────────────────

_PRICE_RE = [
    r"(?:bs\.?|bolivianos?)[:\s]*(\d+(?:[.,]\d{1,2})?)",
    r"(\d+(?:[.,]\d{1,2})?)\s*(?:bs\.?|bolivianos?)",
    r"[-–:;,\t]\s*(\d+(?:[.,]\d{1,2})?)\s*$",
    r"\s(\d{1,6}(?:\.\d{1,2})?)\s*$",
]


def _parse_line(line: str) -> tuple[str, float]:
    for pat in _PRICE_RE:
        m = re.search(pat, line, re.IGNORECASE)
        if m:
            try:
                price = float(m.group(1).replace(",", "."))
                name  = line[: m.start()].strip()
                name  = re.sub(r"[-–:;,]\s*$", "", name).strip()
                return name, price
            except ValueError:
                pass
    return line.strip(), 0.0


def extract_from_text(raw: str) -> list[dict]:
    products = []
    for line in raw.splitlines():
        line = line.strip()
        if not line or len(line) < 2:
            continue
        name, price = _parse_line(line)
        if name:
            products.append({
                "descripcion":    name.upper(),
                "precio":         price,
                "cantidad":       1,
                "marca_text":     "",
                "categoria_text": "",
                "ubicacion_text": "",
            })
    return products


# Heurística de columnas para Excel / CSV
_COL_NAME  = ["nombre", "producto", "descripcion", "descripción", "item",
               "articulo", "artículo", "name", "product"]
_COL_PRICE = ["precio", "price", "valor", "pvp", "tarifa", "importe", "monto", "venta"]
_COL_COST  = ["costo", "cost", "compra"]
_COL_BRAND = ["marca", "brand", "fabricante"]
_COL_CAT   = ["categoria", "categoría", "category", "tipo", "grupo",
               "familia", "family"]
_COL_QTY   = ["cantidad", "qty", "stock", "existencia", "quantity", "unidades"]
_COL_EXIST_MIN = ["existencia minima", "existencia mínima", "exist minima",
                   "exist mínima", "stock minimo", "stock mínimo",
                   "cantidad minima", "cantidad mínima", "minimo", "mínimo"]
_COL_EXIST_MAX = ["existencia maxima", "existencia máxima", "exist maxima",
                   "exist máxima", "stock maximo", "stock máximo",
                   "cantidad maxima", "cantidad máxima", "maximo", "máximo"]
_COL_UBIC  = ["ubicacion", "ubicación", "location", "sede", "almacen",
               "almacén", "bodega", "deposito", "depósito"]
_COL_DESC_SHORT = ["descripcion corta", "descripción corta", "desc corta",
                     "descripcion breve", "descripción breve",
                     "short description", "short desc"]
_COL_CODE  = ["codigo", "código", "cod", "sku", "item_code", "item code",
               "product_code", "code", "codebar", "barcode"]
_COL_UNIDAD = ["unidad", "presentacion", "presentación", "empaque"]
_COL_FACTOR = ["factor", "conversion", "conversión", "equivalencia", "multiplicador"]


def _find_col(df, keywords: list[str], exclude: set | None = None) -> str | None:
    for col in df.columns:
        if exclude and col in exclude:
            continue
        if any(kw in str(col).lower() for kw in keywords):
            return col
    return None


def _find_cols(df, keywords: list[str]) -> list[str]:
    """Todas las columnas cuyo encabezado matchea alguna keyword (en orden)."""
    return [col for col in df.columns if any(kw in str(col).lower() for kw in keywords)]


def _safe_str(val) -> str:
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "") else s


def df_to_products(df, log=None) -> list[dict]:
    # "descripcion corta" se busca antes que el nombre y se excluye de esa
    # búsqueda, porque también matchea la keyword "descripcion" de _COL_NAME
    # (mismo patrón que exist_min/exist_max con "existencia")
    col_desc_short = _find_col(df, _COL_DESC_SHORT)
    col_name  = _find_col(df, _COL_NAME,
                           exclude={col_desc_short} if col_desc_short else None)
    col_cost  = _find_col(df, _COL_COST)
    # el precio de venta se busca excluyendo la columna ya asignada a costo,
    # para evitar que "PRECIO COMPRA" (que también matchea "precio") se
    # confunda con "PRECIO VENTA".
    # Algunos orígenes (ej. DUNAMIS) traen una segunda columna "PRECIO" después
    # de FACTOR, dedicada al precio de las filas de continuación (presentación
    # alterna): esa segunda columna se guarda aparte como col_price_alt.
    _price_cols = [c for c in _find_cols(df, _COL_PRICE) if c != col_cost]
    col_price = _price_cols[0] if _price_cols else None
    col_price_alt = _price_cols[1] if len(_price_cols) > 1 else col_price
    col_brand = _find_col(df, _COL_BRAND)
    col_cat   = _find_col(df, _COL_CAT)
    # min/max se buscan antes que la cantidad normal, y se excluyen de esa
    # búsqueda, porque "existencia minima/maxima" también matchea la keyword
    # "existencia" de _COL_QTY
    col_exist_min = _find_col(df, _COL_EXIST_MIN)
    col_exist_max = _find_col(df, _COL_EXIST_MAX,
                               exclude={col_exist_min} if col_exist_min else None)
    col_qty   = _find_col(
        df, _COL_QTY,
        exclude={c for c in (col_exist_min, col_exist_max) if c},
    )
    col_ubic  = _find_col(df, _COL_UBIC)

    # unidad/factor: columnas opcionales para presentaciones alternas
    # (ej. archivos DUNAMIS con filas de continuación CAJA/DOCENA/etc.)
    col_factor = _find_col(df, _COL_FACTOR)
    col_unidad = _find_col(df, _COL_UNIDAD, exclude={col_qty} if col_qty else None)
    has_multi_unit = col_unidad is not None and col_factor is not None

    code_cols = _find_cols(df, _COL_CODE)
    col_code  = code_cols[0] if code_cols else None
    col_code2 = code_cols[1] if len(code_cols) > 1 else None

    if col_name is None:
        for c in df.columns:
            if df[c].dtype == object:
                col_name = c
                break

    if log:
        log(f"  nombre={col_name}  precio={col_price}  precio_alt={col_price_alt}  costo={col_cost}  "
            f"marca={col_brand}  cat={col_cat}  qty={col_qty}  "
            f"exist_min={col_exist_min}  exist_max={col_exist_max}  "
            f"ubic={col_ubic}  codigo={col_code}  codigo_item={col_code2}  "
            f"desc_corta={col_desc_short}  unidad={col_unidad}  factor={col_factor}")

    products: list[dict] = []
    last_product: dict | None = None
    for _, row in df.iterrows():
        name = _safe_str(row[col_name]) if col_name else ""
        if name:
            try:
                price = float(_safe_str(row[col_price]).replace(",", ".")) if col_price else 0.0
            except ValueError:
                price = 0.0
            try:
                cost = float(_safe_str(row[col_cost]).replace(",", ".")) if col_cost else 0.0
            except ValueError:
                cost = 0.0
            try:
                qty = int(float(_safe_str(row[col_qty]).replace(",", "."))) if col_qty else 0
            except ValueError:
                qty = 0
            try:
                exist_min = int(float(_safe_str(row[col_exist_min]).replace(",", "."))) if col_exist_min else 0
            except ValueError:
                exist_min = 0
            try:
                exist_max = int(float(_safe_str(row[col_exist_max]).replace(",", "."))) if col_exist_max else 0
            except ValueError:
                exist_max = 0
            product = {
                "descripcion":    name.upper(),
                "precio":         price,
                "costo":          cost,
                "cantidad":       max(qty, 0),
                "existencia_minima": max(exist_min, 0),
                "existencia_maxima": max(exist_max, 0),
                "marca_text":     _safe_str(row[col_brand]).upper() if col_brand else "",
                "categoria_text": _safe_str(row[col_cat]).upper()   if col_cat   else "",
                "ubicacion_text": _safe_str(row[col_ubic]).upper()  if col_ubic  else "",
                "codigo_origen":  _safe_str(row[col_code])          if col_code  else "",
                "codigo_item_origen": _safe_str(row[col_code2])     if col_code2 else "",
                "descripcion_corta_origen": _safe_str(row[col_desc_short]) if col_desc_short else "",
                "unidad_texto":   _safe_str(row[col_unidad]) if col_unidad else "",
                "unidades_alternas": [],
            }
            products.append(product)
            last_product = product
            continue

        # Fila sin nombre: posible fila de continuación (presentación alterna)
        if has_multi_unit and last_product is not None:
            unidad_text = _safe_str(row[col_unidad])
            factor_text = _safe_str(row[col_factor])
            if not unidad_text and not factor_text:
                continue  # fila de relleno completamente vacía
            try:
                factor_val = float(factor_text.replace(",", ".")) if factor_text else 0.0
            except ValueError:
                factor_val = 0.0
            try:
                alt_precio = float(_safe_str(row[col_price_alt]).replace(",", ".")) if col_price_alt else 0.0
            except ValueError:
                alt_precio = 0.0
            if factor_val <= 0 and log:
                log(f"  ATENCION: factor invalido ({factor_text!r}) en unidad alterna "
                    f"'{unidad_text}' de '{last_product['descripcion']}' — se registra igual.", "warn")
            alt_costo = _safe_str(row[col_cost]) if col_cost else ""
            if alt_costo and log:
                log(f"  ATENCION: costo ({alt_costo}) en fila de unidad alterna '{unidad_text}' "
                    f"de '{last_product['descripcion']}' se descarta (sin campo destino).", "warn")
            last_product["unidades_alternas"].append({
                "descripcion": unidad_text,
                "factor":      factor_val,
                "precio":      alt_precio,
            })
    return products


def extract_from_excel(path: str, log=None) -> list[dict]:
    xl = pd.ExcelFile(path)
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        if df.empty:
            continue
        if log:
            log(f"Hoja '{sheet}': {len(df)} filas, {len(df.columns)} columnas")
        return df_to_products(df, log=log)
    return []


def extract_from_csv(path: str, log=None) -> list[dict]:
    for sep in (",", ";", "\t", "|"):
        try:
            df = pd.read_csv(path, sep=sep, encoding="utf-8", errors="ignore")
            if len(df.columns) > 1:
                if log:
                    log(f"CSV sep='{sep}': {len(df)} filas")
                return df_to_products(df, log=log)
        except Exception:
            pass
    return []


def extract_from_image(path: str, log=None) -> list[dict]:
    if not PIL_OK:
        if log:
            log("ERROR: Instale pillow y pytesseract para OCR.")
        return []
    try:
        img = Image.open(path).convert("L")
        if CV2_OK:
            arr = np.array(img)
            arr = cv2.equalizeHist(arr)
            _, arr = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            h, w = arr.shape
            scale = max(1, 2000 // max(h, w, 1))
            if scale > 1:
                arr = cv2.resize(arr, (w * scale, h * scale),
                                 interpolation=cv2.INTER_CUBIC)
            img = Image.fromarray(arr)
            if log:
                log("Imagen preprocesada con OpenCV.")
        try:
            text = pytesseract.image_to_string(
                img, lang="spa+eng", config="--oem 3 --psm 6")
        except Exception:
            text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")
        if log:
            log(f"OCR: {len(text)} caracteres extraidos.")
        return extract_from_text(text)
    except Exception as e:
        if log:
            log(f"Error OCR: {e}")
        return []


# ─── Constructor de hojas oficiales ──────────────────────────────────────────

def build_sheets(
    products:   list[dict],
    codigo_ini: int = 10001,
    sede_id:    int = 1,
    unidad_id:  int = 57,
) -> dict[str, "pd.DataFrame"]:
    """
    Construye todas las hojas del formato oficial.
    CATEGORIAS, MARCAS y UBICACION se generan dinámicamente a partir de los productos.
    UNIDADES es fija (UNIDADES_REF). UNIDAD_MEDIDA y PRODUCTOS_UNIDADES_MEDIDAS se
    generan solo si los productos traen presentaciones alternas (unidades_alternas).
    """
    cat_tbl  = DynamicTable()
    brand_tbl = DynamicTable()
    ubic_tbl  = DynamicTable()
    unit_tbl  = UnidadMedidaTable()

    prods, sedes, detalles, prod_units = [], [], [], []
    _auto_codigo = codigo_ini  # contador para productos sin código de origen

    for i, p in enumerate(products, start=1):
        desc   = p.get("descripcion", f"PRODUCTO {i}")
        precio = p.get("precio", 0.0)
        costo  = round(p.get("costo", 0.0), 5)
        cant   = max(p.get("cantidad", 1), 0)

        brand_name, cat_name = _resolve_brand_cat(p)

        # Ubicación: usar columna explícita si existe, sino GENERAL
        ubic_text = (p.get("ubicacion_text") or "").strip().upper()

        c_id = cat_tbl.get_or_add(cat_name)    if cat_name  != "GENERAL" else 1
        m_id = brand_tbl.get_or_add(brand_name) if brand_name != "GENERAL" else 1
        u_id = ubic_tbl.get_or_add(ubic_text)   if ubic_text              else 1

        # Unidad base: resuelta por producto desde el texto de origen si existe,
        # o el valor global de la GUI como respaldo.
        unidad_texto = (p.get("unidad_texto") or "").strip()
        producto_unidad_id = resolve_unidad_id(unidad_texto, default=unidad_id) if unidad_texto else unidad_id

        # Código: respetar el del origen; generar incremental si está vacío
        codigo_origen = (p.get("codigo_origen") or "").strip()
        if codigo_origen:
            codigo_val = codigo_origen
        else:
            codigo_val = str(_auto_codigo)
            _auto_codigo += 1

        prods.append({
            "id":                 i,
            "descripcion":        desc,
            "codigo":             codigo_val,
            "codigo_actividad":   "",
            "codigo_item":        (p.get("codigo_item_origen") or "").strip(),
            "codigo_linea":       "",
            "unidad_id":          producto_unidad_id,
            "imagen":             "",
            "estado":             1,
            "es_servicio":        0,
            "categoria_id":       c_id,
            "marca_id":           m_id,
            "ubicacion_id":       u_id,
            "descripcion_larga":  "",
            "codigo_sin":         "",
            "created_at":         "",
            "updated_at":         "",
            "slug":               "",
            "descripcion_corta":  (p.get("descripcion_corta_origen") or "").strip(),
            "principio_activo":   "",
            "registro_sanitario": "",
            "venta_controlada":   0,
            "existencia_minima":  p.get("existencia_minima", 0),
            "tipo":               "",
            "novedoso":           0,
            "codigo_volvo":       "",
            "es_bonificado":      0,
            "existencia_maxima":  p.get("existencia_maxima", 0),
            "config_adicionales": "",
            "medidas":            "",
        })

        # Margen de utilidad: ((precio_unitario / costo) - 1) * 100
        # Sin costo pero con precio -> 100 por defecto. Con costo pero sin precio -> 0.
        if costo > 0 and precio > 0:
            utilidad_val = round(((precio / costo) - 1) * 100, 2)
        elif costo <= 0 and precio > 0:
            utilidad_val = 100
        else:
            utilidad_val = 0

        sedes.append({
            "id":                          i,
            "producto_id":                 i,
            "sede_id":                     sede_id,
            "precio_unitario":             precio,
            "esta_activo":                 1,
            "cantidad":                    1,  # el precio es por unidad; la cantidad real va en PRODUCTOS DETALLES
            "precio_unitario2":            0,
            "cantidad2":                   0,
            "precio_unitario3":            0,
            "cantidad3":                   0,
            "precio_unitario4":            0,
            "cantidad4":                   0,
            "utilidad":                    utilidad_val,
            "utilidad2":                   "",
            "utilidad3":                   "",
            "utilidad4":                   "",
            "costo":                       costo,
            "precio_factura":              "",
            "utilidad_factura":            "",
            "aplicar_lista_precios":       "",
            "factor":                      None,
            "calcular_cantidad_por_precio":"",
            "mostrar_en_ecommerce":        "",
            "es_alquilable":               "",
            "tarifa_alquiler":             "",
        })

        detalles.append({
            "id":          i,
            "producto_id": i,
            "sede_id":     sede_id,
            "cantidad":    cant,
            "costo":       costo,
            "created_at":  "",
            "updated_at":  "",
            "lote":        "",
            "vencimiento": "",
        })

        # Presentaciones alternas del producto (hoja PRODUCTOS_UNIDADES_MEDIDAS)
        for orden, alt in enumerate(p.get("unidades_alternas", []), start=1):
            um_id = unit_tbl.get_or_add(alt["descripcion"], alt["factor"])
            prod_units.append({
                "id":                len(prod_units) + 1,
                "producto_id":       i,
                "unidad_medida_id":  um_id,
                "precio_unitario":   alt.get("precio", 0.0),
                "orden":             orden,
                "calcular_precio":   0,
            })

    # Construir DataFrames de referencia desde las tablas dinámicas
    cats = [
        {"id": r["id"], "descripcion": r["descripcion"],
         "categoria_padre_id": "",
         "slug": "",
         "descripcion_larga": "", "descripcion_corta": "", "es_para_menu": 0}
        for r in cat_tbl.items()
    ]

    marcas = [
        {"id": r["id"], "descripcion": r["descripcion"],
         "slug": ""}
        for r in brand_tbl.items()
    ]

    ubicaciones = [
        {"id": r["id"], "descripcion": r["descripcion"], "color": ""}
        for r in ubic_tbl.items()
    ]

    unidades_medida = [
        {"id": r["id"], "unidad_id": r["unidad_id"], "descripcion": r["descripcion"],
         "factor": r["factor"], "estado": 1, "created_at": "", "updated_at": ""}
        for r in unit_tbl.items()
    ]

    return {
        "CLIENTE PROVEEDORES": pd.DataFrame(columns=CLIENTE_PROV_COLS),
        "PRODUCTOS":           pd.DataFrame(prods).reindex(columns=PRODUCTOS_COLS),
        "PRODUCTOS DETALLES":  pd.DataFrame(detalles),
        "PRODUCTO SEDES":      pd.DataFrame(sedes).reindex(columns=PRODUCTO_SEDES_COLS),
        "UNIDADES":            pd.DataFrame(UNIDADES_REF).reindex(columns=UNIDADES_EXPORT_COLS),
        "UNIDAD_MEDIDA":       pd.DataFrame(unidades_medida).reindex(columns=UNIDAD_MEDIDA_COLS),
        "PRODUCTOS_UNIDADES_MEDIDAS": pd.DataFrame(prod_units).reindex(columns=PRODUCTOS_UM_COLS),
        "CATEGORIAS":          pd.DataFrame(cats).reindex(columns=CATEGORIAS_COLS),
        "MARCAS":              pd.DataFrame(marcas).reindex(columns=MARCAS_COLS),
        "UBICACION":           pd.DataFrame(ubicaciones).reindex(columns=UBICACION_COLS),
    }


# ─── Escritura Excel ──────────────────────────────────────────────────────────

def write_excel(sheets: dict[str, "pd.DataFrame"], path: str) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name in SHEET_ORDER:
            if name in sheets:
                sheets[name].to_excel(writer, sheet_name=name, index=False)

        wb        = writer.book
        hdr_font  = Font(bold=True, color="FFFFFF", size=10)
        hdr_fill  = PatternFill("solid", fgColor="1F4E79")
        hdr_align = Alignment(horizontal="center", vertical="center")

        for ws in wb.worksheets:
            if ws.max_row < 1:
                continue
            for cell in ws[1]:
                cell.font      = hdr_font
                cell.fill      = hdr_fill
                cell.alignment = hdr_align
            ws.row_dimensions[1].height = 22
            ws.freeze_panes = "A2"
            for col in ws.columns:
                length = max((len(str(c.value or "")) for c in col), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(length + 4, 52)


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatizador de Inventarios — SI ESAM  v1.2")
        self.geometry("1050x700")
        self.minsize(820, 560)
        self.configure(bg="#F0F4F8")
        self._input_file: str | None = None
        self._placeholder_active = True
        self._build_ui()
        self._log("Sistema iniciado  (v1.2).")
        self._check_deps()

    # ── Construcción de la interfaz ───────────────────────────────────────────

    def _build_ui(self):
        self._setup_styles()
        self._build_header()
        pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        pane.add(self._build_left(pane),  weight=55)
        pane.add(self._build_right(pane), weight=45)
        self._build_statusbar()

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",            background="#F0F4F8")
        s.configure("TLabel",            background="#F0F4F8", font=("Segoe UI", 10))
        s.configure("TLabelframe",       background="#F0F4F8")
        s.configure("TLabelframe.Label", background="#F0F4F8", font=("Segoe UI", 10, "bold"))
        s.configure("TButton",           font=("Segoe UI", 10), padding=6)
        s.configure("TEntry",            font=("Segoe UI", 10))
        s.configure("TCheckbutton",      background="#F0F4F8", font=("Segoe UI", 10))
        s.configure("Treeview",          font=("Segoe UI", 9), rowheight=24)
        s.configure("Treeview.Heading",  font=("Segoe UI", 9, "bold"))

    def _build_header(self):
        fr = tk.Frame(self, bg="#1F4E79", pady=10)
        fr.pack(fill=tk.X)
        tk.Label(fr, text="CONVERTIDOR DE INVENTARIOS PARA SI ESAM",
                 bg="#1F4E79", fg="white", font=("Segoe UI", 16, "bold")).pack()
        tk.Label(fr,
                 text="Convierte Excel · CSV · TXT · Imagen  →  Formato oficial SI ESAM",
                 bg="#1F4E79", fg="#BDD7EE", font=("Segoe UI", 9)).pack()

    def _build_left(self, parent) -> ttk.LabelFrame:
        lf = ttk.LabelFrame(parent, text="  Entrada de datos  ", padding=10)

        ttk.Label(lf,
                  text="1.  Archivo de inventario (Excel, CSV, TXT o imagen JPG/PNG):"
                  ).pack(anchor=tk.W)
        row = ttk.Frame(lf)
        row.pack(fill=tk.X, pady=(4, 2))
        self._file_var = tk.StringVar(value="Sin archivo seleccionado")
        ttk.Entry(row, textvariable=self._file_var, state="readonly").pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="Examinar…",
                   command=self._pick_file).pack(side=tk.LEFT, padx=(6, 2))
        ttk.Button(row, text="✗", width=3,
                   command=self._clear_file).pack(side=tk.LEFT)

        ttk.Separator(lf, orient="horizontal").pack(fill=tk.X, pady=10)

        ttk.Label(lf,
                  text="2.  O pegar inventario en texto plano (uno por línea):"
                  ).pack(anchor=tk.W)
        ttk.Label(lf,
                  text='      Formatos: "Producto - Bs 99"  |  "Producto: 99"  |  "Producto; 99"',
                  foreground="#666666", font=("Segoe UI", 8)).pack(anchor=tk.W)

        self._txt = ScrolledText(lf, height=16, font=("Consolas", 10),
                                  wrap=tk.WORD, relief="solid", bd=1)
        self._txt.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self._txt.insert("1.0", PLACEHOLDER)
        self._txt.config(fg="#999999")
        self._txt.bind("<FocusIn>",  self._txt_focus_in)
        self._txt.bind("<FocusOut>", self._txt_focus_out)

        return lf

    def _build_right(self, parent) -> ttk.Frame:
        fr = ttk.Frame(parent)

        # Opciones
        opts = ttk.LabelFrame(fr, text="  Opciones de conversión  ", padding=10)
        opts.pack(fill=tk.X, pady=(0, 8))

        def opt_row(label: str, var: tk.StringVar, tip: str = ""):
            r = ttk.Frame(opts)
            r.pack(fill=tk.X, pady=3)
            ttk.Label(r, text=label, width=24, anchor=tk.W).pack(side=tk.LEFT)
            ttk.Entry(r, textvariable=var, width=12).pack(side=tk.LEFT)
            if tip:
                ttk.Label(r, text=tip, foreground="#888",
                          font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=6)

        self._v_codigo = tk.StringVar(value="10001")
        self._v_sede   = tk.StringVar(value="1")
        self._v_unidad = tk.StringVar(value="57")
        opt_row("Código inicial:", self._v_codigo, "(ej. 10001 → 10002 → …)")
        opt_row("Sede ID:",             self._v_sede,   "(por defecto: 1)")
        opt_row("Unidad ID:",           self._v_unidad, "(57 = UNIDAD BIENES)")

        self._v_preview = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Previsualizar productos antes de guardar",
                        variable=self._v_preview).pack(anchor=tk.W, pady=(8, 0))

        ttk.Button(opts,
                   text="ℹ  Ver columnas detectables y Unidades",
                   command=self._show_config).pack(fill=tk.X, pady=(10, 0))

        # Botón principal
        tk.Button(fr, text="⚡  CONVERTIR INVENTARIO",
                  bg="#1F4E79", fg="white", font=("Segoe UI", 12, "bold"),
                  relief="flat", pady=12, cursor="hand2",
                  activebackground="#2E75B6", activeforeground="white",
                  command=self._convert).pack(fill=tk.X, pady=8)

        # Log
        log_fr = ttk.LabelFrame(fr, text="  Registro  ", padding=6)
        log_fr.pack(fill=tk.BOTH, expand=True)

        self._log_box = ScrolledText(
            log_fr, height=18, font=("Consolas", 9),
            bg="#1E1E2E", fg="#CDD6F4",
            wrap=tk.WORD, relief="flat",
            state=tk.DISABLED, insertbackground="white")
        self._log_box.pack(fill=tk.BOTH, expand=True)
        self._log_box.tag_config("ok",   foreground="#A6E3A1")
        self._log_box.tag_config("warn", foreground="#F9E2AF")
        self._log_box.tag_config("err",  foreground="#F38BA8")
        self._log_box.tag_config("info", foreground="#89B4FA")

        return fr

    def _build_statusbar(self):
        bar = tk.Frame(self, bg="#DDE3EA", pady=4)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        self._status = tk.StringVar(value="Listo — seleccione un archivo o pegue texto.")
        tk.Label(bar, textvariable=self._status, bg="#DDE3EA",
                 font=("Segoe UI", 9), fg="#3D4451", anchor=tk.W
                 ).pack(side=tk.LEFT, padx=10)

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = ""):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log_box.config(state=tk.NORMAL)
        self._log_box.insert(tk.END, f"[{ts}] {msg}\n", tag if tag else ())
        self._log_box.see(tk.END)
        self._log_box.config(state=tk.DISABLED)
        self.update_idletasks()

    def _set_status(self, msg: str):
        self._status.set(msg)
        self.update_idletasks()

    def _check_deps(self):
        if not PANDAS_OK:
            self._log("FALTA: pandas  →  pip install pandas", "err")
        if not OPENPYXL_OK:
            self._log("FALTA: openpyxl  →  pip install openpyxl", "err")
        if not PIL_OK:
            self._log("OCR desactivado  →  pip install pillow pytesseract", "warn")
        if not CV2_OK and PIL_OK:
            self._log("Pre-proc. imagen desactivado  →  pip install opencv-python", "warn")
        if PANDAS_OK and OPENPYXL_OK:
            self._log("Dependencias principales OK.", "ok")

    def _txt_focus_in(self, _):
        if self._placeholder_active:
            self._txt.delete("1.0", tk.END)
            self._txt.config(fg="black")
            self._placeholder_active = False

    def _txt_focus_out(self, _):
        if not self._txt.get("1.0", tk.END).strip():
            self._txt.insert("1.0", PLACEHOLDER)
            self._txt.config(fg="#999999")
            self._placeholder_active = True

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo de inventario",
            filetypes=[
                ("Todos los soportados", "*.xlsx *.xls *.csv *.txt *.jpg *.jpeg *.png *.bmp"),
                ("Excel",  "*.xlsx *.xls"),
                ("CSV",    "*.csv"),
                ("Texto",  "*.txt"),
                ("Imagen", "*.jpg *.jpeg *.png *.bmp"),
                ("Todos",  "*.*"),
            ],
        )
        if path:
            self._input_file = path
            self._file_var.set(Path(path).name)
            self._log(f"Archivo: {path}")

    def _clear_file(self):
        self._input_file = None
        self._file_var.set("Sin archivo seleccionado")

    # ── Ver información del sistema ───────────────────────────────────────────

    def _show_config(self):
        win = tk.Toplevel(self)
        win.title("Información del sistema")
        win.geometry("740x520")
        win.grab_set()

        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Tab 1: detección de columnas
        tab_cols = ttk.Frame(nb, padding=4)
        nb.add(tab_cols, text="  Detección de columnas  ")

        ttk.Label(tab_cols,
                  text="Palabras clave usadas para identificar columnas automáticamente en Excel/CSV:",
                  foreground="#333").pack(anchor=tk.W, pady=(0, 6))

        col_defs = [
            ("Código/SKU",         "codigo, código, cod, sku, item_code, item code, product_code, code, codebar, barcode "
                                    "(si hay 2 columnas que matchean, la 1ra va a 'codigo' y la 2da a 'codigo_item')"),
            ("Nombre/Descripción", "nombre, producto, descripcion, item, articulo, name, product"),
            ("Precio",             "precio, price, costo, cost, valor, pvp, tarifa, importe, monto"),
            ("Marca",              "marca, brand, fabricante"),
            ("Categoría",          "categoria, category, tipo, grupo, familia, family"),
            ("Cantidad",           "cantidad, qty, stock, existencia, quantity, unidades"),
            ("Ubicación",          "ubicacion, location, sede, almacen, bodega, deposito"),
        ]
        cols_h = ("Campo destino", "Palabras clave detectadas en el encabezado")
        tree_c = ttk.Treeview(tab_cols, columns=cols_h, show="headings", height=10)
        tree_c.heading(cols_h[0], text=cols_h[0])
        tree_c.heading(cols_h[1], text=cols_h[1])
        tree_c.column(cols_h[0], width=160)
        tree_c.column(cols_h[1], width=520)
        for campo, kws in col_defs:
            tree_c.insert("", tk.END, values=(campo, kws))
        tree_c.pack(fill=tk.BOTH, expand=True)

        ttk.Label(tab_cols,
                  text="Nota: si el archivo no tiene columna de Marca, Categoría o Ubicación, "
                       "se asigna GENERAL a todos los productos.",
                  foreground="#666", font=("Segoe UI", 8),
                  wraplength=680).pack(anchor=tk.W, pady=(6, 0))

        # Tab 2: unidades (solo consulta)
        tab_unid = ttk.Frame(nb, padding=4)
        nb.add(tab_unid, text=f"  Unidades ({len(UNIDADES_REF)})  ")

        ttk.Label(tab_unid,
                  text="Lista fija de unidades del SIN Bolivia (no se modifica con la entrada).",
                  foreground="#555").pack(anchor=tk.W, pady=(0, 4))

        cols_u = ("ID", "Descripcion", "Estado")
        tree_u = ttk.Treeview(tab_unid, columns=cols_u, show="headings", height=16)
        tree_u.heading("ID",          text="ID")
        tree_u.heading("Descripcion", text="Descripción")
        tree_u.heading("Estado",      text="Estado")
        tree_u.column("ID",          width=50,  anchor="center")
        tree_u.column("Descripcion", width=420)
        tree_u.column("Estado",      width=60,  anchor="center")
        for u in UNIDADES_REF:
            tree_u.insert("", tk.END, values=(u["ID"], u["DESCRIPCION"], u["ESTADO"]))
        sb_u = ttk.Scrollbar(tab_unid, orient=tk.VERTICAL, command=tree_u.yview)
        tree_u.configure(yscrollcommand=sb_u.set)
        tree_u.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_u.pack(side=tk.RIGHT, fill=tk.Y)

        info = ttk.Frame(win, padding=6)
        info.pack(fill=tk.X)
        ttk.Label(info,
                  text="Categoría, Marca y Ubicación se toman exclusivamente de los datos del archivo.",
                  foreground="#555").pack(side=tk.LEFT)
        ttk.Button(info, text="Cerrar", command=win.destroy).pack(side=tk.RIGHT)

    # ── Conversión ────────────────────────────────────────────────────────────

    def _convert(self):
        if not PANDAS_OK or not OPENPYXL_OK:
            messagebox.showerror("Dependencias faltantes",
                "Instale las dependencias requeridas:\n"
                "  pip install pandas openpyxl\n\n"
                "Luego reinicie la aplicación.")
            return

        try:
            codigo_ini = int(self._v_codigo.get())
            sede_id    = int(self._v_sede.get())
            unidad_id  = int(self._v_unidad.get())
        except ValueError:
            messagebox.showerror("Opciones inválidas",
                "Código inicial, Sede ID y Unidad ID deben ser enteros.")
            return

        self._log("─" * 48)
        self._log("Iniciando extracción de productos…", "info")
        self._set_status("Extrayendo…")
        products: list[dict] = []

        if self._input_file:
            ext = Path(self._input_file).suffix.lower()
            try:
                if ext in (".xlsx", ".xls"):
                    products += extract_from_excel(self._input_file, log=self._log)
                elif ext == ".csv":
                    products += extract_from_csv(self._input_file, log=self._log)
                elif ext == ".txt":
                    raw = Path(self._input_file).read_text(encoding="utf-8", errors="ignore")
                    tp  = extract_from_text(raw)
                    self._log(f"TXT: {len(tp)} productos.")
                    products += tp
                elif ext in (".jpg", ".jpeg", ".png", ".bmp"):
                    products += extract_from_image(self._input_file, log=self._log)
                else:
                    self._log(f"Formato '{ext}' no soportado.", "warn")
            except Exception as exc:
                self._log(f"Error leyendo archivo: {exc}", "err")
                traceback.print_exc()

        if not self._placeholder_active:
            raw = self._txt.get("1.0", tk.END).strip()
            if raw:
                tp = extract_from_text(raw)
                self._log(f"Texto pegado: {len(tp)} productos.")
                products += tp

        if not products:
            messagebox.showwarning("Sin datos",
                "No se encontraron productos.\n\n"
                "Verifique el archivo seleccionado o el texto pegado.")
            self._set_status("Sin datos extraídos.")
            return

        self._log(f"Total extraídos: {len(products)} productos.", "ok")
        for p in products[:6]:
            brand, cat = _resolve_brand_cat(p)
            marca_disp = brand if brand != "GENERAL" else "(sin marca)"
            cat_disp   = cat   if cat   != "GENERAL" else "(sin categoria)"
            cod_disp   = p.get("codigo_origen") or "(auto)"
            self._log(f"  • [{cod_disp:<10}] {p['descripcion']:<30} "
                      f"marca={marca_disp:<14} cat={cat_disp}")
        if len(products) > 6:
            self._log(f"  … y {len(products) - 6} producto(s) más.")

        if self._v_preview.get():
            if not self._show_preview(products):
                self._log("Cancelado por el usuario.")
                self._set_status("Cancelado.")
                return

        self._log("Construyendo formato oficial…", "info")
        try:
            sheets = build_sheets(products, codigo_ini, sede_id, unidad_id)
        except Exception as exc:
            self._log(f"Error construyendo inventario: {exc}", "err")
            traceback.print_exc()
            messagebox.showerror("Error", str(exc))
            return

        n_cat   = len(sheets["CATEGORIAS"])
        n_brand = len(sheets["MARCAS"])
        n_ubic  = len(sheets["UBICACION"])
        n_unit_med = len(sheets["UNIDAD_MEDIDA"])
        self._log(f"  Categorías generadas : {n_cat}", "info")
        self._log(f"  Marcas generadas    : {n_brand}", "info")
        self._log(f"  Ubicaciones         : {n_ubic}", "info")
        self._log(f"  Presentaciones (UNIDAD_MEDIDA): {n_unit_med}", "info")

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 2010+", "*.xlsx")],
            initialfile=f"inventario_convertido_{ts}.xlsx",
            title="Guardar inventario convertido",
        )
        if not save_path:
            self._log("Guardado cancelado.")
            self._set_status("Guardado cancelado.")
            return

        try:
            write_excel(sheets, save_path)
        except Exception as exc:
            self._log(f"Error guardando: {exc}", "err")
            traceback.print_exc()
            messagebox.showerror("Error al guardar", str(exc))
            return

        n     = len(products)
        fname = Path(save_path).name
        self._log(f"✓ Guardado: {fname}", "ok")
        self._log(f"  PRODUCTOS:          {n} filas", "ok")
        self._log(f"  PRODUCTO SEDES:     {n} filas", "ok")
        self._log(f"  PRODUCTOS DETALLES: {n} filas", "ok")
        self._log(f"  UNIDADES:           {len(UNIDADES_REF)} filas (fija)", "ok")
        self._set_status(f"✓ {n} productos exportados → {fname}")

        if messagebox.askyesno("¡Éxito!",
            f"Inventario convertido y guardado.\n\n"
            f"  Productos:   {n}\n"
            f"  Categorías: {n_cat}\n"
            f"  Marcas:      {n_brand}\n"
            f"  Archivo:     {fname}\n\n"
            "¿Abrir la carpeta de destino?"):
            os.startfile(str(Path(save_path).parent))

    # ── Previsualización ──────────────────────────────────────────────────────

    def _show_preview(self, products: list[dict]) -> bool:
        win = tk.Toplevel(self)
        win.title("Previzualización de productos")
        win.geometry("920x520")
        win.grab_set()

        ttk.Label(win,
                  text=f"Se encontraron {len(products)} productos. Revise y confirme.",
                  font=("Segoe UI", 11, "bold"), padding=10).pack()

        ttk.Label(win,
                  text="Las columnas Marca y Categoría muestran los valores que se crearán "
                       "en las hojas MARCAS y CATEGORIAS del Excel.",
                  foreground="#555", font=("Segoe UI", 9)).pack()

        fr = ttk.Frame(win, padding=(8, 4, 8, 0))
        fr.pack(fill=tk.BOTH, expand=True)

        cols    = ("#", "Código", "Descripción", "Precio (Bs)", "Marca", "Categoría", "Cant.", "Unid. Alt.")
        widths  = [38, 90, 260, 88, 110, 100, 55, 70]
        anchors = ["center", "center", "w", "center", "w", "w", "center", "center"]

        tree = ttk.Treeview(fr, columns=cols, show="headings", height=14)
        for c, w, a in zip(cols, widths, anchors):
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor=a)

        sb = ttk.Scrollbar(fr, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        for i, p in enumerate(products, 1):
            brand_name, cat_name = _resolve_brand_cat(p)
            tree.insert("", tk.END, values=(
                i,
                p.get("codigo_origen") or "(auto)",
                p.get("descripcion", ""),
                f"{p.get('precio', 0):.2f}",
                brand_name,
                cat_name,
                p.get("cantidad", 1),
                len(p.get("unidades_alternas", [])),
            ))

        result = [False]

        def confirm():
            result[0] = True
            win.destroy()

        btn_fr = ttk.Frame(win, padding=8)
        btn_fr.pack()
        tk.Button(btn_fr, text="✓  Confirmar y generar Excel",
                  bg="#1F4E79", fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=12, pady=6,
                  activebackground="#2E75B6", activeforeground="white",
                  command=confirm).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_fr, text="✗  Cancelar",
                   command=win.destroy).pack(side=tk.LEFT, padx=6)

        self.wait_window(win)
        return result[0]


# ─── Entrada principal ────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not PANDAS_OK or not OPENPYXL_OK:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Dependencias faltantes",
            "Faltan dependencias criticas.\n\n"
            "Ejecute en la terminal:\n"
            "  pip install pandas openpyxl\n\n"
            "Luego reinicie la aplicacion."
        )
        root.destroy()
    else:
        app = App()
        app.mainloop()
