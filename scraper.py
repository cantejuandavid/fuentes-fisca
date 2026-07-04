#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scraper.py
==========
Mantiene actualizadas, desde el sitio de la DIAN, tres listas oficiales:
  - Proveedores ficticios
  - Contadores sancionados por la DIAN
  - Autorretenedores del impuesto sobre la renta

Flujo:
  1. Con Playwright (Chromium headless) abre cada página de origen UNA vez (las fuentes
     se agrupan por página; hoy: la home de la DIAN y la página de Autorretenedores).
     Se bloquean imágenes/fuentes/CSS/medios para cargar solo el DOM (más rápido).
  2. De cada fuente localiza su `selector` (un <li data-content="..."> o un <a href>)
     y obtiene el href ACTUAL del PDF (la DIAN lo cambia en cada actualización).
  3. Descarga cada PDF.
  4. Extrae la(s) tabla(s) con pdfplumber y normaliza a las columnas canónicas de la fuente.
  5. Escritura:
     - Si los DATOS cambiaron: reescribe csv (UTF-8 BOM, ';'), json y meta.
     - Si los datos NO cambiaron pero los METADATOS sí (nueva URL del PDF, cambio de
       esquema del meta): reescribe solo el meta CONSERVANDO fecha_actualizacion, que
       siempre significa "última vez que los datos cambiaron".
     - Si nada cambió: no reescribe nada (el commit del workflow es condicional).

Robustez:
  - Carga de páginas y descargas con reintentos ante fallos transitorios.
  - Cada fuente se procesa de forma INDEPENDIENTE: si una falla, las demás se actualizan
    igual; la que falla conserva su última versión válida.
  - Validación con memoria: además de mínimos absolutos, si el número de registros cae
    más del 50% respecto al meta anterior, se rechaza la extracción (protege contra PDFs
    truncados o enlaces que apuntan a otro documento).
  - Escritura atómica: genera *.tmp, valida y solo entonces reemplaza.

Códigos de salida (contrato con el workflow y el README):
  0 = todas las fuentes OK · 2 = alguna fuente con fallo CONTROLADO (ScraperError)
  1 = fallo INESPERADO (excepción no controlada en alguna fuente o al cargar páginas)

Para agregar/ajustar fuentes o columnas, edita la lista FUENTES.
"""

from __future__ import annotations

import csv
import io
import json
import logging
import re
import sys
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import requests

# ----------------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------------

DIAN_URL = "https://www.dian.gov.co/Paginas/Inicio.aspx"

# Recursos que NO necesitamos para hallar los enlaces (acelera la carga de la página).
RECURSOS_BLOQUEADOS = {"image", "font", "stylesheet", "media"}

# Reintentos ante fallos transitorios (red / render).
REINTENTOS = 3
ESPERA_REINTENTO = 4  # segundos entre intentos

# Caída máxima tolerada del nº de registros vs. la corrida anterior (ver validar()).
MAX_CAIDA_REGISTROS = 0.5

ROOT = Path(__file__).resolve().parent

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# ----------------------------------------------------------------------------
# Definición de fuentes (data-driven)
# ----------------------------------------------------------------------------
#   nombre        : etiqueta legible (para logs y meta)
#   pagina        : URL de la página donde vive el enlace (default: home de la DIAN)
#   selector      : selector CSS del elemento con el enlace; puede ser el propio <a>
#                   o un contenedor (p.ej. <li data-content="...">)
#   csv/json/meta : nombres de archivo de salida (en la raíz del repo)
#   columnas      : columnas canónicas del CSV (orden estable para Excel/VBA)
#   keywords      : por columna canónica, palabras clave para mapear el encabezado del
#                   PDF. El match se resuelve en tres pases (igualdad exacta, prefijo,
#                   subcadena) para que un encabezado como "FECHA DE RESOLUCION" gane
#                   por prefijo en Fecha y no por subcadena en Resolucion.
#   requeridas    : columnas que deben reconocerse para aceptar una fila de encabezado.
#                   Además se exige mapear >= 60% de `columnas`, para que una fila de
#                   datos con subcadenas casuales no pase por encabezado.
#   col_id        : columna identificadora; debe ser numérica (tolera puntos/guiones).
#   min_filas     : mínimo absoluto de filas de datos para aceptar la extracción.
#   min_id_ratio  : mínimo de filas cuyo col_id parece un número.
#   solo_filas_con_id : si True, descarta toda fila cuyo col_id no sea numérico
#                   (útil cuando el PDF repite título/encabezado en cada página).
FUENTES = [
    {
        "nombre": "Proveedores ficticios",
        "selector": 'li[data-content="Proveedores ficticios"]',
        "csv": "proveedores_ficticios.csv",
        "json": "proveedores_ficticios.json",
        "meta": "meta.json",  # nombre histórico (compatibilidad con la macro VBA)
        "columnas": ["NIT", "Razon_Social", "Resolucion", "Fecha", "Estado"],
        "keywords": {
            "NIT": ["nit", "identificacion", "documento", "cedula"],
            "Razon_Social": ["razon", "nombre", "social", "contribuyente", "proveedor"],
            "Resolucion": ["resolucion", "acto"],
            "Fecha": ["fecha", "ano", "vigencia"],
            "Estado": ["estado", "observacion", "situacion"],
        },
        "requeridas": ["NIT", "Razon_Social"],
        "col_id": "NIT",
        "min_filas": 5,
        "min_id_ratio": 0.5,
    },
    {
        "nombre": "Contadores sancionados por la DIAN",
        "selector": 'li[data-content="Contadores sancionados por la DIAN"]',
        "csv": "contadores_sancionados.csv",
        "json": "contadores_sancionados.json",
        "meta": "contadores_sancionados.meta.json",
        "columnas": [
            "No", "Nombre", "Cedula", "Inscripcion_Profesional", "Resolucion",
            "Sancion", "Fecha_Ejecutoria", "Vencimiento", "Autoridad",
        ],
        "keywords": {
            "No": ["no", "num", "numero", "item"],
            "Nombre": ["nombre"],
            "Cedula": ["cedula", "identificacion", "documento"],
            "Inscripcion_Profesional": ["inscripcion", "profesional", "tarjeta"],
            "Resolucion": ["resolucion", "acto"],
            "Sancion": ["sancion"],
            "Fecha_Ejecutoria": ["ejecutoria", "fecha"],
            "Vencimiento": ["vencimiento", "vence"],
            "Autoridad": ["autoridad", "sanciona", "entidad"],
        },
        "requeridas": ["Cedula", "Nombre"],
        "col_id": "Cedula",
        "min_filas": 3,
        "min_id_ratio": 0.5,
    },
    {
        "nombre": "Autorretenedores de renta",
        "pagina": "https://www.dian.gov.co/impuestos/Autorretenedores/Paginas/Autorretenedor-del-Impuesto-sobre-la-Renta.aspx",
        # "Renta" en el href evita agarrar otro documento que la DIAN suba a esa carpeta.
        "selector": 'a[href*="/impuestos/Autorretenedores/Documents/"][href*="Renta"]',
        "csv": "autorretenedores_renta.csv",
        "json": "autorretenedores_renta.json",
        "meta": "autorretenedores_renta.meta.json",
        "columnas": ["NIT", "Razon_Social", "Resolucion", "Fecha"],
        "keywords": {
            "NIT": ["nit", "identificacion", "documento"],
            "Razon_Social": ["razon", "nombre", "social"],
            "Resolucion": ["resolucion", "numero", "acto"],
            "Fecha": ["fecha"],
        },
        "requeridas": ["NIT", "Razon_Social"],
        "col_id": "NIT",
        "min_filas": 100,       # son miles; la validación con memoria cubre el resto
        "min_id_ratio": 0.9,
        # El PDF repite una fila de título y otra de encabezado en cada página:
        # quedarse solo con filas cuyo NIT sea numérico las descarta automáticamente.
        "solo_filas_con_id": True,
    },
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("dian-scraper")


class ScraperError(Exception):
    """Error controlado: aborta la fuente sin sobrescribir sus datos buenos."""


def _reintentar(fn, *, descripcion: str, intentos: int = REINTENTOS, espera: int = ESPERA_REINTENTO):
    """Ejecuta `fn` reintentando ante cualquier excepción transitoria."""
    ultimo_error: Exception | None = None
    for n in range(1, intentos + 1):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            ultimo_error = e
            if n < intentos:
                log.warning("Intento %d/%d de %s falló: %s. Reintentando en %ds...",
                            n, intentos, descripcion, e, espera)
                time.sleep(espera)
            else:
                log.error("Agotados los %d intentos de %s.", intentos, descripcion)
    raise ultimo_error  # type: ignore[misc]


# ----------------------------------------------------------------------------
# Paso 1-2: obtener los enlaces actuales (una carga por página de origen)
# ----------------------------------------------------------------------------

def _href_de(page, selector: str) -> str | None:
    """Devuelve la URL absoluta del enlace ubicado por `selector`, o None.
    Funciona tanto si el selector apunta al propio <a> como a un contenedor (p.ej. un <li>)."""
    el = page.locator(selector).first
    if el.count() == 0:
        return None

    # 1) el propio elemento es un <a> con href
    href = el.get_attribute("href")

    # 2) o contiene un <a href> (caso <li ...><a>...</a></li>)
    if not href:
        anchor = el.locator("a[href]").first
        if anchor.count() > 0:
            href = anchor.get_attribute("href")

    # 3) o trae el enlace en un atributo data-*
    if not href:
        for attr in ("data-href", "data-url", "data-link"):
            val = el.get_attribute(attr)
            if val:
                href = val
                break

    # 4) último recurso: buscar un href embebido en el HTML interno
    if not href:
        inner = el.inner_html() or ""
        m = re.search(r'href=["\']([^"\']+)["\']', inner)
        if m:
            href = m.group(1)

    if not href:
        return None
    return urljoin(page.url, href.strip())


def obtener_urls(fuentes: list[dict]) -> dict[str, str | None]:
    """Carga cada página necesaria (una sola vez por URL) y devuelve {nombre_fuente: url_pdf|None}.

    Las fuentes se agrupan por su página de origen: si una página falla, las fuentes de las
    otras páginas igual se resuelven."""
    from playwright.sync_api import sync_playwright

    # Agrupar fuentes por su página de origen (por defecto, la home de la DIAN).
    paginas: dict[str, list[dict]] = {}
    for f in fuentes:
        paginas.setdefault(f.get("pagina", DIAN_URL), []).append(f)

    resultado: dict[str, str | None] = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            for pagina, fs in paginas.items():
                log.info("Abriendo %s con Playwright...", pagina)

                def intento(pagina=pagina, fs=fs) -> dict[str, str | None]:
                    context = browser.new_context(user_agent=USER_AGENT, locale="es-CO")
                    context.route(
                        "**/*",
                        lambda route: (
                            route.abort()
                            if route.request.resource_type in RECURSOS_BLOQUEADOS
                            else route.continue_()
                        ),
                    )
                    page = context.new_page()
                    try:
                        page.goto(pagina, wait_until="domcontentloaded", timeout=60_000)
                        urls: dict[str, str | None] = {}
                        for f in fs:
                            sel = f["selector"]
                            try:
                                page.wait_for_selector(sel, timeout=30_000)
                            except Exception:
                                log.warning("No apareció el selector de '%s' (%s).", f["nombre"], sel)
                            urls[f["nombre"]] = _href_de(page, sel)
                        return urls
                    finally:
                        context.close()

                try:
                    urls = _reintentar(intento, descripcion=f"cargar {pagina}")
                except Exception as e:  # noqa: BLE001
                    log.error("No se pudo cargar %s: %s", pagina, e)
                    urls = {f["nombre"]: None for f in fs}
                resultado.update(urls)

            for nombre, url in resultado.items():
                log.info("Enlace [%s]: %s", nombre, url or "NO ENCONTRADO")
            return resultado
        finally:
            browser.close()


# ----------------------------------------------------------------------------
# Paso 3: descargar un PDF
# ----------------------------------------------------------------------------

def descargar_pdf(url: str) -> bytes:
    log.info("Descargando PDF: %s", url)
    headers = {"User-Agent": USER_AGENT, "Accept": "application/pdf,*/*"}

    def intento() -> bytes:
        with requests.Session() as s:
            resp = s.get(url, headers=headers, timeout=120, allow_redirects=True)
        resp.raise_for_status()
        content = resp.content
        ctype = resp.headers.get("Content-Type", "")
        if not (content[:5] == b"%PDF-" or "pdf" in ctype.lower()):
            raise ScraperError(
                f"El contenido descargado no parece un PDF (Content-Type={ctype!r}, "
                f"primeros bytes={content[:8]!r})."
            )
        return content

    content = _reintentar(intento, descripcion="descargar el PDF")
    log.info("PDF descargado: %d bytes", len(content))
    return content


# ----------------------------------------------------------------------------
# Paso 4: extraer y normalizar la tabla
# ----------------------------------------------------------------------------

def _norm(s: str) -> str:
    """Minúsculas, sin acentos ni puntuación, sin espacios extra (para comparar encabezados)."""
    s = (s or "").strip().lower()
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _mapear_encabezado(headers: list[str], keywords: dict[str, list[str]],
                       requeridas: list[str], columnas: list[str]) -> dict[int, str] | None:
    """Devuelve {indice_columna -> nombre_canonico} si la fila parece un encabezado real.

    El match se resuelve en tres pases con prioridad decreciente — igualdad exacta,
    prefijo, subcadena — para que "FECHA DE RESOLUCION" gane por prefijo en Fecha y no
    por subcadena en Resolucion (robusto ante reordenamiento de columnas del PDF).

    Se acepta solo si (a) todas las `requeridas` quedaron mapeadas y (b) se mapeó al
    menos el 60% de `columnas`; esto evita que una fila de DATOS con subcadenas
    casuales ("...CORAZON...", "...DOCUMENTO...") pase por encabezado.
    """
    norm = [_norm(h) for h in headers]
    mapeo: dict[int, str] = {}
    usados: set[str] = set()

    for pase in ("igual", "prefijo", "sub"):
        for idx, h in enumerate(norm):
            if not h or idx in mapeo:
                continue
            for canon, kws in keywords.items():
                if canon in usados:
                    continue
                hit = False
                for kw in kws:
                    if pase == "igual":
                        hit = (h == kw)
                    elif pase == "prefijo":
                        hit = len(kw) >= 4 and h.startswith(kw)
                    else:  # subcadena
                        hit = len(kw) >= 4 and kw in h
                    if hit:
                        break
                if hit:
                    mapeo[idx] = canon
                    usados.add(canon)
                    break

    min_mapeadas = max(len(requeridas), (3 * len(columnas) + 4) // 5)  # ~60%
    if all(c in usados for c in requeridas) and len(mapeo) >= min_mapeadas:
        return mapeo
    return None


def _limpiar(celda) -> str:
    if celda is None:
        return ""
    return re.sub(r"\s+", " ", str(celda)).strip()


def _parece_id(valor: str) -> bool:
    """True si el valor es un identificador numérico plausible (NIT/cédula).
    Tolera puntos, espacios y guiones como separadores, pero exige que el resto
    sean SOLO dígitos (un título con una fecha embebida ya no pasa)."""
    v = re.sub(r"[.\s\-]", "", valor or "")
    return v.isdigit() and 5 <= len(v) <= 15


def extraer_tabla(pdf_bytes: bytes, fuente: dict) -> list[dict]:
    """Extrae filas de todas las tablas del PDF y las normaliza a las columnas de la fuente."""
    import pdfplumber

    columnas = fuente["columnas"]
    keywords = fuente["keywords"]
    requeridas = fuente["requeridas"]
    col_id = fuente["col_id"]
    solo_con_id = fuente.get("solo_filas_con_id", False)

    registros: list[dict] = []
    mapeo_global: dict[int, str] | None = None

    log.info("[%s] Extrayendo tablas con pdfplumber...", fuente["nombre"])
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pnum, page in enumerate(pdf.pages, start=1):
            tablas = page.extract_tables() or []
            for tabla in tablas:
                if not tabla:
                    continue
                # Buscar la fila de encabezado en las primeras filas (hay PDF con una fila
                # de título antes del encabezado, y el encabezado se repite por página).
                inicio = 0
                for hi in range(min(5, len(tabla))):
                    posible = _mapear_encabezado(
                        [_limpiar(c) for c in tabla[hi]], keywords, requeridas, columnas
                    )
                    if posible:
                        if mapeo_global is not None and posible != mapeo_global:
                            log.warning(
                                "[%s] El mapeo de columnas cambió en la página %d: %s -> %s",
                                fuente["nombre"], pnum,
                                sorted(mapeo_global.items()), sorted(posible.items()),
                            )
                        if len(posible) < len(columnas):
                            faltan = [c for c in columnas if c not in posible.values()]
                            log.warning(
                                "[%s] Encabezado detectado sin mapear todas las columnas; "
                                "faltan: %s (quedarán vacías).", fuente["nombre"], faltan,
                            )
                        mapeo_global = posible
                        inicio = hi + 1
                        break

                mapeo = mapeo_global
                for fila in tabla[inicio:]:
                    celdas = [_limpiar(c) for c in fila]
                    if not any(celdas):
                        continue

                    reg = {col: "" for col in columnas}
                    if mapeo:
                        for idx, canon in mapeo.items():
                            if idx < len(celdas):
                                reg[canon] = celdas[idx]
                    else:
                        # Fallback posicional: usa el orden de `columnas`.
                        for i, canon in enumerate(columnas):
                            if i < len(celdas):
                                reg[canon] = celdas[i]

                    if solo_con_id:
                        # Mantener solo filas cuyo identificador sea numérico
                        # (descarta filas de título/encabezado repetidas por página).
                        if not _parece_id(reg.get(col_id, "")):
                            continue
                    else:
                        # Descartar filas no-datos: todas las columnas requeridas vacías.
                        if all(not reg.get(c) for c in requeridas):
                            continue
                    registros.append(reg)

            log.info("[%s] Página %d procesada (%d registros).",
                     fuente["nombre"], pnum, len(registros))

    return registros


def _leer_meta_anterior(fuente: dict) -> dict | None:
    meta_path = ROOT / fuente["meta"]
    if not meta_path.exists():
        return None
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def validar(registros: list[dict], fuente: dict) -> None:
    nombre = fuente["nombre"]
    min_filas = fuente["min_filas"]
    col_id = fuente["col_id"]
    min_ratio = fuente["min_id_ratio"]

    if len(registros) < min_filas:
        raise ScraperError(
            f"Se extrajeron muy pocas filas ({len(registros)} < {min_filas}). "
            "Posible cambio de estructura del PDF; se conserva la versión anterior."
        )
    con_id = sum(1 for r in registros if _parece_id(r.get(col_id, "")))
    ratio = con_id / len(registros)
    if ratio < min_ratio:
        raise ScraperError(
            f"Solo {ratio:.0%} de las filas tienen '{col_id}' plausible "
            f"(mínimo {min_ratio:.0%}). Se conserva la versión anterior."
        )

    # Validación con memoria: una caída fuerte frente a la corrida anterior sugiere un
    # PDF truncado o un enlace que apunta a otro documento. Si la reducción es legítima
    # (la DIAN depuró la lista), borra el meta de la fuente o ajusta MAX_CAIDA_REGISTROS.
    meta_ant = _leer_meta_anterior(fuente)
    prev = (meta_ant or {}).get("num_registros")
    if isinstance(prev, int) and prev > 0 and len(registros) < prev * (1 - MAX_CAIDA_REGISTROS):
        raise ScraperError(
            f"Caída sospechosa de registros: {len(registros)} vs {prev} de la corrida "
            f"anterior (>{MAX_CAIDA_REGISTROS:.0%}). Se conserva la versión anterior."
        )

    log.info("[%s] Validación OK: %d filas, %.0f%% con '%s' plausible.",
             nombre, len(registros), ratio * 100, col_id)


# ----------------------------------------------------------------------------
# Paso 5: escribir salidas (atómico y solo si algo cambió)
# ----------------------------------------------------------------------------

def _csv_str(registros: list[dict], columnas: list[str]) -> str:
    """Serializa los registros a CSV (separador ';', salto '\\n' determinista)."""
    buf = io.StringIO()
    writer = csv.DictWriter(
        buf, fieldnames=columnas, delimiter=";",
        extrasaction="ignore", lineterminator="\n",
    )
    writer.writeheader()
    writer.writerows(registros)
    return buf.getvalue()


def _escribir_atomico(path: Path, contenido: str, encoding: str = "utf-8") -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(contenido, encoding=encoding, newline="")
    tmp.replace(path)


def escribir_salidas(registros: list[dict], url_pdf: str, fuente: dict) -> bool:
    """Escribe las salidas de la fuente. Devuelve True si algo cambió en disco.

    - Datos nuevos  -> reescribe csv + json + meta (fecha_actualizacion = ahora).
    - Solo metadatos (nueva URL de PDF, cambio de esquema) -> reescribe SOLO el meta,
      conservando fecha_actualizacion (= último cambio real de los datos).
    - Nada cambió   -> no toca ningún archivo.
    """
    nombre = fuente["nombre"]
    columnas = fuente["columnas"]
    csv_path = ROOT / fuente["csv"]
    json_path = ROOT / fuente["json"]
    meta_path = ROOT / fuente["meta"]

    nuevo_csv = _csv_str(registros, columnas)

    csv_cambio = True
    if csv_path.exists():
        try:
            actual = csv_path.read_text(encoding="utf-8-sig")
        except OSError:
            actual = None
        csv_cambio = (actual != nuevo_csv)

    # Esquema ÚNICO del meta para todas las fuentes (ver README).
    meta_nuevo = {
        "fuente": nombre,
        "fecha_actualizacion": datetime.now(timezone.utc).isoformat(),
        "url_pdf": url_pdf,
        "num_registros": len(registros),
        "columnas": columnas,
        "pagina": fuente.get("pagina", DIAN_URL),
    }

    if not csv_cambio:
        meta_viejo = _leer_meta_anterior(fuente)
        if meta_viejo and meta_viejo.get("fecha_actualizacion"):
            # Los datos no cambiaron: la fecha del último cambio real se conserva.
            meta_nuevo["fecha_actualizacion"] = meta_viejo["fecha_actualizacion"]
        if meta_viejo == meta_nuevo:
            log.info("[%s] Datos y metadatos idénticos (%d filas): no se reescribe nada.",
                     nombre, len(registros))
            return False
        _escribir_atomico(meta_path, json.dumps(meta_nuevo, ensure_ascii=False, indent=2))
        log.info("[%s] Datos sin cambios; metadatos actualizados (%s).", nombre, meta_path.name)
        return True

    _escribir_atomico(csv_path, nuevo_csv, encoding="utf-8-sig")  # BOM para Excel
    _escribir_atomico(json_path, json.dumps(registros, ensure_ascii=False, indent=2))
    _escribir_atomico(meta_path, json.dumps(meta_nuevo, ensure_ascii=False, indent=2))
    log.info("[%s] Archivos actualizados: %s, %s, %s",
             nombre, csv_path.name, json_path.name, meta_path.name)
    return True


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------

def procesar_fuente(fuente: dict, url: str | None) -> str:
    """Procesa una fuente. Devuelve 'ok', 'controlado' o 'inesperado'."""
    nombre = fuente["nombre"]
    try:
        if not url:
            raise ScraperError(f"No se encontró el enlace de '{nombre}' en la página.")
        pdf = descargar_pdf(url)
        registros = extraer_tabla(pdf, fuente)
        validar(registros, fuente)
        cambio = escribir_salidas(registros, url, fuente)
        log.info("[%s] %s (%d registros).",
                 nombre, "ACTUALIZADO" if cambio else "sin cambios", len(registros))
        return "ok"
    except ScraperError as e:
        log.error("[%s] Fallo controlado: %s", nombre, e)
        log.error("[%s] Se conserva la última versión válida (no se sobrescribió).", nombre)
        return "controlado"
    except Exception as e:  # noqa: BLE001
        log.exception("[%s] Fallo inesperado: %s", nombre, e)
        log.error("[%s] Se conserva la última versión válida (no se sobrescribió).", nombre)
        return "inesperado"


def main() -> int:
    try:
        urls = obtener_urls(FUENTES)
    except Exception as e:  # noqa: BLE001
        log.exception("No se pudieron cargar las páginas de la DIAN: %s", e)
        log.error("Se conservan todas las versiones válidas anteriores.")
        return 1

    estados = {f["nombre"]: procesar_fuente(f, urls.get(f["nombre"])) for f in FUENTES}
    fallidas = {n: e for n, e in estados.items() if e != "ok"}

    if fallidas:
        log.error("Fuentes con error: %s.",
                  ", ".join(f"{n} ({e})" for n, e in fallidas.items()))
        return 1 if "inesperado" in fallidas.values() else 2
    log.info("Proceso completado: todas las fuentes OK.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
