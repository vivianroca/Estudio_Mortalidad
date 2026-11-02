"""Descarga el Excel anual de defunciones del año indicado (2009–2024)."""

import argparse
import re
import unicodedata
import time
from io import BytesIO
from pathlib import Path
from urllib.parse import urljoin, urlparse
from zipfile import BadZipFile, ZipFile

import requests
from bs4 import BeautifulSoup

ENDPOINT = "https://www.ine.gob.gt/bdatos_cargar.php"
PARAMS = {
    "categoria": "82",
    "periodo": "1",
    "dire": "https://www.ine.gob.gt/sistema/",
}
TITULO = "Principales resultados: Defunciones"
RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "data" / "raw_data"
ERRORES = (requests.RequestException, ValueError, BadZipFile, OSError)


def obtener(url, **kwargs):
    """Hasta tres intentos totales para fallos de conexión o HTTP temporales."""
    for intento in range(3):
        try:
            response = requests.get(url, timeout=30, **kwargs)
            # Detectar el bloqueo antes del estado HTTP y de cualquier reintento.
            tipo = response.headers.get("Content-Type", "").lower()
            inicio = response.content.lstrip()[:100].lower()
            if "html" in tipo or inicio.startswith((b"<!doctype html", b"<html")):
                pagina = response.content[:200000].lower()
                if any(marca in pagina for marca in (
                    b"radware", b"captcha.perfdrive.com", b"radwarebotmanager",
                )):
                    raise ValueError(
                        "El INE devolvió una página CAPTCHA de Radware. "
                        "Se detuvo la ejecución sin reintentos adicionales. "
                        "Intenta ejecutar el script más tarde desde tu terminal."
                    )
            response.raise_for_status()
            return response
        except requests.RequestException as error:
            status = error.response.status_code if error.response is not None else None
            temporal = isinstance(error, (requests.ConnectionError, requests.Timeout)) or (
                status in {408, 429} or (status is not None and 500 <= status < 600)
            )
            if not temporal or intento == 2:
                raise
            print(f"  Error temporal; reintento {intento + 2}/3: {error}", flush=True)
            time.sleep(2 * (intento + 1))


def validar_excel(content, extension, content_type=""):
    if content.lstrip().lower().startswith(
        (b"<!doctype html", b"<html")
    ):
        raise ValueError("El servidor devolvió HTML en lugar de un Excel.")
    # Algunos servidores anuncian text/html incluso para archivos binarios.
    # La decisión se basa en el contenido, no solo en Content-Type.
    if extension == ".xls":
        # Firma OLE y nombre del flujo del libro binario Excel.
        ole = content.startswith(bytes.fromhex("D0CF11E0A1B11AE1")) and any(
            name.encode("utf-16le") in content for name in ("Workbook", "Book")
        )
        # Excel antiguo también puede ser BIFF sin contenedor OLE.
        biff = False
        if content[:2] in (b"\x09\x00", b"\x09\x02", b"\x09\x04", b"\x09\x08"):
            offset = 0
            while offset + 4 <= len(content):
                record = int.from_bytes(content[offset:offset + 2], "little")
                size = int.from_bytes(content[offset + 2:offset + 4], "little")
                offset += 4 + size
                if offset > len(content):
                    break
                if record == 0x000A and size == 0:
                    biff = True
                    break
        if not (ole or biff):
            raise ValueError("Formato Excel inválido: no se reconoce XLS OLE/BIFF.")
    elif extension == ".xlsx":
        with ZipFile(BytesIO(content)) as workbook:
            if not {"[Content_Types].xml", "xl/workbook.xml"} <= set(workbook.namelist()):
                raise ValueError("El archivo no contiene un libro XLSX.")
            if workbook.testzip() is not None:
                raise ValueError("El XLSX está dañado.")
    else:
        raise ValueError("Extensión distinta de XLS/XLSX.")


def normalizar(texto):
    texto = unicodedata.normalize("NFKC", texto)
    texto = " ".join(texto.split())
    return re.sub(r"\s*:\s*", ": ", texto).casefold()


def localizar_url(anio):
    response = obtener(ENDPOINT, params={**PARAMS, "anio": str(anio)})
    print(f"  [{anio}] Consulta: {response.url} | HTTP {response.status_code} | "
          f"{response.headers.get('Content-Type', '')} | {len(response.content)} bytes",
          flush=True)
    soup = BeautifulSoup(response.content, "html.parser")
    if not soup.find() or soup.find("html") and not soup.find(["a", "select", "table"]):
        raise ValueError("Respuesta HTML inesperada del endpoint: "
                         + soup.get_text(" ", strip=True)[:250])

    # Coincidencia completa; tolerar diferencias de espacios, mayúsculas y Unicode.
    etiquetas = [tag for tag in soup.find_all(True)
                 if normalizar(tag.get_text(" ", strip=True)) == normalizar(TITULO)]
    # Conservar solo el elemento más interno con el título completo.
    etiquetas = [tag for tag in etiquetas if not any(
        normalizar(child.get_text(" ", strip=True)) == normalizar(TITULO)
        for child in tag.find_all(True)
    )]
    urls = set()
    for etiqueta in etiquetas:
        enlace = etiqueta if etiqueta.name == "a" else etiqueta.find_parent("a")
        contenedor = enlace or etiqueta.find_parent("tr") or etiqueta.parent
        candidatos = [contenedor, *contenedor.find_all(["a", "button", "input"])]
        for candidato in candidatos:
            referencias = [candidato.get("href", "")]
            # Leer literales de URL; no ejecutar JavaScript ni inferir nombres.
            referencias += re.findall(r"['\"]([^'\"]+)['\"]", candidato.get("onclick", ""))
            for referencia in referencias:
                url = urljoin(response.url, referencia.strip())
                if Path(urlparse(url).path).suffix.lower() in {".xls", ".xlsx"}:
                    urls.add(url)
        print(f"  [{anio}] Elemento asociado: {str(contenedor)[:1600]}", flush=True)
    if not urls:
        candidatos = [str(tag)[:500] for tag in soup.find_all(["a", "button", "td"])
                      if "defunc" in normalizar(tag.get_text(" ", strip=True))]
        print(f"  [{anio}] Enlace no encontrado. Títulos coincidentes: {len(etiquetas)}. "
              f"Elementos con 'defunc': {candidatos[:8]}", flush=True)
        return None
    if len(urls) != 1:
        raise ValueError(f"Enlace ambiguo: {sorted(urls)}")
    url = urls.pop()
    if urlparse(url).hostname not in {"www.ine.gob.gt", "ine.gob.gt"}:
        raise ValueError("El enlace no pertenece al sitio oficial www.ine.gob.gt.")
    return url


def procesar_anio(anio):
    # Revisar ambos formatos y salir antes de consultar la red si ya existe.
    for extension in (".xls", ".xlsx"):
        ruta = DESTINO / f"defunciones_{anio}{extension}"
        if ruta.is_file():
            content = ruta.read_bytes()
            try:
                validar_excel(content, extension)
            except (ValueError, BadZipFile) as error:
                print(f"Archivo local inválido {ruta.name}: {error}", flush=True)
                continue
            print(f"Ya existe un Excel válido: {ruta} ({len(content)} bytes). "
                  "No se sobrescribió.")
            return

    url = localizar_url(anio)
    if url is None:
        raise ValueError(f"No se encontró el Excel de defunciones de {anio}.")
    print(f"URL localizada: {url}", flush=True)
    extension = Path(urlparse(url).path).suffix.lower()
    ruta = DESTINO / f"defunciones_{anio}{extension}"
    archivo = obtener(url)
    validar_excel(archivo.content, extension, archivo.headers.get("Content-Type", ""))
    DESTINO.mkdir(parents=True, exist_ok=True)
    # Escribir únicamente después de validar; sustituir solo archivos inválidos.
    temporal = ruta.with_suffix(ruta.suffix + ".part")
    try:
        temporal.write_bytes(archivo.content)
        temporal.replace(ruta)
    finally:
        temporal.unlink(missing_ok=True)
    print(f"Nombre: {ruta.name}")
    print(f"Tamaño descargado: {len(archivo.content)} bytes")
    print(f"Guardado en: {ruta}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("anio", type=int, help="Año entre 2009 y 2024, inclusive")
    args = parser.parse_args()
    if not 2009 <= args.anio <= 2024:
        parser.error("anio debe estar comprendido entre 2009 y 2024.")
    try:
        procesar_anio(args.anio)
    except ERRORES as error:
        raise SystemExit(f"Error: {error}")


if __name__ == "__main__":
    main()
