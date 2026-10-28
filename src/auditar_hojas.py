"""Audita exclusivamente nombres de hojas de los Excel de 2009 a 2024."""

import csv
import importlib
import unicodedata
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
ENTRADA = RAIZ / "data" / "raw_data"
REPORTES = RAIZ / "reports"
ANIOS = tuple(range(2009, 2025))


def leer_nombres(ruta):
    if ruta.suffix == ".xls":
        xlrd = importlib.import_module("xlrd")
        libro = xlrd.open_workbook(str(ruta), on_demand=True)
        try:
            return libro.sheet_names()
        finally:
            libro.release_resources()
    openpyxl = importlib.import_module("openpyxl")
    libro = openpyxl.load_workbook(ruta, read_only=True, keep_links=False)
    try:
        return libro.sheetnames
    finally:
        libro.close()


def normalizar(nombre):
    # Solo diferencias de caja, espacios y marcas diacríticas.
    texto = unicodedata.normalize("NFD", nombre.casefold())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return " ".join(texto.split())


def listar_anios(anios):
    return ", ".join(map(str, sorted(anios))) or "ninguno"


def main():
    hojas, errores, archivos = {}, {}, {}
    for anio in ANIOS:
        candidatos = [ENTRADA / f"{anio}{ext}" for ext in (".xls", ".xlsx")
                      if (ENTRADA / f"{anio}{ext}").is_file()]
        archivos[anio] = ", ".join(p.name for p in candidatos) or "—"
        if len(candidatos) != 1:
            errores[anio] = ("Archivo faltante" if not candidatos else
                            "Más de un archivo para el año; ninguno seleccionado")
            continue
        try:
            nombres = leer_nombres(candidatos[0])
            if not nombres or not all(isinstance(n, str) for n in nombres):
                raise ValueError("No se pudieron extraer nombres de hojas válidos")
            hojas[anio] = nombres
        except Exception as error:
            # Un libro ilegible no impide auditar los demás años.
            errores[anio] = f"{type(error).__name__}: {error}"

    presencia = {}
    for anio, nombres in hojas.items():
        for nombre in nombres:
            presencia.setdefault(nombre, set()).add(anio)
    orden = sorted(presencia, key=lambda n: (-len(presencia[n]), n))
    todos = [n for n in orden if len(presencia[n]) == len(ANIOS)]
    varios = [n for n in orden if 1 < len(presencia[n]) < len(ANIOS)]
    exclusivos = [n for n in orden if len(presencia[n]) == 1]
    evaluados = set(hojas)
    no_evaluados = set(ANIOS) - evaluados
    REPORTES.mkdir(parents=True, exist_ok=True)

    inventario = REPORTES / "inventario_hojas.txt"
    with inventario.open("w", encoding="utf-8") as salida:
        for anio in ANIOS:
            salida.write(f"AÑO: {anio}\nArchivo: {archivos[anio]}\n")
            if anio in errores:
                salida.write(f"Estado: NO EVALUADO\nMotivo: {errores[anio]}\n\n")
            else:
                salida.write(f"Total de hojas: {len(hojas[anio])}\n\n")
                for numero, nombre in enumerate(hojas[anio], 1):
                    salida.write(f"{numero}. {nombre}\n")
                salida.write("\n")

    comparacion = REPORTES / "comparacion_hojas.csv"
    with comparacion.open("w", newline="", encoding="utf-8-sig") as salida:
        writer = csv.writer(salida)
        writer.writerow(["Hoja", *ANIOS, "total_anios"])
        for nombre in orden:
            writer.writerow([nombre, *[
                int(anio in presencia[nombre]) if anio in hojas else "NO EVALUADO"
                for anio in ANIOS
            ], len(presencia[nombre])])

    resumen = REPORTES / "resumen_hojas.txt"
    with resumen.open("w", encoding="utf-8") as salida:
        salida.write("A. Resumen general\n")
        salida.write(f"Archivos esperados: {len(ANIOS)}\n")
        salida.write(f"Archivos procesados correctamente: {len(hojas)}\n")
        salida.write(f"Total de nombres de hojas diferentes: {len(presencia)}\n")
        salida.write("Archivos faltantes o con errores:\n")
        for anio, error in errores.items():
            salida.write(f"- {anio} ({archivos[anio]}): {error}\n")
        if not errores:
            salida.write("Ninguno.\n")
        salida.write("\n1 = presente; 0 = ausente en un libro evaluado; "
                     "NO EVALUADO = año sin observación.\n")
        salida.write("total_anios cuenta únicamente presencias comprobadas. "
                     "Las exclusividades son provisionales si hay años no evaluados.\n")
        salida.write("\nB. Hojas presentes en todos los años\n")
        for nombre in todos:
            salida.write(f"- {nombre}\n")
        if not todos:
            salida.write("Ninguna comprobada en los 16 años.\n")
        salida.write("\nC. Hojas presentes en varios años\n")
        for nombre in varios:
            salida.write(f"- {nombre}\n  Cantidad de años: {len(presencia[nombre])}\n"
                         f"  Presente: {listar_anios(presencia[nombre])}\n"
                         f"  Ausente: {listar_anios(evaluados - presencia[nombre])}\n"
                         f"  No evaluado: {listar_anios(no_evaluados)}\n")
        if not varios:
            salida.write("Ninguna observada.\n")
        salida.write("\nD. Hojas exclusivas\n")
        for nombre in exclusivos:
            salida.write(f"- {nombre}: {listar_anios(presencia[nombre])}\n")
        if not exclusivos:
            salida.write("Ninguna observada.\n")
        salida.write("\nE. Posibles equivalencias\n")
        grupos = {}
        for nombre in orden:
            grupos.setdefault(normalizar(nombre), []).append(nombre)
        equivalencias = [grupo for grupo in grupos.values() if len(grupo) > 1]
        for grupo in equivalencias:
            salida.write("Grupo para revisión manual:\n")
            for nombre in grupo:
                salida.write(f"- {nombre!r}: {listar_anios(presencia[nombre])}\n")
        if not equivalencias:
            salida.write("Ninguna detectada.\n")

    print(f"Archivos procesados: {len(hojas)}/{len(ANIOS)}")
    print(f"Nombres de hojas diferentes: {len(presencia)}")
    print(f"Nombres presentes en los 16 años: {len(todos)}")
    print(f"Nombres presentes en varios años: {len(varios)}")
    print(f"Nombres exclusivos: {len(exclusivos)}")
    for anio, error in errores.items():
        print(f"NO EVALUADO {anio}: {error}")
    for reporte in (inventario, comparacion, resumen):
        print(f"Reporte: {reporte}")


if __name__ == "__main__":
    main()
