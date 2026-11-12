#!/usr/bin/env python3
"""
Descarga los xlsx oficiales de SUNARP (Registro de propiedad inmueble) y los
convierte a JSON local. No consulta sistemas del Estado desde la app: la app
solo lee los JSON generados aqui.

Uso:  python scripts/sunarp_a_json.py
Variables opcionales:
  SUNARP_PAGINA  URL de la pagina de gob.pe con los xlsx (por defecto, la oficial)
  SUNARP_URLS    JSON {"transferencia_propiedad": "https://...xlsx", ...} para
                 saltarse la deteccion automatica de enlaces
Tambien se lee data/sunarp/urls_manual.json si existe (mismo formato).
"""
import datetime as dt
import hashlib
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

PAGINA = os.environ.get(
    "SUNARP_PAGINA",
    "https://www.gob.pe/institucion/sunarp/informes-publicaciones/"
    "1433948-registro-de-propiedad-inmueble-sunarp",
)
# User-Agent simple (token producto/version): algunos WAF bloquean curl/* o UA con parentesis
UA = "mp-datos-abiertos/1.0"
FUENTE = "SUNARP, Registro de propiedad inmueble (Estadisticas Sunarp)"

# clave de salida -> frases que deben aparecer en el titulo o nombre del archivo
CONJUNTOS = {
    "transferencia_propiedad": ["transferencia de propiedad"],
    "compraventa_predial_sid": ["compra venta predial", "compraventa predial"],
}

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "data" / "sunarp"
RAW = SALIDA / "raw"

NIVELES = {
    "distrito": "distrito",
    "oficina registral": "oficina registral",
    "zona registral": "zona registral",
    "departamento": "departamento",
    "provincia": "provincia",
}


def normalizar(t: str) -> str:
    t = urllib.parse.unquote(t)
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def descargar(url: str, intentos: int = 3) -> bytes:
    ultimo = None
    for i in range(intentos):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=60) as r:
                return r.read()
        except (urllib.error.URLError, TimeoutError) as e:
            ultimo = e
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"No se pudo descargar {url}: {ultimo}")


def detectar_urls() -> dict:
    manual = {}
    ruta_manual = SALIDA / "urls_manual.json"
    if ruta_manual.exists():
        manual.update(json.loads(ruta_manual.read_text(encoding="utf-8")))
    if os.environ.get("SUNARP_URLS"):
        manual.update(json.loads(os.environ["SUNARP_URLS"]))
    if len(manual) == len(CONJUNTOS):
        return manual

    html = descargar(PAGINA).decode("utf-8", errors="replace")
    enlaces = re.findall(
        r"""(?:https?://|/)[^\s"'<>]+?\.xlsx""",
        html,
        flags=re.I,
    )
    base = PAGINA
    vistos, encontrados = set(), {}
    for e in enlaces:
        url = urllib.parse.urljoin(base, e.replace("&amp;", "&"))
        if url in vistos:
            continue
        vistos.add(url)
        pos = html.find(e)
        # texto propio del enlace: desde el href hasta el cierre </a> (max 300 caracteres)
        resto = html[pos: pos + 300]
        cierre = resto.lower().find("</a>")
        anchor = resto if cierre == -1 else resto[:cierre]
        contexto = normalizar(anchor + " " + url)
        for clave, frases in CONJUNTOS.items():
            if clave in manual or clave in encontrados:
                continue
            if any(normalizar(f) in contexto for f in frases):
                encontrados[clave] = url
                break  # un archivo pertenece a un solo conjunto
    manual.update(encontrados)
    faltan = [c for c in CONJUNTOS if c not in manual]
    if faltan:
        print("No se detectaron enlaces para:", ", ".join(faltan), file=sys.stderr)
        print("Enlaces xlsx vistos en la pagina:", file=sys.stderr)
        for u in vistos:
            print("  ", u, file=sys.stderr)
        print(
            "Solucion: crear data/sunarp/urls_manual.json con las URLs reales.",
            file=sys.stderr,
        )
        sys.exit(2)
    return manual


def valor(c):
    if c is None:
        return None
    if isinstance(c, (dt.datetime, dt.date)):
        return c.isoformat()
    if isinstance(c, float) and c.is_integer():
        return int(c)
    return c


def leer_hojas(ruta: Path) -> list:
    wb = load_workbook(ruta, data_only=True, read_only=True)
    hojas = []
    for ws in wb.worksheets:
        filas = [[valor(c) for c in fila] for fila in ws.iter_rows(values_only=True)]
        while filas and all(c is None for c in filas[-1]):
            filas.pop()
        ncols = max((max((i for i, c in enumerate(f) if c is not None), default=-1) + 1)
                    for f in filas) if filas else 0
        filas = [f[:ncols] for f in filas]
        hojas.append({"nombre": ws.title, "filas": filas})
    wb.close()
    return hojas


def niveles_detectados(hojas: list) -> list:
    texto = normalizar(" ".join(
        str(c) for h in hojas for f in h["filas"][:40] for c in f if isinstance(c, str)
    ))
    return sorted({etq for clave, etq in NIVELES.items() if clave in texto})


def main():
    SALIDA.mkdir(parents=True, exist_ok=True)
    RAW.mkdir(parents=True, exist_ok=True)
    urls = detectar_urls()
    ahora = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    perfil, cambios = {}, 0

    for clave in CONJUNTOS:
        url = urls[clave]
        datos = descargar(url)
        sha = hashlib.sha256(datos).hexdigest()
        destino = SALIDA / f"{clave}.json"
        previo = json.loads(destino.read_text(encoding="utf-8")) if destino.exists() else {}

        if previo.get("sha256") == sha:
            print(f"[{clave}] sin cambios")
            hojas = previo["hojas"]
        else:
            (RAW / f"{clave}.xlsx").write_bytes(datos)
            hojas = leer_hojas(RAW / f"{clave}.xlsx")
            salida = {
                "conjunto": clave,
                "fuente": FUENTE,
                "url_origen": url,
                "descargado_utc": ahora,
                "sha256": sha,
                "hojas": hojas,
            }
            destino.write_text(
                json.dumps(salida, ensure_ascii=False, separators=(",", ":")),
                encoding="utf-8",
            )
            cambios += 1
            print(f"[{clave}] actualizado ({len(hojas)} hojas)")

        perfil[clave] = {
            "url_origen": url,
            "niveles_geograficos_detectados": niveles_detectados(hojas),
            "hojas": [
                {
                    "nombre": h["nombre"],
                    "filas": len(h["filas"]),
                    "columnas": len(h["filas"][0]) if h["filas"] else 0,
                    "primeras_filas": h["filas"][:8],
                }
                for h in hojas
            ],
        }

    (SALIDA / "_perfil.json").write_text(
        json.dumps(perfil, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if cambios or not (SALIDA / "meta.json").exists():
        (SALIDA / "meta.json").write_text(
            json.dumps(
                {"fuente": FUENTE, "actualizado_utc": ahora,
                 "texto_pie": f"Datos al {ahora[:10]}, fuente: SUNARP"},
                ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    print("Listo. Revisa data/sunarp/_perfil.json para ver columnas y nivel geografico.")


if __name__ == "__main__":
    main()
