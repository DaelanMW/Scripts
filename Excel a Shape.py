import argparse
import datetime
import math
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

def _instalar_dependencias():
    """Instala lo que falte usando el mismo Python que ejecuta el script."""
    import importlib
    import subprocess
    faltan = []
    for modulo, paquete in (("pandas", "pandas"), ("openpyxl", "openpyxl"),
                            ("pyproj", "pyproj"), ("shapefile", "pyshp")):
        try:
            importlib.import_module(modulo)
        except ImportError:
            faltan.append(paquete)
    if not faltan:
        return
    print("Faltan librerías:", ", ".join(faltan))
    print(f"Instalando con: {sys.executable} -m pip install {' '.join(faltan)}\n")
    r = subprocess.call([sys.executable, "-m", "pip", "install", *faltan])
    if r != 0:
        sys.exit("\nNo se pudo instalar automáticamente. Ejecutá manualmente:\n"
                 f'    "{sys.executable}" -m pip install {" ".join(faltan)}')
    print("\nInstalación lista.\n")


_instalar_dependencias()

import pandas as pd
import shapefile  # pyshp
from pyproj import CRS, Transformer

# ============================================================== constantes
SISTEMAS = {
    "wgs84": dict(nombre="WGS84", geo=4326, proj=None),
    "posgar07": dict(nombre="POSGAR 2007", geo=5340, proj=[5343 + i for i in range(7)]),
    "posgar94": dict(nombre="POSGAR 94", geo=4694, proj=[22181 + i for i in range(7)]),
    "inchauspe": dict(nombre="Campo Inchauspe", geo=4221, proj=[22191 + i for i in range(7)]),
}
# Caja aproximada de Argentina (lon_min, lon_max, lat_min, lat_max) para validar/deducir
CAJA_ARG = (-76.0, -52.0, -57.0, -20.0)

VACIOS = {"", "nan", "none", "null", "nd", "s/d", "sd", "n/a", "na", "-", "--",
          "#n/a", "#value!", "s/n", "nat"}


# ============================================================== utilidades
def sin_acentos(t):
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")


def norm(t):
    """Normaliza un encabezado: separa camelCase/números, minúsculas, sin acentos ni símbolos."""
    t = str(t)
    t = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", t)
    t = re.sub(r"(?<=[A-Za-z])(?=\d)", " ", t)
    t = re.sub(r"(?<=\d)(?=[A-Za-z])", " ", t)
    t = sin_acentos(t).lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def es_vacio(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str):
        return v.strip().lower() in VACIOS
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


_RE_NUM = re.compile(r"^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$")


def parse_numero(v):
    """Convierte a float admitiendo coma decimal, separadores de miles y unidades (m, mts)."""
    if es_vacio(v) or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    if not isinstance(v, str):
        try:
            x = float(v)
            return x if math.isfinite(x) else None
        except (TypeError, ValueError):
            return None
    s = v.strip().replace("\u00a0", " ").replace("\u202f", " ")
    s = s.replace("−", "-").replace("–", "-").replace("—", "-")
    s = re.sub(r"\s*(mts?|metros?|m)\.?$", "", s, flags=re.I)
    if re.fullmatch(r"[-+]?\d{1,3}( \d{3})+([.,]\d+)?", s):
        s = s.replace(" ", "")
    if " " in s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if re.fullmatch(r"[-+]?\d{1,3}(,\d{3}){2,}", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif s.count(".") > 1:
        if re.fullmatch(r"[-+]?\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
        else:
            return None
    if _RE_NUM.match(s):
        x = float(s)
        return x if math.isfinite(x) else None
    return None


# ============================================================== ángulos / DMS
def _compacto(x):
    """DDMMSS(.s), DDDMMSS(.s) o DDMM(.m) escritos como un único número."""
    a = abs(x)
    ent = int(a)
    frac = a - ent
    n = len(str(ent))
    if n in (6, 7):
        ss = ent % 100 + frac
        mm = (ent // 100) % 100
        dd = ent // 10000
    elif n in (4, 5):
        mm = ent % 100 + frac
        ss = 0
        dd = ent // 100
    else:
        return None
    if mm >= 60 or ss >= 60 or dd > 180:
        return None
    return dd + mm / 60 + ss / 3600


def _armar(partes, hem):
    if hem:
        hem = {"O": "W"}.get(hem.upper(), hem.upper())
    if not 1 <= len(partes) <= 3:
        return None, None
    try:
        nums = [float(p) for p in partes]
    except ValueError:
        return None, None
    neg = partes[0].startswith("-") or math.copysign(1, nums[0]) < 0 or hem in ("S", "W")
    d = abs(nums[0])
    if len(nums) == 1:
        if d >= 1000:
            val = _compacto(d)
            if val is None:
                return None, None
        elif d > 180:
            return None, None
        else:
            val = d
    else:
        if d != int(d):
            return None, None
        m = nums[1]
        if m < 0 or m >= 60:
            return None, None
        val = d + m / 60
        if len(nums) == 3:
            s = nums[2]
            if m != int(m) or s < 0 or s >= 60:
                return None, None
            val += s / 3600
    return (-val if neg else val), hem


_PALABRAS_HEM = [
    (r"\bsur\b|\bsouth\b", "S"),
    (r"\bnorte\b|\bnorth\b", "N"),
    (r"\boeste\b|\bwest\b", "W"),
    (r"\beste\b|\beast\b", "E"),
]
_RE_LETRAS_DMS = re.compile(
    r"\s*([NSEWOnsewo])?[\s:=]*([-+]?\d+(?:[.,]\d+)?)\s*[dD]\s*"
    r"(?:(\d+(?:[.,]\d+)?)\s*[mM]\s*(?:(\d+(?:[.,]\d+)?)\s*[sS]?)?)?\s*([NSEWO])?\s*")


def parse_angulo_texto(s):
    """Texto -> (grados_decimales, hemisferio|None). Devuelve (None, None) si no se entiende."""
    t = s.strip()
    if not t:
        return None, None
    t = t.replace("−", "-").replace("–", "-").replace("—", "-")
    t = re.sub(r"[º˚°]", "°", t)
    t = re.sub(r"[′’´`‘]", "'", t)
    t = re.sub(r"[″”“]", '"', t)
    t = t.replace("''", '"')
    hem = None
    for pat, h in _PALABRAS_HEM:
        if re.search(pat, t, re.I):
            hem = h
            t = re.sub(pat, " ", t, flags=re.I)
    # 31d38m12.5s S
    m = _RE_LETRAS_DMS.fullmatch(t)
    if m:
        partes = [p.replace(",", ".") for p in (m.group(2), m.group(3), m.group(4)) if p]
        return _armar(partes, m.group(1) or m.group(5) or hem)
    # hemisferio como prefijo o sufijo
    m = re.match(r"^\s*([NSEWOnsewo])[\s:=.]*(?=[-+]?\d)", t)
    if m:
        hem = hem or m.group(1).upper()
        t = t[m.end():]
    else:
        m = re.search(r"(?<=[\d'\"°.\s])\s*([NSEWOnsewo])\s*$", t)
        if m:
            hem = hem or m.group(1).upper()
            t = t[:m.start()]
    t = re.sub(r"\b(grados?|degrees?|deg|minutos?|mins?|minutes?|segundos?|segs?|secs?|seconds?|"
               r"latitud|latitude|lat|longitud|longitude|long|lon|lng|coord\w*)\b", " ", t, flags=re.I)
    if re.search(r"[A-Za-z]", t):
        return None, None
    t = re.sub(r"[°'\"]", " ", t).replace(":", " ")
    t = re.sub(r"(?<=\d)\s*-\s*(?=\d)", " ", t)
    if t.count(",") >= 2 and "." not in t:
        t = t.replace(",", " ")
    else:
        t = t.replace(",", ".")
    patron = r"[-+]?\d+(?:\.\d+)?"
    partes = re.findall(patron, t)
    if re.sub(patron, "", t).strip(" \t.;"):
        return None, None
    return _armar(partes, hem)


def _angulo_numerico(x):
    if abs(x) <= 180:
        return x
    if abs(x) < 1000:
        return None
    r = _compacto(x)
    return None if r is None else math.copysign(r, x)


def parse_angulo(v, permitir_compacto=True):
    """Cualquier celda -> (grados_decimales, hemisferio|None)."""
    if es_vacio(v) or isinstance(v, bool):
        return None, None
    if isinstance(v, (int, float)):
        x = float(v)
        if abs(x) <= 180:
            return x, None
        return (_angulo_numerico(x), None) if permitir_compacto else (None, None)
    if isinstance(v, str):
        n = parse_numero(v)
        if n is not None:
            if abs(n) <= 180:
                return n, None
            return (_angulo_numerico(n), None) if permitir_compacto else (None, None)
        return parse_angulo_texto(v)
    return None, None


# ============================================================== par en una celda
def _dividir_par(t):
    partes = [p for p in re.split(r"\s*[;|\t\n]\s*|\s+/\s+|\s+y\s+", t) if p.strip()]
    if len(partes) == 2:
        return partes
    partes = [p for p in re.split(r",\s+", t) if p.strip()]
    if len(partes) == 2:
        return partes
    if t.count(",") == 1 and re.fullmatch(r"\s*\S+,\S+\s*", t):
        return t.split(",")
    if re.match(r"^\s*[NSEWOnsewo]\s*[-+]?\d", t):
        p = [x for x in re.split(r"\s+(?=[NSEWOnsewo]\s*[-+]?\d)", t) if x.strip()]
        if len(p) == 2:
            return p
    p = [x for x in re.split(r"(?<=[NSEWO])\s+(?=[-+]?\d)|(?<=\d[sS])\s+(?=[-+]?\d)", t) if x.strip()]
    if len(p) == 2:
        return p
    toks = re.findall(r"[-+]?\d+(?:[.,]\d+)?", t)
    if len(toks) == 2:
        return toks
    if len(toks) == 4:
        return [" ".join(toks[:2]), " ".join(toks[2:])]
    if len(toks) == 6:
        return [" ".join(toks[:3]), " ".join(toks[3:])]
    return None


def parse_par(v):
    """Celda con un par de coordenadas -> dict(tipo='geo'|'proj', ...) o None."""
    if es_vacio(v) or not isinstance(v, str):
        return None
    t = v.strip()
    # WKT
    m = re.match(r"^\s*POINT\s*Z?\s*\(\s*([-+\d.eE]+)\s+([-+\d.eE]+)(?:\s+[-+\d.eE]+)?\s*\)\s*$", t, re.I)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        if abs(a) <= 180 and abs(b) <= 90:
            return dict(tipo="geo", lat=b, lon=a, hem=False)
        return dict(tipo="proj", a=a, b=b, orden="EN", zona=None)
    t = re.sub(r"^\s*(geo:|coordenadas?:?|coords?:?|ubicaci[oó]n:?|posici[oó]n:?|lat\s*/?\s*lon:?)\s*",
               "", t, flags=re.I)
    t = t.strip("()[]{} ")
    if not t:
        return None

    # ---- ¿proyectadas?
    crudos = re.findall(r"[-+]?\d[\d.,]*", t)
    nums = [parse_numero(x.rstrip(".,")) for x in crudos]
    nums = [x for x in nums if x is not None]
    grandes = [x for x in nums if abs(x) >= 100000]
    if len(grandes) == 2 and max(abs(x) for x in grandes) > 1_800_000 and len(nums) <= 4:
        zona = None
        m = re.match(r"^\s*(\d{1,2})\s*([C-HJ-NP-Xc-hj-np-x])\s+\d", t)
        if m:
            zona = int(m.group(1))
        else:
            m = re.search(r"\b(?:faja|zona|zone|huso|utm|f)\s*[:=]?\s*(\d{1,2})\b", t, re.I)
            if m:
                zona = int(m.group(1))
        etiquetas = re.findall(r"(este|east|easting|norte|north|northing|\b[xyen]\b)\s*[:=]?\s*[-+]?\d", t, re.I)
        etiquetas = [e.lower() for e in etiquetas]
        orden = None
        if len(etiquetas) == 2:
            e0 = etiquetas[0]
            if e0[0] == "e" and etiquetas[1][0] == "n":
                orden = "EN"
            elif e0[0] == "n" and etiquetas[1][0] == "e":
                orden = "NE"
        return dict(tipo="proj", a=grandes[0], b=grandes[1], orden=orden, zona=zona)

    # ---- geográficas
    orden_lab = [x.lower()[:3] for x in re.findall(r"\b(lat\w*|lon\w*|lng)\b", t, re.I)]
    partes = _dividir_par(t)
    if not partes or len(partes) != 2:
        return None
    (va, ha), (vb, hb) = parse_angulo(partes[0], False), parse_angulo(partes[1], False)
    if va is None or vb is None:
        # reintenta permitiendo formato compacto DDMMSS
        (va, ha), (vb, hb) = parse_angulo(partes[0]), parse_angulo(partes[1])
        if va is None or vb is None:
            return None
    if ha in ("N", "S") or hb in ("E", "W"):
        lat, lon = va, vb
    elif ha in ("E", "W") or hb in ("N", "S"):
        lat, lon = vb, va
    elif orden_lab[:2] == ["lon", "lat"] or orden_lab[:2] == ["lng", "lat"]:
        lat, lon = vb, va
    elif orden_lab[:2] in (["lat", "lon"], ["lat", "lng"]):
        lat, lon = va, vb
    else:
        lat, lon = va, vb
        if abs(va) > 90 or (57 <= abs(va) <= 76 and abs(vb) < 57):
            lat, lon = vb, va
    if abs(lat) > 90 or abs(lon) > 180:
        return None
    return dict(tipo="geo", lat=lat, lon=lon, hem=bool(ha or hb))


# ============================================================== encabezados
_SIST = [
    ("posgar94", r"\bposgar\s*(19)?\s*94\b|\bpg\s*94\b|\bpos\s*94\b"),
    ("posgar07", r"\bposgar\b|\bpg\s*07\b|\bpos\s*07\b|\bpg07\b|\bposgar07\b"),
    ("inchauspe", r"\binchauspe\b|\bci\s*69\b"),
    ("wgs84", r"\bwgs\b|\bgps\b|\bw\s*g\s*s\b"),
]
_PART = r"(g|gr|grad|grados?|deg|d|m|min|mins|minutos?|s|seg|segs|sec|segundos?|hem|h|dd|dec|gd|gms|dms|wgs|utm|gk)?"
_RE_LAT = re.compile(r"(lat|lati|latit|latitud|latitude|latitudes)" + _PART)
_RE_LON = re.compile(r"(lon|long|lng|longit|longitud|longitude|longitudes)" + _PART)
_ESTE = {"este", "east", "easting", "eastings", "e"}
_NORTE = {"norte", "north", "northing", "northings", "n"}
_X = {"x", "xcoord", "coordx", "xutm", "xgk", "xposgar", "xwgs", "coordenadax", "abscisa", "abscisas"}
_Y = {"y", "ycoord", "coordy", "yutm", "ygk", "yposgar", "ywgs", "coordenaday", "ordenada", "ordenadas"}
_ALT = {"alt", "altura", "altitud", "altitude", "elevacion", "elevation", "elev", "cota", "z", "msnm", "height"}
_ZONA = {"faja", "fajas", "zona", "zone", "huso", "husos", "fuso"}
_PAR = {"coordenadas", "coordenada", "coord", "coords", "coordinates", "coordinate", "ubicacion",
        "posicion", "location", "geom", "geometry", "wkt", "latlon", "latlong", "latlng", "lonlat",
        "xy", "geo", "geoposicion", "georreferencia", "gps"}
_P_DEG = {"g", "gr", "grad", "grado", "grados", "deg", "degree", "degrees", "d"}
_P_MIN = {"m", "min", "mins", "minuto", "minutos", "minute", "minutes"}
_P_SEC = {"s", "seg", "segs", "sec", "secs", "segundo", "segundos", "second", "seconds"}
_P_HEM = {"hem", "hemi", "hemisferio", "h", "sentido", "ns", "ew", "signo", "sur", "norte", "oeste", "dir",
          "direccion"}
_DMS_TOK = {"gms", "dms", "sexagesimal", "sexagesimales", "sexa", "ggmmss", "ddmmss"}


def clasificar_encabezado(h):
    info = dict(axis=None, part="full", sist=None, proy=None, dms=False, zona_kind=None, txt="")
    if h is None:
        return info
    txt = norm(h)
    info["txt"] = txt
    toks = txt.split()
    if not toks or len(toks) > 6:
        return info
    for nombre, pat in _SIST:
        if re.search(pat, txt):
            info["sist"] = nombre
            break
    if re.search(r"\butm\b|\bhuso\b", txt):
        info["proy"] = "utm"
    elif re.search(r"\bgauss\b|\bkruger\b|\bgk\b|\bfaja\b", txt):
        info["proy"] = "gk"
    if any(t in _DMS_TOK for t in toks) or re.search(r"grados minutos segundos|g m s|dd mm ss", txt):
        info["dms"] = True

    lat_tok = lon_tok = None
    for t in toks:
        m = _RE_LAT.fullmatch(t)
        if m and lat_tok is None:
            lat_tok = m
        m = _RE_LON.fullmatch(t)
        if m and lon_tok is None:
            lon_tok = m
    if lat_tok and lon_tok:
        info["axis"] = "par_geo"
        return info
    if lat_tok or lon_tok:
        m = lat_tok or lon_tok
        info["axis"] = "lat" if lat_tok else "lon"
        suf = m.group(2) or ""
        rest = [t for t in toks if t != m.group(0)]
        cand = ([suf] if suf else []) + rest
        for c in cand:
            if c in _P_DEG:
                info["part"] = "deg"
            elif c in _P_MIN:
                info["part"] = "min"
            elif c in _P_SEC:
                info["part"] = "sec"
            elif c in _P_HEM:
                info["part"] = "hem"
            elif c in ("gms", "dms"):
                info["dms"] = True
            elif c in ("dd", "dec", "gd"):
                info["part"] = "full"
            if info["part"] != "full":
                break
        if info["part"] == "full" and any(c in ("gms", "dms") for c in cand):
            info["dms"] = True
        return info

    ejes = []
    for t in toks:
        if t in _ESTE:
            ejes.append("east")
        elif t in _NORTE:
            ejes.append("north")
        elif t in _X:
            ejes.append("x")
        elif t in _Y:
            ejes.append("y")
    if len(ejes) == 2 and set(ejes) in ({"east", "north"}, {"x", "y"}):
        info["axis"] = "par_proj"
        return info
    if ejes:
        info["axis"] = ejes[0]
        return info
    if any(t in _ZONA for t in toks) and len(toks) <= 3:
        info["axis"] = "zona"
        info["zona_kind"] = "faja" if "faja" in toks or "fajas" in toks else (
            "utm" if "huso" in toks or "husos" in toks else "zona")
        return info
    if any(t in _ALT for t in toks) and len(toks) <= 3:
        info["axis"] = "alt"
        return info
    if len(toks) <= 3 and any(t in _PAR for t in toks):
        info["axis"] = "par"
        return info
    return info


EJES_COORD = {"lat", "lon", "east", "north", "x", "y", "par", "par_geo", "par_proj", "zona"}


# ============================================================== contexto
def elegir_opcion(titulo, texto, opciones, defecto=0):
    try:
        import tkinter as tk
        from tkinter import ttk
        res = {}
        root = tk.Tk()
        root.title(titulo)
        root.resizable(False, False)
        ttk.Label(root, text=texto, wraplength=380, justify="left").pack(padx=15, pady=(15, 5))
        combo = ttk.Combobox(root, values=opciones, state="readonly", width=45)
        combo.current(defecto)
        combo.pack(padx=15, pady=5)

        def ok():
            res["i"] = combo.current()
            root.destroy()

        ttk.Button(root, text="Aceptar", command=ok).pack(pady=(5, 15))
        root.protocol("WM_DELETE_WINDOW", ok)
        root.mainloop()
        return res.get("i", defecto)
    except Exception:
        print(texto)
        for n, o in enumerate(opciones, 1):
            print(f"  {n}. {o}")
        s = input(f"Número [{defecto + 1}]: ").strip()
        return int(s) - 1 if s.isdigit() and 1 <= int(s) <= len(opciones) else defecto


class Contexto:
    def __init__(self, interactivo=True):
        self.interactivo = interactivo
        self.faja = None
        self.zona_utm = None
        self.hemi_utm = "S"
        self.sist_archivo = None
        self._avisado = set()

    def aviso(self, msg):
        if msg not in self._avisado:
            self._avisado.add(msg)
            print("   [!] " + msg)

    def pedir_utm(self):
        if self.zona_utm is None:
            if self.interactivo:
                ops = ["18 (78°O–72°O)", "19 (72°O–66°O)", "20 (66°O–60°O)", "21 (60°O–54°O)"]
                i = elegir_opcion("Huso UTM", "Hay coordenadas UTM sin huso indicado.\n"
                                  "¿A qué huso (WGS84, hemisferio sur) corresponden?", ops, 1)
                self.zona_utm = 18 + i
            else:
                self.zona_utm = 19
                self.aviso("UTM sin huso: se asume huso 19S.")
        return self.zona_utm

    def pedir_faja(self):
        if self.faja is None:
            if self.interactivo:
                ops = [f"Faja {n} (meridiano central {-72 + 3 * (n - 1)}°)" for n in range(1, 8)]
                i = elegir_opcion("Faja Gauss-Krüger", "Hay coordenadas Gauss-Krüger sin faja identificable.\n"
                                  "¿A qué faja corresponden?", ops, 1)
                self.faja = i + 1
            else:
                self.faja = 2
                self.aviso("Coordenadas planas sin faja: se asume faja 2.")
        return self.faja


def contexto_desde_texto(ctx, texto):
    t = norm(texto)
    m = re.search(r"\bfaja\s*(\d)\b", t) or (re.search(r"\bf\s*(\d)\b", t) if re.search(r"posgar|gk|gauss", t) else None)
    if m and 1 <= int(m.group(1)) <= 7 and ctx.faja is None:
        ctx.faja = int(m.group(1))
    m = re.search(r"\b(?:utm|huso|zona|zone)\s*(\d{1,2})\b", t) if re.search(r"utm|huso", t) else None
    if m and 1 <= int(m.group(1)) <= 60 and ctx.zona_utm is None:
        ctx.zona_utm = int(m.group(1))
    if ctx.sist_archivo is None:
        for nombre, pat in _SIST:
            if re.search(pat, t):
                ctx.sist_archivo = nombre
                break


# ============================================================== CRS
_TRANSF = {}


def transformador(a, b):
    k = (a, b)
    if k not in _TRANSF:
        _TRANSF[k] = Transformer.from_crs(CRS.from_epsg(a), CRS.from_epsg(b), always_xy=True)
    return _TRANSF[k]


def epsg_geo(sist):
    return SISTEMAS.get(sist or "wgs84", SISTEMAS["wgs84"])["geo"]


def epsg_proj(e, zona, zona_kind, sist, proy, ctx, interactivo=True):
    """Deduce el EPSG proyectado a partir del Este, la zona y las pistas disponibles."""
    if proy == "utm" or sist == "wgs84":
        z = zona if (zona and 1 <= zona <= 60) else ctx.zona_utm
        if z is None:
            if not interactivo:
                return None
            z = ctx.pedir_utm()
        return (32700 if ctx.hemi_utm == "S" else 32600) + int(z)
    tabla = SISTEMAS[sist or "posgar07"]["proj"]
    if e >= 1_000_000:
        fj = int(e // 1_000_000)
        return tabla[fj - 1] if 1 <= fj <= 7 else None
    if zona is not None and 1 <= zona <= 7 and (zona_kind in ("faja", "zona") or sist or proy == "gk"):
        return tabla[int(zona) - 1]
    if ctx.faja is not None:
        return tabla[ctx.faja - 1]
    if zona is not None and zona > 7:
        return (32700 if ctx.hemi_utm == "S" else 32600) + int(zona)
    if sist or proy == "gk":
        if not interactivo:
            return None
        return tabla[ctx.pedir_faja() - 1]
    if not interactivo:
        return None
    return (32700 if ctx.hemi_utm == "S" else 32600) + ctx.pedir_utm()


def en_caja(lon, lat):
    return CAJA_ARG[0] <= lon <= CAJA_ARG[1] and CAJA_ARG[2] <= lat <= CAJA_ARG[3]


def faja_de_lon(lon):
    return min(max(round((lon + 72) / 3) + 1, 1), 7)


# ============================================================== lectura de tablas
def leer_hojas(ruta: Path):
    """Devuelve {nombre_hoja: lista de filas (listas de celdas)}."""
    ext = ruta.suffix.lower()
    if ext == ".csv":
        df = None
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(ruta, header=None, dtype=str, sep=None, engine="python", encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        if df is None:
            raise ValueError("No se pudo leer el CSV")
        return {"csv": df.astype(object).where(df.notna(), None).values.tolist()}
    hojas = pd.read_excel(ruta, sheet_name=None, header=None, dtype=object)
    return {n: d.astype(object).where(d.notna(), None).values.tolist() for n, d in hojas.items()}


def detectar_fila_encabezado(filas):
    mejor, idx = 0, None
    for r, fila in enumerate(filas[:25]):
        s = 0
        for v in fila:
            if isinstance(v, str) and 0 < len(v.strip()) <= 40:
                if clasificar_encabezado(v)["axis"] in EJES_COORD:
                    s += 1
        if s > mejor:
            mejor, idx = s, r
    return idx


def armar_tabla(filas):
    filas = [list(f) for f in filas]
    ncol = max((len(f) for f in filas), default=0)
    filas = [f + [None] * (ncol - len(f)) for f in filas]
    h = detectar_fila_encabezado(filas)
    if h is None:
        headers = [""] * ncol
        inicio = 0
    else:
        headers = [("" if es_vacio(v) else str(v).strip()) for v in filas[h]]
        inicio = h + 1
    datos, nums = [], []
    for k in range(inicio, len(filas)):
        if any(not es_vacio(v) for v in filas[k]):
            datos.append(filas[k])
            nums.append(k + 1)
    return dict(headers=headers, infos=[clasificar_encabezado(x) for x in headers],
                filas=datos, nums=nums, ncol=ncol, con_encabezado=h is not None)


# ============================================================== validaciones por valores
def _muestra(vals, n=300):
    return [v for v in vals if not es_vacio(v)][:n]


def valida_geo(vals, tipo):
    m = _muestra(vals)
    if not m:
        return False
    lim = 90 if tipo == "lat" else 180
    ok = sum(1 for v in m if (lambda a: a is not None and abs(a) <= lim)(parse_angulo(v)[0]))
    return ok / len(m) >= 0.7


def valida_proj(vals):
    m = _muestra(vals)
    if not m:
        return False
    ok = sum(1 for v in m if (lambda a: a is not None and 1e5 <= abs(a) <= 1.2e7)(parse_numero(v)))
    return ok / len(m) >= 0.7


def valida_zona(vals):
    m = _muestra(vals)
    if not m:
        return False
    ok = 0
    for v in m:
        z = _leer_zona(v)
        if z is not None:
            ok += 1
    return ok / len(m) >= 0.7


def _leer_zona(v):
    if es_vacio(v):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v) if float(v).is_integer() and 1 <= int(v) <= 60 else None
    mm = re.search(r"(\d{1,2})", str(v))
    return int(mm.group(1)) if mm and 1 <= int(mm.group(1)) <= 60 else None


def valida_par(vals):
    m = _muestra(vals, 60)
    if not m:
        return False
    ok = sum(1 for v in m if parse_par(v) is not None)
    return ok / len(m) >= 0.7


# ============================================================== especificaciones
def construir_especificaciones(tabla, ctx):
    H = tabla["infos"]
    n = tabla["ncol"]
    filas = tabla["filas"]

    def col(i):
        return [f[i] for f in filas]

    specs = []
    usados_cols = set()

    # ---------- componentes G/M/S por columnas
    comps = {}
    for i in range(n):
        if H[i]["axis"] in ("lat", "lon") and H[i]["part"] in ("deg", "min", "sec", "hem"):
            comps.setdefault(H[i]["axis"], {}).setdefault(H[i]["part"], i)
    lat_cols = [i for i in range(n) if H[i]["axis"] == "lat" and H[i]["part"] == "full"]
    lon_cols = [i for i in range(n) if H[i]["axis"] == "lon" and H[i]["part"] == "full"]
    for eje, lista in (("lat", lat_cols), ("lon", lon_cols)):
        c = comps.get(eje)
        if c and "deg" in c and "min" not in c:
            lista.append(c["deg"])          # "Lat_G" sin minutos = decimal
            del c["deg"]
    if "lat" in comps and "lon" in comps and all("deg" in comps[e] and "min" in comps[e] for e in comps):
        cl, co = comps["lat"], comps["lon"]
        ok = True
        for c_ in (cl, co):
            d = [parse_numero(v) for v in _muestra(col(c_["deg"]))]
            mi = [parse_numero(v) for v in _muestra(col(c_["min"]))]
            if not d or any(x is None or abs(x) > 180 for x in d) or any(x is None or not 0 <= x < 60 for x in mi):
                ok = False
        if ok:
            specs.append(dict(tipo="comp", lat=cl, lon=co, prio=3, sist=H[cl["deg"]]["sist"] or H[co["deg"]]["sist"]))
            usados_cols.update(cl.values())
            usados_cols.update(co.values())

    # ---------- lat/lon por columnas (decimal o DMS texto)
    lats = [i for i in lat_cols if valida_geo(col(i), "lat")]
    lons = [i for i in lon_cols if valida_geo(col(i), "lon")]
    libres = list(lons)
    for a in lats:
        if not libres:
            break
        libres.sort(key=lambda b: (H[a]["dms"] != H[b]["dms"], abs(a - b)))
        b = libres.pop(0)
        es_texto = H[a]["dms"] or any(isinstance(v, str) and parse_numero(v) is None for v in _muestra(col(a), 20))
        specs.append(dict(tipo="geo", lat=a, lon=b, prio=4 if es_texto else 0,
                          sist=H[a]["sist"] or H[b]["sist"]))

    # ---------- proyectadas / X-Y
    E = [i for i in range(n) if H[i]["axis"] == "east"]
    N = [i for i in range(n) if H[i]["axis"] == "north"]
    X = [i for i in range(n) if H[i]["axis"] == "x"]
    Y = [i for i in range(n) if H[i]["axis"] == "y"]
    zonas = [i for i in range(n) if H[i]["axis"] == "zona" and valida_zona(col(i))]

    def emparejar(A, B, etiquetado):
        A = [i for i in A if valida_proj(col(i)) or True]
        B = list(B)
        for a in A:
            if not B:
                break
            B.sort(key=lambda b: abs(a - b))
            b = B.pop(0)
            yield a, b, etiquetado

    pares_proj, pares_geo_xy = [], []
    usados_xy = set()
    for a, b, et in list(emparejar(E, N, True)) + list(emparejar(X, Y, False)) + \
            list(emparejar(E, Y, True)) + list(emparejar(X, N, True)):
        if a in usados_xy or b in usados_xy:
            continue
        va, vb = col(a), col(b)
        if valida_proj(va) and valida_proj(vb):
            usados_xy.update((a, b))
            zc = min(zonas, key=lambda z: abs(z - a)) if zonas else None
            specs.append(dict(
                tipo="proj", e=a, n=b, etiquetado=et, prio=1, zona=zc,
                zona_kind=H[zc]["zona_kind"] if zc is not None else None,
                sist=H[a]["sist"] or H[b]["sist"], proy=H[a]["proy"] or H[b]["proy"]))
        elif not et and valida_geo(va, "lon") and valida_geo(vb, "lat"):
            usados_xy.update((a, b))
            specs.append(dict(tipo="geo", lat=b, lon=a, prio=0, sist=H[a]["sist"] or H[b]["sist"]))

    # ---------- par en una sola columna
    for i in range(n):
        if H[i]["axis"] in ("par", "par_geo", "par_proj") and valida_par(col(i)):
            specs.append(dict(tipo="par", col=i, prio=2, sist=H[i]["sist"], proy=H[i]["proy"]))

    # ---------- deducción por valores si no hubo nada
    if not specs:
        specs = detectar_por_valores(tabla, ctx)

    specs.sort(key=lambda s: s["prio"])
    return specs


def detectar_por_valores(tabla, ctx):
    """Sin encabezados reconocibles: se deduce por los rangos y formatos de los valores."""
    n = tabla["ncol"]
    filas = tabla["filas"]
    specs = []

    def col(i):
        return [f[i] for f in filas]

    dec = {}
    for i in range(n):
        m = _muestra(col(i))
        if len(m) < 1:
            continue
        nums = [parse_numero(v) for v in m]
        if sum(x is not None for x in nums) / len(m) >= 0.8:
            dec[i] = [x for x in nums if x is not None]

    def frac_no_entera(v):
        return sum(1 for x in v if abs(x - round(x)) > 1e-9) / len(v)

    def todos(v, lo, hi, umbral=0.9):
        return sum(1 for x in v if lo <= x <= hi) / len(v) >= umbral

    # 1) decimal Argentina (con o sin signo)
    lat_a = [i for i, v in dec.items() if frac_no_entera(v) > 0.5 and
             (todos(v, -57, -20) or todos(v, 20, 57))]
    lon_a = [i for i, v in dec.items() if frac_no_entera(v) > 0.5 and
             (todos(v, -76, -52) or todos(v, 52, 76))]
    par = None
    for a in lat_a:
        for b in lon_a:
            if a != b:
                par = (a, b)
                break
        if par:
            break
    if par is None:
        # 2) decimal genérico: dos columnas contiguas
        for i in sorted(dec):
            j = i + 1
            if j in dec and frac_no_entera(dec[i]) > 0.5 and frac_no_entera(dec[j]) > 0.5:
                if todos(dec[i], -90, 90) and todos(dec[j], -180, 180):
                    par = (i, j)
                    break
                if todos(dec[j], -90, 90) and todos(dec[i], -180, 180):
                    par = (j, i)
                    break
    if par:
        specs.append(dict(tipo="geo", lat=par[0], lon=par[1], prio=0, sist=ctx.sist_archivo))

    # 3) proyectadas
    E_like = [i for i, v in dec.items() if todos(v, 1e5, 8.5e6)]
    N_like = [i for i, v in dec.items() if todos(v, 1e6, 1.1e7)]
    for a in E_like:
        b = next((j for j in N_like if j > a), None)
        if b is not None:
            specs.append(dict(tipo="proj", e=a, n=b, etiquetado=False, prio=1, zona=None, zona_kind=None,
                              sist=ctx.sist_archivo, proy=None))
            break

    # 4) par en una sola columna de texto / DMS en dos columnas de texto
    for i in range(n):
        if i in dec:
            continue
        if valida_par(col(i)):
            specs.append(dict(tipo="par", col=i, prio=2, sist=ctx.sist_archivo, proy=None))
    textos = [i for i in range(n) if i not in dec and _muestra(col(i))]
    for i in textos:
        j = i + 1
        if j in textos:
            ai = [parse_angulo(v, False)[0] for v in _muestra(col(i))]
            aj = [parse_angulo(v, False)[0] for v in _muestra(col(j))]
            if ai and aj and all(x is not None for x in ai) and all(x is not None for x in aj):
                lat, lon = (i, j)
                if statistics_med([abs(x) for x in ai]) > 57 > statistics_med([abs(x) for x in aj]):
                    lat, lon = j, i
                specs.append(dict(tipo="geo", lat=lat, lon=lon, prio=4, sist=ctx.sist_archivo))
                break

    # 5) tres columnas G | M | S consecutivas, dos veces
    if not specs:
        trip = []
        for i in range(n - 2):
            if i in dec and i + 1 in dec and i + 2 in dec:
                d, m, s = dec[i], dec[i + 1], dec[i + 2]
                if all(abs(x) <= 180 and x == int(x) for x in d) and \
                        all(0 <= x < 60 and x == int(x) for x in m) and all(0 <= x < 60 for x in s):
                    trip.append(i)
        if len(trip) >= 2 and trip[1] >= trip[0] + 3:
            a, b = trip[0], trip[1]
            cl = dict(deg=a, min=a + 1, sec=a + 2)
            co = dict(deg=b, min=b + 1, sec=b + 2)
            if statistics_med([abs(x) for x in dec[a]]) > 57 > statistics_med([abs(x) for x in dec[b]]):
                cl, co = co, cl
            specs.append(dict(tipo="comp", lat=cl, lon=co, prio=3, sist=ctx.sist_archivo))
    return specs


def statistics_med(v):
    v = sorted(v)
    return v[len(v) // 2] if v else 0


# ============================================================== evaluación
def _signo_texto(v):
    return isinstance(v, str) and v.strip().startswith("-")


def _hem_cel(v):
    if es_vacio(v):
        return None
    t = sin_acentos(str(v)).strip().upper()
    if not t:
        return None
    if t.startswith("SUR") or t == "S":
        return "S"
    if t.startswith("NORTE") or t == "N":
        return "N"
    if t.startswith("OESTE") or t in ("W", "O", "WEST"):
        return "W"
    if t.startswith("ESTE") or t == "E" or t.startswith("EAST"):
        return "E"
    return t[0] if t[0] in "NSEWO" else None


def _leer_comp(f, c):
    d = parse_numero(f[c["deg"]])
    if d is None:
        return None, None
    mi = parse_numero(f[c["min"]]) if "min" in c and not es_vacio(f[c["min"]]) else 0.0
    se = parse_numero(f[c["sec"]]) if "sec" in c and not es_vacio(f[c["sec"]]) else 0.0
    if mi is None or se is None or not (0 <= mi < 60) or not (0 <= se < 60):
        return None, None
    if (mi or se) and d != int(d):
        return None, None
    hem = _hem_cel(f[c["hem"]]) if "hem" in c else None
    neg = d < 0 or math.copysign(1, d) < 0 or _signo_texto(f[c["deg"]]) or hem in ("S", "W")
    val = abs(d) + mi / 60 + se / 3600
    return (-val if neg else val), hem


def postproceso_geo(items, ctx):
    """items: [(k, lat, lon, hem_explicita)] -> corrige intercambio y signo omitido (Argentina)."""
    if not items:
        return items
    lats = [it[1] for it in items]
    lons = [it[2] for it in items]
    if statistics_med([abs(x) for x in lats]) > 57 and statistics_med([abs(x) for x in lons]) < 57 and \
            all(abs(x) <= 90 for x in lons):
        ctx.aviso("Latitud y longitud parecían estar invertidas: se intercambiaron.")
        items = [(k, lo, la, h) for k, la, lo, h in items]
        lats, lons = lons, lats
    if not any(h for *_, h in items) and all(x >= 0 for x in lats) and all(x >= 0 for x in lons) and \
            all(20 <= x <= 57 for x in lats) and all(52 <= x <= 76 for x in lons):
        ctx.aviso("Coordenadas sin signo: se asumió hemisferio Sur y longitud Oeste (valores negativos).")
        items = [(k, -la, -lo, h) for k, la, lo, h in items]
    return items


def procesar_proj(items, sist, proy, zona_kind, etiquetado, ctx):
    """items: [(k, a, b, zona, fijo_EN)] -> {k: (epsg, x_este, y_norte)}."""
    if not items:
        return {}

    def puntaje(muestra, invertir):
        bien = evaluados = 0
        for k, a, b, z, fijo in muestra:
            e, nn = (b, a) if invertir else (a, b)
            ep = epsg_proj(e, z, zona_kind, sist, proy, ctx, interactivo=False)
            if ep is None:
                continue
            evaluados += 1
            try:
                lon, lat = transformador(ep, 4326).transform(e, nn)
            except Exception:
                continue
            if math.isfinite(lon) and math.isfinite(lat) and en_caja(lon, lat):
                bien += 1
        return bien / len(muestra) if muestra else 0

    libres = [it for it in items if not it[4]]
    invertir = False
    if libres:
        muestra = libres[:40]
        s1, s2 = puntaje(muestra, False), puntaje(muestra, True)
        if etiquetado:
            if s1 < 0.5 and s2 >= 0.8:
                invertir = True
                ctx.aviso("Las columnas Este/Norte parecen estar intercambiadas: se corrigió.")
        else:
            if s2 > s1:
                invertir = True
            elif s2 == s1:
                ma = statistics_med([abs(x[1]) for x in libres])
                mb = statistics_med([abs(x[2]) for x in libres])
                invertir = ma > mb          # el Este es el valor menor
            if invertir:
                ctx.aviso("Columnas X/Y: se interpretó Y=Este y X=Norte (convención argentina).")
    res = {}
    for k, a, b, z, fijo in items:
        e, nn = (b, a) if (invertir and not fijo) else (a, b)
        ep = epsg_proj(e, z, zona_kind, sist, proy, ctx, interactivo=True)
        if ep is None:
            continue
        res[k] = (ep, e, nn)
    return res


def evaluar(spec, tabla, ctx, pendientes):
    """Devuelve {k: (epsg, x, y)} solo para las filas de 'pendientes'."""
    filas = tabla["filas"]
    sist = spec.get("sist") or ctx.sist_archivo
    geo_items, proj_items, res = [], [], {}
    zona_kind = spec.get("zona_kind")

    if spec["tipo"] == "geo":
        for k in pendientes:
            a, b = filas[k][spec["lat"]], filas[k][spec["lon"]]
            if es_vacio(a) or es_vacio(b):
                continue
            (va, ha), (vb, hb) = parse_angulo(a), parse_angulo(b)
            if va is None or vb is None:
                continue
            if ha in ("E", "W") or hb in ("N", "S"):
                va, vb, ha, hb = vb, va, hb, ha
            if abs(va) > 90 or abs(vb) > 180:
                continue
            geo_items.append((k, va, vb, bool(ha or hb)))
    elif spec["tipo"] == "comp":
        for k in pendientes:
            la, ha = _leer_comp(filas[k], spec["lat"])
            lo, hb = _leer_comp(filas[k], spec["lon"])
            if la is None or lo is None or abs(la) > 90 or abs(lo) > 180:
                continue
            geo_items.append((k, la, lo, bool(ha or hb)))
    elif spec["tipo"] == "proj":
        for k in pendientes:
            a, b = parse_numero(filas[k][spec["e"]]), parse_numero(filas[k][spec["n"]])
            if a is None or b is None:
                continue
            z = _leer_zona(filas[k][spec["zona"]]) if spec.get("zona") is not None else None
            proj_items.append((k, a, b, z, False))
    elif spec["tipo"] == "par":
        proj_par = []
        for k in pendientes:
            p = parse_par(filas[k][spec["col"]])
            if not p:
                continue
            if p["tipo"] == "geo":
                geo_items.append((k, p["lat"], p["lon"], p["hem"]))
            else:
                if p["orden"] == "NE":
                    proj_par.append((k, p["b"], p["a"], p["zona"], True))
                else:
                    proj_par.append((k, p["a"], p["b"], p["zona"], p["orden"] == "EN"))
        proj_items = proj_par

    if geo_items:
        for k, la, lo, _ in postproceso_geo(geo_items, ctx):
            res[k] = (epsg_geo(sist), lo, la)
    if proj_items:
        et = spec.get("etiquetado", False)
        res.update(procesar_proj(proj_items, sist, spec.get("proy"), zona_kind, et, ctx))
    return res


def describir_spec(spec, tabla):
    def nombre(i):
        h = tabla["headers"][i]
        return f"'{h}'" if h else f"columna {i + 1}"
    t = spec["tipo"]
    if t == "geo":
        return f"Lat/Lon → {nombre(spec['lat'])} / {nombre(spec['lon'])}"
    if t == "comp":
        return "Lat/Lon en grados-minutos-segundos por columnas"
    if t == "proj":
        if spec.get("etiquetado"):
            return f"Este/Norte → {nombre(spec['e'])} / {nombre(spec['n'])}"
        return f"Coordenadas planas → {nombre(spec['e'])} / {nombre(spec['n'])}"
    return f"Par de coordenadas en {nombre(spec['col'])}"


def extraer_puntos(tabla, ctx):
    total = len(tabla["filas"])
    specs = construir_especificaciones(tabla, ctx)
    if not specs:
        return {}, []
    resultado, usados = {}, []
    for spec in specs:
        pend = [k for k in range(total) if k not in resultado]
        if not pend:
            break
        r = evaluar(spec, tabla, ctx, pend)
        if r:
            usados.append((spec, len(r)))
            resultado.update(r)
    return resultado, usados


# ============================================================== atributos DBF
def _nombre_dbf(base, usados):
    s = re.sub(r"[^A-Za-z0-9_]+", "_", sin_acentos(str(base))).strip("_") or "CAMPO"
    if s[0].isdigit():
        s = "C_" + s
    s = s[:10]
    orig, i = s, 1
    while s.upper() in usados:
        suf = f"_{i}"
        s = orig[:10 - len(suf)] + suf
        i += 1
    usados.add(s.upper())
    return s


def definir_campos(tabla):
    usados = {"LAT_WGS84", "LON_WGS84", "FILA_XLS"}
    campos = []
    for c in range(tabla["ncol"]):
        vals = [f[c] for f in tabla["filas"]]
        nv = [v for v in vals if not es_vacio(v)]
        if not nv:
            continue
        base = tabla["headers"][c] or f"COL{c + 1}"
        nombre = _nombre_dbf(base, usados)
        es_num = all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in nv)
        if es_num:
            f = [float(v) for v in nv]
            if all(x.is_integer() for x in f):
                mx = max(abs(x) for x in f)
                size = min(20, len(str(int(mx))) + (1 if any(x < 0 for x in f) else 0))
                campos.append((c, nombre, "N", max(size, 1), 0))
            else:
                dec = 0
                for x in f:
                    r = repr(x)
                    dec = max(dec, len(r.split(".")[1]) if "e" not in r and "." in r else 8)
                dec = min(dec, 10)
                mx = max(abs(x) for x in f)
                ent = len(str(int(mx)))
                neg = 1 if any(x < 0 for x in f) else 0
                if ent + neg + 1 + dec > 20:
                    dec = max(0, 20 - ent - neg - 1)
                campos.append((c, nombre, "N", min(20, ent + neg + 1 + dec), dec))
        else:
            ln = max(len(_a_texto(v).encode("utf-8")) for v in nv)     # el DBF mide bytes
            campos.append((c, nombre, "C", max(1, min(254, ln)), 0))
    return campos


def _a_texto(v):
    if es_vacio(v):
        return ""
    if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp)):
        return v.isoformat()[:19].replace("T", " ")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ============================================================== salida
def sufijo_salida(clave, faja):
    return {"wgs84": "WGS84", "posgar_geo": "POSGAR2007_geograficas",
            "posgar_faja": f"POSGAR2007_F{faja}"}[clave]


def limpiar_nombre(s):
    return re.sub(r"[^\w\-]+", "_", sin_acentos(s)).strip("_")


def convertir(puntos, destino):
    """puntos: [(epsg, x, y)] -> [(x', y')] en el CRS destino."""
    out = [None] * len(puntos)
    grupos = {}
    for i, (ep, x, y) in enumerate(puntos):
        grupos.setdefault(ep, []).append(i)
    for ep, idxs in grupos.items():
        xs = [puntos[i][1] for i in idxs]
        ys = [puntos[i][2] for i in idxs]
        if ep == destino:
            nx, ny = xs, ys
        else:
            nx, ny = transformador(ep, destino).transform(xs, ys)
        for j, i in enumerate(idxs):
            x, y = float(nx[j]), float(ny[j])
            out[i] = (x, y) if math.isfinite(x) and math.isfinite(y) else None
    return out


def escribir_shp(ruta: Path, coords, epsg, tabla, campos, ks, wgs):
    w = shapefile.Writer(str(ruta), shapeType=shapefile.POINT, encoding="utf-8")
    for _, nombre, tipo, size, dec in campos:
        w.field(nombre, tipo, size, dec)
    w.field("LAT_WGS84", "N", 14, 9)
    w.field("LON_WGS84", "N", 15, 9)
    w.field("FILA_XLS", "N", 8, 0)
    for (x, y), k, (lon, lat) in zip(coords, ks, wgs):
        w.point(x, y)
        reg = []
        for c, _, tipo, size, dec in campos:
            v = tabla["filas"][k][c]
            if tipo == "N":
                x_ = parse_numero(v) if not es_vacio(v) else None
                reg.append(None if x_ is None else (int(x_) if dec == 0 else x_))
            else:
                txt = _a_texto(v)
                while len(txt.encode("utf-8")) > size:
                    txt = txt[:-1]
                reg.append(txt)
        reg += [lat, lon, tabla["nums"][k]]
        w.record(*reg)
    w.close()
    ruta.with_suffix(".prj").write_text(CRS.from_epsg(epsg).to_wkt(version="WKT1_ESRI"), encoding="utf-8")
    ruta.with_suffix(".cpg").write_text("UTF-8", encoding="ascii")


def procesar_hoja(nombre_base, hoja, filas, salidas, faja_sel, carpeta, ctx_base, varias_hojas):
    ctx = Contexto(ctx_base.interactivo)
    ctx.faja, ctx.zona_utm, ctx.sist_archivo = ctx_base.faja, ctx_base.zona_utm, ctx_base.sist_archivo
    contexto_desde_texto(ctx, hoja)
    tabla = armar_tabla(filas)
    if not tabla["filas"]:
        print(f"   [{hoja}] hoja vacía.")
        return 0
    for h in tabla["headers"]:
        contexto_desde_texto(ctx, h)

    resultado, usados = extraer_puntos(tabla, ctx)
    if not resultado:
        print(f"   [{hoja}] no se reconocieron coordenadas.")
        return 0
    ks = sorted(resultado)
    puntos = [resultado[k] for k in ks]
    wgs = convertir(puntos, 4326)
    validos = [i for i, p in enumerate(wgs) if p is not None]
    if len(validos) < len(ks):
        print(f"   [!] {len(ks) - len(validos)} punto(s) descartado(s) por conversión inválida.")
    ks = [ks[i] for i in validos]
    puntos = [puntos[i] for i in validos]
    wgs = [wgs[i] for i in validos]

    total_filas = len(tabla["filas"])
    print(f"   [{hoja}] {len(ks)} de {total_filas} filas con coordenadas.")
    for spec, cnt in usados:
        print(f"       · {describir_spec(spec, tabla)}  ({cnt} filas)")
    sin = [tabla["nums"][k] for k in range(total_filas) if k not in resultado]
    if sin:
        muestra = ", ".join(map(str, sin[:10])) + (" …" if len(sin) > 10 else "")
        print(f"   [!] {len(sin)} fila(s) sin coordenadas válidas (filas Excel: {muestra}).")

    campos = definir_campos(tabla)
    prefijo = nombre_base + (f"_{limpiar_nombre(hoja)}" if varias_hojas else "")
    hechos = 0
    for clave in salidas:
        if clave == "wgs84":
            destino, faja = 4326, None
        elif clave == "posgar_geo":
            destino, faja = SISTEMAS["posgar07"]["geo"], None
        else:
            if faja_sel == "auto":
                cnt = Counter(faja_de_lon(lo) for lo, la in wgs)
                faja = cnt.most_common(1)[0][0]
                if len(cnt) > 1:
                    print(f"   [!] Los puntos abarcan varias fajas {dict(cnt)}: se usa la faja {faja} para todos.")
            else:
                faja = int(faja_sel)
            destino = SISTEMAS["posgar07"]["proj"][faja - 1]
        coords = convertir(puntos, destino)
        ok = [i for i, c in enumerate(coords) if c is not None]
        salida = carpeta / f"{prefijo}_{sufijo_salida(clave, faja)}.shp"
        try:
            escribir_shp(salida, [coords[i] for i in ok], destino, tabla, campos,
                         [ks[i] for i in ok], [wgs[i] for i in ok])
        except PermissionError:
            print(f"   [ERROR] No se pudo escribir {salida.name}: cerralo en QGIS/ArcGIS y reintentá.")
            continue
        print(f"   [OK] {salida.name}  (EPSG:{destino}, {len(ok)} puntos)")
        hechos += 1
    return hechos


# ============================================================== interfaz
EXTENSIONES = (".xlsx", ".xlsm", ".xls", ".csv")


def buscar_archivos(carpeta):
    return sorted([p for p in carpeta.iterdir()
                   if p.is_file() and p.suffix.lower() in EXTENSIONES and not p.name.startswith("~$")],
                  key=lambda p: p.name.lower())


def ventana_principal(archivos):
    """Devuelve (lista_archivos, [claves_salida], faja 'auto'|1..7)."""
    nombres = [a.name for a in archivos]
    opciones = nombres + (["» Todos los archivos «"] if len(archivos) > 1 else [])
    try:
        import tkinter as tk
        from tkinter import ttk
        res = {}
        root = tk.Tk()
        root.title("Excel → Shapefile")
        root.resizable(False, False)
        frm = ttk.Frame(root, padding=15)
        frm.pack()
        ttk.Label(frm, text="Archivo a transformar:").grid(row=0, column=0, sticky="w")
        combo = ttk.Combobox(frm, values=opciones, state="readonly", width=55)
        combo.current(0)
        combo.grid(row=1, column=0, columnspan=2, pady=(2, 10))
        ttk.Label(frm, text="Sistema(s) de referencia de salida:").grid(row=2, column=0, sticky="w")
        v1, v2, v3 = tk.BooleanVar(value=True), tk.BooleanVar(value=False), tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="WGS84 (EPSG:4326)", variable=v1).grid(row=3, column=0, sticky="w")
        ttk.Checkbutton(frm, text="POSGAR 2007 geográficas (EPSG:5340)", variable=v2).grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(frm, text="POSGAR 2007 plana Gauss-Krüger, faja:", variable=v3).grid(row=5, column=0, sticky="w")
        cf = ttk.Combobox(frm, values=["Automática"] + [str(i) for i in range(1, 8)], state="readonly", width=12)
        cf.current(0)
        cf.grid(row=5, column=1, sticky="w")

        def ok():
            sel = []
            if v1.get():
                sel.append("wgs84")
            if v2.get():
                sel.append("posgar_geo")
            if v3.get():
                sel.append("posgar_faja")
            if not sel:
                return
            res["archivos"] = archivos if combo.current() == len(archivos) else [archivos[combo.current()]]
            res["salidas"] = sel
            res["faja"] = "auto" if cf.current() == 0 else cf.current()
            root.destroy()

        ttk.Button(frm, text="Aceptar", command=ok).grid(row=6, column=0, columnspan=2, pady=(12, 0))
        root.mainloop()
        if "archivos" not in res:
            sys.exit("Cancelado.")
        return res["archivos"], res["salidas"], res["faja"]
    except Exception as e:
        if isinstance(e, SystemExit):
            raise
        print("Archivos encontrados:")
        for n, o in enumerate(opciones, 1):
            print(f"  {n}. {o}")
        s = input("Archivo [1]: ").strip()
        i = int(s) - 1 if s.isdigit() and 1 <= int(s) <= len(opciones) else 0
        sel_a = archivos if i == len(archivos) else [archivos[i]]
        print("Salida: 1=WGS84  2=POSGAR2007 geográficas  3=POSGAR2007 faja  (ej: 1,3)")
        s = input("Salidas [1,3]: ").strip() or "1,3"
        m = {"1": "wgs84", "2": "posgar_geo", "3": "posgar_faja"}
        sel = [m[x] for x in re.split(r"[ ,;]+", s) if x in m] or ["wgs84"]
        f = input("Faja (auto/1-7) [auto]: ").strip().lower() or "auto"
        return sel_a, sel, ("auto" if f == "auto" or not f.isdigit() else int(f))


def main():
    ap = argparse.ArgumentParser(description="Excel/CSV -> Shapefile")
    ap.add_argument("--archivo", help="archivo a convertir (por defecto: elegir en el menú)")
    ap.add_argument("--salida", default="wgs84,posgar_faja", help="wgs84,posgar_geo,posgar_faja")
    ap.add_argument("--faja", default="auto", help="auto o 1..7")
    ap.add_argument("--sin-menu", action="store_true", help="no abrir ventanas")
    args = ap.parse_args()

    carpeta = Path(__file__).resolve().parent
    if args.archivo:
        archivos = [Path(args.archivo).resolve()]
    else:
        archivos = buscar_archivos(carpeta)
        if not archivos:
            sys.exit(f"No hay archivos Excel ni CSV en {carpeta}")

    if args.sin_menu:
        seleccion = archivos
        salidas = [s for s in args.salida.split(",") if s in ("wgs84", "posgar_geo", "posgar_faja")]
        faja = "auto" if args.faja == "auto" else int(args.faja)
    else:
        seleccion, salidas, faja = ventana_principal(archivos)

    ctx = Contexto(interactivo=not args.sin_menu)
    print(f"Carpeta: {carpeta}")
    for a in seleccion:
        print(f"\n▶ {a.name}")
        try:
            hojas = leer_hojas(a)
        except Exception as e:
            print(f"   [ERROR] No se pudo leer: {e}")
            continue
        ctx_a = Contexto(ctx.interactivo)
        contexto_desde_texto(ctx_a, a.stem)
        for hoja, filas in hojas.items():
            try:
                procesar_hoja(limpiar_nombre(a.stem), hoja, filas, salidas, faja, a.parent, ctx_a, len(hojas) > 1)
            except Exception as e:
                print(f"   [ERROR] hoja '{hoja}': {e}")
    print("\nListo.")


if __name__ == "__main__":
    main()