# -*- coding: utf-8 -*-
"""
VALIDADOR DE BASES SSC (hechos de tránsito) - versión Python v2.0
Se empaqueta en un único .exe (ver build_exe.bat / README).

Uso:
  ValidadorSSC.exe                      -> abre selector de archivo
  ValidadorSSC.exe "C:\\ruta\\base.xlsx"  -> valida directamente (o arrastra el .xlsx sobre el .exe)

Salida: resultados/validacion_<base>_<fecha>.xlsx (Resumen + una hoja por error)
Catálogos: carpeta "catalogos" junto al .exe (cuadrantes.csv + 5 xlsx).
"""
import os
import re
import sys
import json
import pickle
import traceback
import unicodedata
import warnings
import datetime as dt

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ------------------------------------------------------------------------------
# 0. CONFIGURACIÓN
# ------------------------------------------------------------------------------
SEPARAR_ARCHIVOS = False   # True = además un .xlsx por cada validación
ABRIR_CARPETA = True       # abrir carpeta de resultados al terminar (Windows)

# PUNTO 1 es SIEMPRE la vialidad principal donde ocurrió el hecho: la clasificación de la
# vialidad se valida únicamente contra PUNTO 1. PUNTO 2 (calle de cruce/referencia) no se usa.
VIALIDAD_COLUMNAS = ["punto_1"]
VIALIDAD_MODO = "punto_1"

# Validación de la clasificación por GEORREFERENCIA (coordenadas del hecho vs. red vial SEMOVI/SSC).
# Requiere catalogos/red_vial.npz (se genera con convertir_red_vial.py). Si no existe, se valida solo por nombre.
USAR_GEORREFERENCIA = True
GEO_TOL_M = 30   # distancia máxima (m) entre la coordenada y la vialidad para considerarla "en ese punto"

# Reglas manuales por nombre (catalogos/reglas_vialidad.csv): si el PUNTO 1 contiene el patrón,
# esa clasificación manda sobre los catálogos y la georreferencia. Puedes agregar filas sin tocar el código.
USAR_REGLAS_NOMBRE = True

# Anexo 2 del Reglamento de Tránsito de la CDMX (catalogos/anexo2_reglamento_transito.csv): radiales (VAC) y
# avenidas primarias, por nombre.
USAR_ANEXO2 = True

# Clasificación por georreferencia (capas SEMOVI + KMZ de ejes). La clase de cada tramo sale de su CLASIF_15.
# 1) Si el nombre del PUNTO 1 es una vialidad mayor (Anexo 2, catálogos, reglas manuales), su clase manda.
# 2) Si el nombre es una calle secundaria o no se reconoce, la vialidad de mayor jerarquía a <= GEO_TOL_M de la
#    coordenada define la clase. Jerarquía: VAC > acceso carretero > eje vial > primaria > secundaria.
#    Ej.: un punto entre el Eje 4 Norte y Av. Gran Canal (secundaria) es EJE VIAL;
#    "CIRCUITO INTERIOR" es VAC aunque junto a la coordenada solo pase una calle.
GEO_TOL_CLASE = {0: 40, 2: 40}   # VAC y ejes son anchos (laterales): tolerancia mayor; el resto usa GEO_TOL_M

# Zonas viales SCT (catalogos/zonas_viales_sct.kmz): zona (Z1 a Z5) según el polígono donde cae la coordenada.
USAR_ZONAS_VIALES = True
COLUMNAS_ZONA_VIAL = ["zona_vial", "zona", "zona_vial_sct", "zonas_viales", "zona_sct"]

# Calles homónimas: los catálogos y el Anexo 2 solo traen nombres. Si el registro tiene coordenada válida, el
# nombre del PUNTO 1 cuenta como vialidad mayor (VAC, carretera, eje o primaria) únicamente si hay un tramo de ESE
# nombre a <= NOMBRE_CERCA_M de la coordenada (así "5 DE MAYO" del Centro no se confunde con la Av. 5 de Mayo
# de Álvaro Obregón). Sin coordenada válida se usa el nombre solo.
NOMBRE_CERCA_M = 60
NOMBRE_CERCA_VAC_M = 150   # las VAC son anchas: la coordenada puede quedar lejos de su eje

LON_MIN, LON_MAX = -99.40, -98.80
LAT_MIN, LAT_MAX = 19.00, 19.65

APP_DIR = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
           else os.path.dirname(os.path.abspath(__file__)))


def log(*a):
    print(*a, flush=True)


def ruta_catalogo(archivo):
    p = os.path.join(APP_DIR, "catalogos", archivo)
    return p if os.path.exists(p) else os.path.join(APP_DIR, archivo)


# ------------------------------------------------------------------------------
# 1. NORMALIZACIÓN DE TEXTO
# ------------------------------------------------------------------------------
def _prep(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _ascii_upper(s):
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c)).upper()


def _squish(s):
    return " ".join(s.split())


def map_unique(serie, f):
    """Aplica f solo a valores únicos (rápido en columnas grandes)."""
    codes, uniques = pd.factorize(serie)
    mapped = [f(v) for v in uniques]
    vacio = f(None)
    return pd.Series([mapped[c] if c >= 0 else vacio for c in codes], index=serie.index, dtype=object)


def _norm_fn(quitar, reemplazo):
    rx = re.compile(quitar)
    return lambda v: _squish(rx.sub(reemplazo, _ascii_upper(_prep(v))))


_f_sector = _norm_fn(r"[^A-Z0-9 -]", "")
_f_alcaldia = _norm_fn(r"[^A-Z ]", "")
_f_clean = _norm_fn(r"[^A-Z0-9 ]", " ")


def norm_sector(s):   return map_unique(s, _f_sector)
def norm_alcaldia(s): return map_unique(s, _f_alcaldia)
def clean_txt(s):     return map_unique(s, _f_clean)
def up_strip(s):      return map_unique(s, lambda v: _prep(v).strip().upper())


def norm_vehiculo(s):
    def f(v):
        x = _ascii_upper(_prep(v).strip())
        return x if x != "" else "EN BLANCO"
    return map_unique(s, f)


def num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def clean_name(c):
    c = unicodedata.normalize("NFKD", str(c))
    c = "".join(ch for ch in c if not unicodedata.combining(ch)).lower()
    c = re.sub(r"[^a-z0-9]+", "_", c).strip("_")
    return c or "col"


def regex_palabras(patrones):
    pats = sorted({p for p in patrones if p}, key=len, reverse=True)
    if not pats:
        return re.compile(r"a^")  # nunca coincide
    return re.compile(r"\b(" + "|".join(re.escape(p) for p in pats) + r")\b")


def convertir_fecha(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    n = pd.to_numeric(s, errors="coerce")
    a = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")  # serial de Excel
    b = pd.to_datetime(s.where(n.isna()), errors="coerce", dayfirst=True, format="mixed")
    return a.fillna(b)


# ------------------------------------------------------------------------------
# 2. CATÁLOGOS FIJOS
# ------------------------------------------------------------------------------
CATALOGO_ALCALDIAS_CLEAN = {_f_alcaldia(a) for a in [
    "ÁLVARO OBREGÓN", "AZCAPOTZALCO", "BENITO JUÁREZ", "COYOACÁN", "CUAJIMALPA",
    "CUAUHTÉMOC", "GUSTAVO A. MADERO", "IZTACALCO", "IZTAPALAPA",
    "MAGDALENA CONTRERAS", "MIGUEL HIDALGO", "MILPA ALTA", "TLÁHUAC", "TLALPAN",
    "VENUSTIANO CARRANZA", "XOCHIMILCO"]}

CATALOGO_VEHICULOS = {
    "AUTOBUS DE PASAJEROS", "AUTOMOVIL", "CAMION DE CARGA", "METROBUS", "TROLEBUS",
    "BICICLETA", "CAMIONETA", "OBJETO FIJO", "MOTOCICLETA", "MICROBUS", "TAXI", "SD",
    "EN BLANCO", "TRAILER", "TREN SUBURBANO", "FERROCARRIL", "TREN", "TREN LIGERO",
    "MONOPATIN"}

CATALOGO_MARCAS = [
    "NISSAN", "CHEVROLET", "VOLKSWAGEN", "FORD", "HONDA", "TOYOTA", "HYUNDAI", "KIA",
    "MAZDA", "MERCEDES BENZ", "MERCEDES-BENZ", "BMW", "AUDI", "SUZUKI", "DODGE", "JEEP",
    "CHRYSLER", "RAM", "PEUGEOT", "RENAULT", "SEAT", "FIAT", "MITSUBISHI", "SUBARU",
    "VOLVO", "ACURA", "INFINITI", "LEXUS", "PORSCHE", "MINI", "LINCOLN", "CADILLAC",
    "LAND ROVER", "JAGUAR", "ALFA ROMEO", "MASERATI", "GEO", "PONTIAC", "MERCURY",
    "HUMMER", "SAAB", "SMART", "SATURN", "MG", "JAC", "BAIC", "BYD", "CHANGAN", "GEELY",
    "OMODA", "JAECOO", "GAC", "JETOUR", "GWM", "GREAT WALL", "TANK", "HAVAL", "NETA",
    "ZEEKR", "DEEPAL", "FORTHING", "MAXUS", "SERES", "SEV", "TESLA", "CUPRA",
    "ASTON MARTIN", "FERRARI", "LAMBORGHINI", "MCLAREN", "LOTUS", "ITALIKA", "YAMAHA",
    "HONDA MOTO", "SUZUKI MOTO", "BMW MOTO", "BUELL", "DUCATI", "HARLEY DAVIDSON",
    "HARLEY-DAVIDSON", "KTM", "ROYAL ENFIELD", "TRIUMPH", "BAJAJ", "VENTO", "KURAZAI",
    "DINAMO", "HERO", "ISLO", "CARABELA", "VESPA", "TVS", "CFMOTO", "ZONTES",
    "HUSQVARNA", "APRILIA", "FREIGHTLINER", "KENWORTH", "INTERNATIONAL", "MACK", "ISUZU",
    "HINO", "VOLVO TRUCKS", "SCANIA", "MAN", "MERCEDES BENZ TRUCKS", "PETERBILT", "DINA",
    "FOTON", "SHACMAN", "FAW", "YUTONG", "KING LONG", "SITRAK"]
RX_MARCAS = regex_palabras(CATALOGO_MARCAS)

VALORES_SD = {"SD", "S/D", "S.D.", "SIN DATOS", "EN BLANCO", "NA", "N/A", ""}

CLASES_VALIDAS = {"VIA DE ACCESO CONTROLADO", "ACCESO CARRETERO", "EJE VIAL",
                  "VIA PRIMARIA", "VIA SECUNDARIA"}

# Patrones adicionales que complementan a los catálogos (editables)
EXTRA_VAC = ["PERIFERICO", "ANILLO PERIFERICO", "CIRCUITO INTERIOR", "RIO CHURUBUSCO",
             "VIADUCTO", "TLALPAN", "ZARAGOZA", "CONSTITUYENTES", "AQUILES SERDAN",
             "RIO SAN JOAQUIN"]
EXTRA_CARRETERA = ["CUERNAVACA", "MEXICO CUERNAVACA", "FED A CUERNAVACA", "CARR FED",
                   "CARR CUERNAVACA", "TOLUCA", "MEXICO TOLUCA", "PUEBLA", "MEXICO PUEBLA",
                   "PACHUCA", "MEXICO PACHUCA", "QUERETARO", "MEXICO QUERETARO",
                   "PICACHO AJUSCO", "AJUSCO", "XOCHIMILCO TULYEHUALCO", "OAXTEPEC",
                   "SAN BARTOLO XICOMULCO", "CARRETERA", "AUTOPISTA", "CARR"]
EXTRA_PRIMARIA = ["INSURGENTES", "INSURGENTES SUR", "INSURGENTES NORTE", "IZTACCIHUATL",
                  "GRANJAS", "ARMAS", "UNIVERSIDAD", "NUEVO LEON", "TLAHUAC", "BALDERAS",
                  "COYUYA", "REFORMA", "CHAPULTEPEC", "CIEN METROS"]
SIN_NOMBRE = {"", "SN", "S N", "SIN NOMBRE", "NO REFERENCIADO"}
RX_PREFIJO = re.compile(r"^(LATERAL DE|LATERALES DE|LATERALES|LATERAL|LAT|PROLONGACION|PROL|"
                        r"AVENIDA|AV|CALZADA|CALZ|CALLE|CARRETERA|CARR|AUTOPISTA|EJE)\s+")


# ------------------------------------------------------------------------------
# 3. CACHÉ Y LECTURA DE CATÁLOGOS
# ------------------------------------------------------------------------------
def _cache_path(nombre):
    d = os.path.join(APP_DIR, "cache")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, nombre)


def _cache_valido(cache, fuente):
    return os.path.exists(cache) and os.path.getmtime(cache) > os.path.getmtime(fuente)


def nombres_catalogo(archivo, columnas):
    """Lista de nombres limpios de un catálogo xlsx (con caché)."""
    ruta = ruta_catalogo(archivo)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el catálogo: {ruta}")
    cache = _cache_path(archivo + ".names.pkl")
    if _cache_valido(cache, ruta):
        with open(cache, "rb") as f:
            return pickle.load(f)
    log(f"  Leyendo catálogo {archivo} (solo la primera vez)...")
    df = pd.read_excel(ruta, engine="openpyxl")
    nombres = set()
    for c in columnas:
        if c in df.columns:
            nombres.update(clean_txt(df[c]).unique())
    nombres = sorted(n for n in nombres if len(n) >= 3)
    with open(cache, "wb") as f:
        pickle.dump(nombres, f)
    return nombres


def cargar_poligonos():
    """Cuadrantes como anillos numpy (exterior + huecos) con caja envolvente."""
    csv = ruta_catalogo("cuadrantes.csv")
    cache = _cache_path("poligonos.pkl")
    if _cache_valido(cache, csv):
        with open(cache, "rb") as f:
            return pickle.load(f)
    log("  Construyendo polígonos desde cuadrantes.csv (solo la primera vez)...")
    dfc = pd.read_csv(csv, encoding="utf-8")
    polys = []
    for _, r in dfc.iterrows():
        try:
            g = json.loads(r["geo_shape"])
        except Exception:
            continue
        if g.get("type") == "Polygon":
            partes = [g["coordinates"]]
        elif g.get("type") == "MultiPolygon":
            partes = g["coordinates"]
        else:
            continue
        sec, alc = _f_sector(r["sector"]), _f_alcaldia(r["alcaldia"])
        for parte in partes:
            anillos = [np.asarray(a, dtype=float)[:, :2] for a in parte]
            ext = anillos[0]
            polys.append({
                "sector": sec, "alcaldia": alc,
                "ext": ext, "huecos": anillos[1:],
                "bbox": (ext[:, 0].min(), ext[:, 1].min(), ext[:, 0].max(), ext[:, 1].max()),
            })
    with open(cache, "wb") as f:
        pickle.dump(polys, f)
    return polys


def sectores_catalogo():
    csv = ruta_catalogo("cuadrantes.csv")
    dfc = pd.read_csv(csv, encoding="utf-8")
    return set(norm_sector(dfc["sector"].dropna()).unique())


# ------------------------------------------------------------------------------
# 4. PUNTO EN POLÍGONO (ray casting vectorizado, sin dependencias externas)
# ------------------------------------------------------------------------------
def _en_anillo(px, py, ring):
    x, y = ring[:, 0], ring[:, 1]
    x2, y2 = np.roll(x, -1), np.roll(y, -1)
    dentro = np.zeros(px.shape, dtype=bool)
    with np.errstate(all="ignore"):
        for i in range(len(x)):
            if y[i] == y2[i]:
                continue
            cruza = (y[i] > py) != (y2[i] > py)
            xint = (x2[i] - x[i]) * (py - y[i]) / (y2[i] - y[i]) + x[i]
            dentro ^= cruza & (px < xint)
    return dentro


def ubicar_puntos(lon, lat, polys):
    """Devuelve, por punto, el índice del polígono que lo contiene (-1 si ninguno)."""
    lon = np.asarray(lon, dtype=float)
    lat = np.asarray(lat, dtype=float)
    res = np.full(len(lon), -1, dtype=int)
    for k, p in enumerate(polys):
        xmin, ymin, xmax, ymax = p["bbox"]
        cand = np.where((res < 0) & (lon >= xmin) & (lon <= xmax) & (lat >= ymin) & (lat <= ymax))[0]
        if cand.size == 0:
            continue
        px, py = lon[cand], lat[cand]
        dentro = _en_anillo(px, py, p["ext"])
        for h in p["huecos"]:
            dentro &= ~_en_anillo(px, py, h)
        res[cand[dentro]] = k
    return res


# ------------------------------------------------------------------------------
# 5. VALIDACIONES  (cada una recibe el DataFrame y devuelve errores o None)
# ------------------------------------------------------------------------------
_ctx = {}


def get_coords(df):
    if "coords" in _ctx:
        return _ctx["coords"]
    cx_raw = map_unique(df["coordenada_x"], lambda v: _prep(v).strip())
    cy_raw = map_unique(df["coordenada_y"], lambda v: _prep(v).strip())
    cx = pd.to_numeric(cx_raw, errors="coerce")
    cy = pd.to_numeric(cy_raw, errors="coerce")
    malas = cx.isna() | cy.isna() | (cx_raw == "") | (cy_raw == "")
    invalidas = df.loc[malas, ["folio", "coordenada_x", "coordenada_y"]].copy()
    invalidas["error"] = "Coordenada vacía, nula o con formato no numérico"

    num_df = df.loc[~malas].copy()
    cxo, cyo = cx[~malas], cy[~malas]
    num_df["lon"] = np.where(cxo < 0, cxo, cyo)
    num_df["lat"] = np.where(cxo < 0, cyo, cxo)
    fuera_mask = ((num_df["lon"] < LON_MIN) | (num_df["lon"] > LON_MAX) |
                  (num_df["lat"] < LAT_MIN) | (num_df["lat"] > LAT_MAX))
    fuera = num_df.loc[fuera_mask, ["folio", "coordenada_x", "coordenada_y"]].copy()
    fuera["error"] = "Coordenada fuera del territorio de la CDMX"

    _ctx["coords"] = {"reporte": pd.concat([invalidas, fuera], ignore_index=True),
                      "ok": num_df.loc[~fuera_mask]}
    return _ctx["coords"]


def chk_coordenadas(df):
    return get_coords(df)["reporte"]


def chk_espacial(df):
    ok = get_coords(df)["ok"]
    if ok.empty:
        return None
    polys = cargar_poligonos()
    idx = ubicar_puntos(ok["lon"].values, ok["lat"].values, polys)
    sec_real = [polys[i]["sector"] if i >= 0 else None for i in idx]
    alc_real = [polys[i]["alcaldia"] if i >= 0 else None for i in idx]
    sec_rep, alc_rep = norm_sector(ok["sector"]).tolist(), norm_alcaldia(ok["alcaldia"]).tolist()
    sec_raw, alc_raw = ok["sector"].tolist(), ok["alcaldia"].tolist()

    filas = []
    for j, (folio, lat, lon) in enumerate(zip(ok["folio"], ok["lat"], ok["lon"])):
        sr, ar, sp, ap = sec_real[j], alc_real[j], sec_rep[j], alc_rep[j]
        if sr is None:
            err = "La coordenada cae fuera de la cobertura de sectores CDMX"
        elif sp != sr and ap != ar:
            err = f"Ni el sector ni la alcaldía corresponden. Real: Sector {sr}, Alcaldía {ar}"
        elif sp != sr:
            err = f"El sector reportado ({_prep(sec_raw[j])}) no corresponde a la coordenada. Sector real: {sr}"
        elif ap != ar:
            err = f"La alcaldía reportada ({_prep(alc_raw[j])}) no corresponde a la coordenada. Alcaldía real: {ar}"
        else:
            continue
        filas.append({"folio": folio, "lat": lat, "lon": lon,
                      "alcaldia_reportada": _prep(alc_raw[j]), "alcaldia_real": ar,
                      "sector_reportado": _prep(sec_raw[j]), "sector_real": sr, "error": err})
    return pd.DataFrame(filas)


def chk_alcaldia(df):
    m = ~norm_alcaldia(df["alcaldia"]).isin(CATALOGO_ALCALDIAS_CLEAN)
    out = df.loc[m, ["folio", "alcaldia"]].copy()
    out["error"] = "Valor de alcaldía no registrado en el catálogo oficial CDMX"
    return out


def chk_sector(df):
    cat = sectores_catalogo()
    m = ~norm_sector(df["sector"]).isin(cat)
    out = df.loc[m, ["folio"] + [c for c in ["alcaldia"] if c in df.columns] + ["sector"]].copy()
    out["error"] = "Valor de sector no registrado en el catálogo de cuadrantes de la CDMX o vacío"
    return out


def chk_fecha(df):
    folio = map_unique(df["folio"], lambda v: _prep(v).strip())
    f_raw = folio.str.extract(r"(?<=/)(\d{6,8})(?=/)")[0]
    f_raw = f_raw.fillna(folio.str.extract(r"(\d{6,8})")[0])

    def a_fecha(r):
        if not isinstance(r, str):
            return None
        if len(r) == 8:
            return r
        if len(r) == 6:
            return "20" + r
        return None

    f_folio = f_raw.map(a_fecha)
    f_evento = df["fecha_evento"].dt.strftime("%Y%m%d")
    distinta = f_folio.isna() | f_evento.isna() | (f_folio != f_evento)
    out = df.loc[distinta, ["folio", "fecha_evento"]].copy()
    out["fecha_evento"] = out["fecha_evento"].dt.strftime("%Y-%m-%d")
    out["error"] = "Fecha evento no corresponde con la fecha del folio"
    return out


def chk_prioridad_baja(df):
    prio = up_strip(df["prioridad"])
    trasl = up_strip(df["trasladado_s_lesionado"])
    m = (prio == "BAJA") & ((num(df["total_occisos"]) > 0) | (~trasl.isin(["NO", ""])))
    out = df.loc[m, ["folio", "prioridad", "total_occisos", "trasladado_s_lesionado"]].copy()
    out["error"] = "Prioridad es BAJA pero registra occisos o trasladados"
    return out


def chk_tipo_vehiculo(df):
    cols = ["tipo_de_vehiculo_1", "tipo_de_vehiculo_2", "tipo_de_vehiculo_3"]
    bad = pd.Series(False, index=df.index)
    for c in cols:
        bad |= ~norm_vehiculo(df[c]).isin(CATALOGO_VEHICULOS)
    out = df.loc[bad, ["folio"] + cols].copy()
    out["error"] = "Categoría de vehículo no válida"
    return out


def chk_marca2(df):
    tipo2 = up_strip(df["tipo_de_vehiculo_2"])
    marca2 = up_strip(df["marca_de_vehiculo_2"])
    es_obj = tipo2.isin(["OBJETO FIJO", "SD", "S/D", "S.D.", "SIN DATOS", "EN BLANCO", ""]) | \
        tipo2.str.contains("OBJETO", regex=False)
    marca_ok = marca2.isin(CATALOGO_MARCAS) | marca2.map(lambda v: bool(RX_MARCAS.search(v)))
    m = es_obj & ~marca2.isin(VALORES_SD) & ~marca_ok
    cols = ["folio", "tipo_de_vehiculo_1", "tipo_de_vehiculo_2", "marca_de_vehiculo_1", "marca_de_vehiculo_2"]
    out = df.loc[m, cols].copy()
    out["error"] = "marca_vehiculo_2 no es una marca válida del catálogo ni un valor SD/NA"
    return out


def chk_moto_vs_vehiculo(df):
    afectado = (num(df["motociclista_occiso"]) > 0) | (num(df["motociclista_lesionado"]) > 0)
    m = afectado & (up_strip(df["tipo_de_vehiculo_1"]) != "MOTOCICLETA")
    out = df.loc[m, ["folio", "tipo_de_vehiculo_1", "motociclista_occiso", "motociclista_lesionado"]].copy()
    out["error"] = "Registra motociclista afectado pero Tipo de Vehículo 1 no es MOTOCICLETA"
    return out


def chk_lesionados_usuario(df):
    cols = ["peaton_lesionado", "ciclista_lesionado", "motociclista_lesionado",
            "pasajero_lesionado", "conductor_lesionado"]
    suma = sum(num(df[c]) for c in cols)
    m = num(df["total_lesionados"]) != suma
    out = df.loc[m, ["folio", "total_lesionados"] + cols].copy()
    out.insert(2, "suma_usuarios", suma[m])
    out["error"] = "Suma por usuario de la vía no coincide con total_lesionados"
    return out


def chk_lesionados_sexo(df):
    opciones = [["lesionados_femeninos", "lesionados_masculinos", "lesionados_se_desconoce"],
                ["lesionados_hombres", "lesionados_mujeres"]]
    cols = next((c for c in ([x for x in o if x in df.columns] for o in opciones) if len(c) >= 2), None)
    if cols is None:
        raise ValueError("No se encontraron las columnas de lesionados por sexo "
                         "(LESIONADOS FEMENINOS / MASCULINOS / SE DESCONOCE).")
    suma = sum(num(df[c]) for c in cols)
    m = num(df["total_lesionados"]) != suma
    out = df.loc[m, ["folio", "total_lesionados"] + cols].copy()
    out.insert(2, "suma_sexo", suma[m])
    out["error"] = "Suma por sexo no coincide con total_lesionados"
    return out


def chk_tipo_evento(df):
    def g(c):
        return num(df[c]) if c in df.columns else pd.Series(0.0, index=df.index)
    evento = map_unique(df["tipo_evento"], lambda v: _ascii_upper(_prep(v)))
    p_les, c_les, m_les, tot_les = g("peaton_lesionado"), g("ciclista_lesionado"), g("motociclista_lesionado"), g("total_lesionados")
    p_occ, c_occ, m_occ, tot_occ = g("peaton_occiso"), g("ciclista_occiso"), g("motociclista_occiso"), g("total_occisos")
    m = (
        (evento.str.contains("SIN LESIONADOS", regex=False) & (tot_les > 0)) |
        (evento.str.contains("CON LESIONADOS", regex=False) & (tot_les == 0) & (tot_occ == 0)) |
        (evento.str.contains("ATROPELLA", regex=False) &
         ((p_les + c_les + m_les + tot_les) == 0) & ((p_occ + c_occ + m_occ + tot_occ) == 0))
    )
    cols = [c for c in ["folio", "tipo_evento", "total_lesionados", "total_occisos",
                        "peaton_lesionado", "peaton_occiso"] if c in df.columns]
    out = df.loc[m, cols].copy()
    out["error"] = "Inconsistencia entre tipo_evento y los lesionados/occisos registrados"
    return out


def chk_folios_duplicados(df):
    f = map_unique(df["folio"], lambda v: _prep(v).strip())
    out = df.loc[f.duplicated(keep=False), ["folio"]].copy()
    out["error"] = "Folio duplicado en la base"
    return out


# --- Nombres de vialidad: se comparan por su "núcleo" (sin AV/CALZ/CALLE, sin artículos,
#     con NTE->NORTE, etc.) tanto en la base como en los catálogos, para que
#     "CALZ ACOXPA" == "Calzada Acoxpa" y "AV INSURGENTES NTE" == "Av. de los Insurgentes Norte".
_ABREV = {"CAM": "CAMINO", "NTE": "NORTE", "PTE": "PONIENTE", "OTE": "ORIENTE", "ORTE": "ORIENTE",
          "GRAL": "GENERAL", "STA": "SANTA", "STO": "SANTO"}
_PREF_MAYOR = {"AV", "AVENIDA", "CALZ", "CALZADA", "BLVD", "BOULEVARD", "BULEVAR", "CIRCUITO",
               "CIRC", "PROL", "PROLONGACION", "CONTINUACION", "CARR", "CARRETERA", "AUTOPISTA"}
_PREF_MENOR = {"CALLE", "CDA", "CERRADA", "PRIV", "PRIVADA", "ANDADOR", "RETORNO", "CJON",
               "CALLEJON", "PEATONAL", "PASAJE"}
_PREF_LAT = {"LATERAL", "LATERALES", "LAT"}
_ARTICULOS = {"DE", "DEL", "LA", "LAS", "LOS", "EL"}
_TITULOS = {"LIC", "LICENCIADO", "ING", "INGENIERO", "DR", "DOCTOR", "PROF", "PROFR", "PROFESOR", "MTRO", "MAESTRO",
            "GENERAL", "PDTE", "PRESIDENTE", "CNEL", "CORONEL", "CAP", "CAPITAN", "TTE", "TENIENTE", "SR", "SRA", "PDE", "PRES"}   # GRAL -> GENERAL
_TIPOVIAL_MAYOR = {"AVENIDA", "CALZADA", "CIRCUITO", "PROLONGACION", "CONTINUACION", "CORREDOR"}
RX_EJE_PREF = re.compile(r"^EJE (?:\d+ )?(?:NORTE|SUR|ORIENTE|PONIENTE|CENTRAL)(?: \d+)?\s+")


def tipo_y_nucleo(v):
    """Devuelve (tipo de prefijo: 'mayor' | 'menor' | 'sin', nombre núcleo)."""
    toks = [_ABREV.get(t, t) for t in (v or "").split()]
    tipo, i = "sin", 0
    while i < len(toks) and (toks[i] in _PREF_MAYOR or toks[i] in _PREF_MENOR
                             or toks[i] in _PREF_LAT or toks[i] in _ARTICULOS):
        if tipo == "sin" and toks[i] in _PREF_MAYOR:
            tipo = "mayor"
        elif tipo == "sin" and toks[i] in _PREF_MENOR:
            tipo = "menor"
        i += 1
    resto = [t for t in toks[i:] if t not in _ARTICULOS and t not in _TITULOS]
    resto = [t for k, t in enumerate(resto) if not (t == "A" and 0 < k < len(resto) - 1)]   # "CAMINO REAL A TOLUCA"
    return tipo, " ".join(resto)


def nucleos_catalogo(archivo, columnas, ejes=False):
    """Conjunto de nombres núcleo de un catálogo (con caché)."""
    ruta = ruta_catalogo(archivo)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el catálogo: {ruta}")
    cache = _cache_path(archivo + ".nucleos.pkl")
    if _cache_valido(cache, ruta):
        with open(cache, "rb") as f:
            return pickle.load(f)
    log(f"  Leyendo catálogo {archivo} (solo la primera vez)...")
    df = pd.read_excel(ruta, engine="openpyxl")
    nucleos = set()
    for c in columnas:
        if c not in df.columns:
            continue
        for n in clean_txt(df[c].dropna()).unique():
            nucleos.add(tipo_y_nucleo(n)[1])
            if ejes:   # "Eje 9 Sur Miguel Angel de Quevedo" -> también "MIGUEL ANGEL QUEVEDO"
                m = RX_EJE_PREF.match(n)
                if m:
                    nucleos.add(tipo_y_nucleo(n[m.end():])[1])
    nucleos = {n for n in nucleos if len(n) >= 3}
    with open(cache, "wb") as f:
        pickle.dump(nucleos, f)
    return nucleos


# --- Clase de cada vialidad según el campo CLASIF_15 de las capas SEMOVI --------------------------------
def clase_desde_clasif15(txt, defecto):
    """0 VAC, 2 eje vial, 3 primaria (arteria principal), 4 secundaria. 'Carriles laterales' o vacío -> clase de la capa."""
    t = _f_clean(txt or "")
    if "ACCESO CONTROLADO" in t:
        return 0
    if "ARTERIA" in t:
        return 3
    if t.startswith("EJE"):
        return 2
    if "SECUNDARIA" in t:
        return 4
    return defecto


def nucleos_por_clasif(archivo, columnas, defecto, ejes=False):
    """{clase: {nombres núcleo}} de un catálogo, clasificando cada fila por su clasif_15 (con caché)."""
    ruta = ruta_catalogo(archivo)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el catálogo: {ruta}")
    cache = _cache_path(archivo + ".clasif15.pkl")
    if _cache_valido(cache, ruta):
        with open(cache, "rb") as f:
            return pickle.load(f)
    log(f"  Leyendo catálogo {archivo} (solo la primera vez)...")
    df = pd.read_excel(ruta, engine="openpyxl")
    out = {0: set(), 1: set(), 2: set(), 3: set(), 4: set()}
    clases = df["clasif_15"] if "clasif_15" in df.columns else pd.Series([""] * len(df))
    for cl15, fila in zip(clases, df[[c for c in columnas if c in df.columns]].itertuples(index=False)):
        cl = clase_desde_clasif15("" if pd.isna(cl15) else str(cl15), defecto)
        for valor in fila:
            if pd.isna(valor):
                continue
            n = _f_clean(_prep(valor))
            if not n:
                continue
            out[cl].add(tipo_y_nucleo(n)[1])
            if ejes:
                m = RX_EJE_PREF.match(n)
                if m:
                    out[cl].add(tipo_y_nucleo(n[m.end():])[1])
    out = {k: {n for n in v if len(n) >= 3} for k, v in out.items()}
    with open(cache, "wb") as f:
        pickle.dump(out, f)
    return out


def nucleos_secundarias():
    """{nombre núcleo: True si existe con tipo Avenida/Calzada/Circuito...} (con caché)."""
    archivo = "vias_secundarias.xlsx"
    ruta = ruta_catalogo(archivo)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el catálogo: {ruta}")
    cache = _cache_path(archivo + ".nucleos.pkl")
    if _cache_valido(cache, ruta):
        with open(cache, "rb") as f:
            return pickle.load(f)
    log(f"  Leyendo catálogo {archivo} (solo la primera vez)...")
    df = pd.read_excel(ruta, engine="openpyxl", usecols=["NOMVIAL", "TIPOVIAL"])
    df = df.dropna(subset=["NOMVIAL"]).drop_duplicates()
    nombres = clean_txt(df["NOMVIAL"])
    mayor = df["TIPOVIAL"].map(lambda t: _ascii_upper(_prep(t)).strip() in _TIPOVIAL_MAYOR)
    res = {}
    for n, m in zip(nombres, mayor):
        core = tipo_y_nucleo(n)[1]
        if len(core) >= 3:
            res[core] = res.get(core, False) or bool(m)
    with open(cache, "wb") as f:
        pickle.dump(res, f)
    return res


# --- Georreferencia: proyección a UTM 14N y distancia punto-segmento contra la red vial -------------
def a_utm14(lon, lat):
    """lon/lat WGS84 (grados) -> (este, norte) UTM zona 14N en metros (serie de Krüger)."""
    a, f, k0 = 6378137.0, 1 / 298.257223563, 0.9996
    n = f / (2 - f)
    A = a / (1 + n) * (1 + n ** 2 / 4 + n ** 4 / 64)
    al = (n / 2 - 2 * n ** 2 / 3 + 5 * n ** 3 / 16, 13 * n ** 2 / 48 - 3 * n ** 3 / 5, 61 * n ** 3 / 240)
    phi, lam = np.radians(np.asarray(lat, float)), np.radians(np.asarray(lon, float) + 99.0)
    c = 2 * np.sqrt(n) / (1 + n)
    t = np.sinh(np.arctanh(np.sin(phi)) - c * np.arctanh(c * np.sin(phi)))
    xi0 = np.arctan2(t, np.cos(lam))
    eta0 = np.arctanh(np.sin(lam) / np.sqrt(1 + t ** 2))
    xi = xi0 + sum(al[j - 1] * np.sin(2 * j * xi0) * np.cosh(2 * j * eta0) for j in (1, 2, 3))
    eta = eta0 + sum(al[j - 1] * np.cos(2 * j * xi0) * np.sinh(2 * j * eta0) for j in (1, 2, 3))
    return 500000.0 + k0 * A * eta, k0 * A * xi


_CELDA = 200.0
_DESP = 1000


def cargar_red():
    """Carga red_vial.npz y arma un índice en cuadrícula (None si no existe el archivo)."""
    if "red" in _ctx:
        return _ctx["red"]
    ruta = ruta_catalogo("red_vial.npz")
    if not os.path.exists(ruta):
        _ctx["red"] = None
        return None
    z = np.load(ruta)
    red = {k: z[k] for k in ("x1", "y1", "x2", "y2", "clase", "tmayor", "c1", "c2", "c3")}
    red["origen"] = z["origen"]
    red["rot"] = z["rot"] if "rot" in z.files else None
    red["sub"] = z["sub"] if "sub" in z.files else None
    red["etiquetas"] = z["etiquetas"].tolist() if "etiquetas" in z.files else []
    red["vocab"] = {s_: i for i, s_ in enumerate(z["vocab"].tolist())}
    x1, y1, x2, y2 = (red[k].astype(np.float64) for k in ("x1", "y1", "x2", "y2"))
    cx0 = np.floor(np.minimum(x1, x2) / _CELDA).astype(np.int64) + _DESP
    cx1 = np.floor(np.maximum(x1, x2) / _CELDA).astype(np.int64) + _DESP
    cy0 = np.floor(np.minimum(y1, y2) / _CELDA).astype(np.int64) + _DESP
    cy1 = np.floor(np.maximum(y1, y2) / _CELDA).astype(np.int64) + _DESP
    nx, ny = cx1 - cx0 + 1, cy1 - cy0 + 1
    cnt = nx * ny
    seg = np.repeat(np.arange(len(x1)), cnt)
    loc = np.arange(cnt.sum()) - np.repeat(np.cumsum(cnt) - cnt, cnt)
    cxx = np.repeat(cx0, cnt) + loc % np.repeat(nx, cnt)
    cyy = np.repeat(cy0, cnt) + loc // np.repeat(nx, cnt)
    clave = cxx * 100000 + cyy
    orden = np.argsort(clave, kind="stable")
    red["clave"], red["seg_idx"] = clave[orden], seg[orden]
    # cuántos tramos mayores (VAC/carretero/eje/primaria) y secundarios tiene cada nombre (detecta homónimos)
    V_ = len(red["vocab"])
    mayor_seg, sec_seg = red["clase"] < 4, red["clase"] == 4
    red["cnt_mayor"] = sum(np.bincount(red[c][(red[c] >= 0) & mayor_seg], minlength=V_) for c in ("c1", "c2", "c3"))
    red["cnt_sec"] = sum(np.bincount(red[c][(red[c] >= 0) & sec_seg], minlength=V_) for c in ("c1", "c2", "c3"))
    _ctx["red"] = red
    return red


def consultar_red(red, lon, lat, nombre_ids, exige_mayor, tol):
    """Por punto: clases de vialidad a <= tol m (G), las que además coinciden con el nombre
    del PUNTO 1 (Gn) y la distancia mínima a cada clase (D). Clases: 0 VAC, 1 carretero, 2 eje, 3 primaria, 4 secundaria."""
    n = len(lon)
    G, Gn = np.zeros((n, 5), bool), np.zeros((n, 5), bool)
    D, Dn = np.full((n, 5), np.inf), np.full((n, 5), np.inf)
    J = np.full((n, 5), -1, np.int64)
    if n == 0:
        return G, Gn, D, Dn, J
    E, N = a_utm14(lon, lat)
    px, py = E - red["origen"][0], N - red["origen"][1]
    cx = np.floor(px / _CELDA).astype(np.int64) + _DESP
    cy = np.floor(py / _CELDA).astype(np.int64) + _DESP
    celda = cx * 100000 + cy
    orden = np.argsort(celda, kind="stable")
    unicas, ini = np.unique(celda[orden], return_index=True)
    fin = np.append(ini[1:], n)
    X1, Y1, X2, Y2 = (red[k] for k in ("x1", "y1", "x2", "y2"))
    clave, seg_idx = red["clave"], red["seg_idx"]
    for u, a, b in zip(unicas, ini, fin):
        ux, uy = divmod(int(u), 100000)
        cand = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                k = (ux + dx) * 100000 + (uy + dy)
                lo, hi = np.searchsorted(clave, k, "left"), np.searchsorted(clave, k, "right")
                if hi > lo:
                    cand.append(seg_idx[lo:hi])
        if not cand:
            continue
        s_ = np.unique(np.concatenate(cand))
        ax, ay = X1[s_].astype(np.float64), Y1[s_].astype(np.float64)
        vx, vy = X2[s_].astype(np.float64) - ax, Y2[s_].astype(np.float64) - ay
        v2 = vx * vx + vy * vy + 1e-9
        cl, c1, c2, c3 = red["clase"][s_], red["c1"][s_], red["c2"][s_], red["c3"][s_]
        tm = red["tmayor"][s_]
        ids = orden[a:b]
        for j in range(0, len(ids), 400):
            ii = ids[j:j + 400]
            wx, wy = px[ii, None] - ax[None, :], py[ii, None] - ay[None, :]
            t = np.clip((wx * vx + wy * vy) / v2, 0, 1)
            dist = np.hypot(wx - t * vx, wy - t * vy)
            nid = nombre_ids[ii][:, None]
            mism = (c1[None, :] == nid) | (c2[None, :] == nid) | (c3[None, :] == nid)
            # "CALZ X" / "AV X" no se compara con una "Calle X" de secundarias
            compat = ~exige_mayor[ii][:, None] | tm[None, :] | (cl != 4)[None, :]
            for c in range(5):
                m = cl == c
                if not m.any():
                    continue
                d = dist[:, m]
                D[ii, c] = d.min(axis=1)
                J[ii, c] = s_[m][d.argmin(axis=1)]
                Dn[ii, c] = np.where((mism & compat)[:, m], d, np.inf).min(axis=1)
    G, Gn = D <= tol, Dn <= tol
    return G, Gn, D, Dn, J


def cargar_reglas_nombre():
    """[(regex, clase normalizada)] desde catalogos/reglas_vialidad.csv (columnas: patron, clasificacion)."""
    if "reglas" in _ctx:
        return _ctx["reglas"]
    reglas = []
    ruta = ruta_catalogo("reglas_vialidad.csv")
    if USAR_REGLAS_NOMBRE and os.path.exists(ruta):
        r = pd.read_csv(ruta, encoding="utf-8-sig", dtype=str).fillna("")
        for pat, cl in zip(r.get("patron", []), r.get("clasificacion", [])):
            pat = _f_clean(pat)
            if pat and cl.strip():
                reglas.append((re.compile(r"(?<![A-Z0-9])" + re.escape(pat).replace("\\ ", " ") + r"(?![A-Z0-9])"),
                               _f_clean(cl)))
    _ctx["reglas"] = reglas
    return reglas


def cargar_anexo2():
    """[(nombres núcleo, regex multipalabra)] para (VAC, primaria) del Anexo 2; None si no hay archivo."""
    if "anexo2" in _ctx:
        return _ctx["anexo2"]
    ruta = ruta_catalogo("anexo2_reglamento_transito.csv")
    res = None
    if USAR_ANEXO2 and os.path.exists(ruta):
        r = pd.read_csv(ruta, encoding="utf-8-sig", dtype=str).fillna("")
        conj = {"VIA DE ACCESO CONTROLADO": set(), "VIA PRIMARIA": set()}
        for cl, nombres in zip(r["clase"], r["nombres"]):
            for nom in nombres.split("|"):
                core = tipo_y_nucleo(_f_clean(nom))[1]
                if len(core) >= 3 and _f_clean(cl) in conj:
                    conj[_f_clean(cl)].add(core)
        res = [(conj[k], regex_palabras([c for c in conj[k] if " " in c]))
               for k in ("VIA DE ACCESO CONTROLADO", "VIA PRIMARIA")]
    _ctx["anexo2"] = res
    return res


# --- Zonas viales SCT (KMZ con polígonos por alcaldía agrupados en Z1..Z5) ----------------------------
def cargar_zonas():
    """[(zona 'Z1'.., nombre alcaldía, [anillos (n,2) lon/lat])] desde catalogos/zonas_viales_sct.kmz."""
    if "zonas" in _ctx:
        return _ctx["zonas"]
    import zipfile
    import xml.etree.ElementTree as ET
    ruta = ruta_catalogo("zonas_viales_sct.kmz")
    res = []
    if USAR_ZONAS_VIALES and os.path.exists(ruta):
        with zipfile.ZipFile(ruta) as z:
            kml = next(n for n in z.namelist() if n.lower().endswith(".kml"))
            raiz = ET.fromstring(z.read(kml))
        ns = "{http://www.opengis.net/kml/2.2}"
        for pm in raiz.iter(ns + "Placemark"):
            dat = {d.get("name"): (d.findtext(ns + "value") or "") for d in pm.iter(ns + "Data")}
            zona = (dat.get("layer") or pm.findtext(ns + "name") or "").strip().upper()
            anillos = []
            for lr in pm.iter(ns + "LinearRing"):
                pts = [[float(a) for a in q.split(",")[:2]] for q in (lr.findtext(ns + "coordinates") or "").split()]
                if len(pts) >= 3:
                    anillos.append(np.array(pts))
            if zona and anillos:
                res.append((zona, dat.get("nomgeo", ""), anillos))
    _ctx["zonas"] = res
    return res


def _en_anillo(lon, lat, ring):
    """Ray casting vectorizado: puntos (lon, lat) dentro del anillo (n,2)."""
    x1, y1 = ring[:-1, 0], ring[:-1, 1]
    x2, y2 = ring[1:, 0], ring[1:, 1]
    dentro = np.zeros(len(lon), bool)
    for a in range(0, len(x1), 2000):
        sl = slice(a, a + 2000)
        cruza = ((y1[sl][None, :] > lat[:, None]) != (y2[sl][None, :] > lat[:, None]))
        xi = (x2[sl] - x1[sl])[None, :] * (lat[:, None] - y1[sl][None, :]) / (y2[sl] - y1[sl] + 1e-30)[None, :] + x1[sl][None, :]
        dentro ^= (cruza & (lon[:, None] < xi)).sum(axis=1) % 2 == 1
    return dentro


def zona_de_coordenadas(lon, lat):
    """Devuelve (zona, alcaldía) por punto; None si cae fuera de todas las zonas."""
    zonas = cargar_zonas()
    n = len(lon)
    zona, alc = np.full(n, None, dtype=object), np.full(n, None, dtype=object)
    for z, nombre, anillos in zonas:
        todos = np.vstack(anillos)
        cand = np.flatnonzero((lon >= todos[:, 0].min()) & (lon <= todos[:, 0].max()) &
                              (lat >= todos[:, 1].min()) & (lat <= todos[:, 1].max()) & (zona == None))  # noqa: E711
        if not len(cand):
            continue
        d = np.zeros(len(cand), bool)
        for r in anillos:                 # exterior + huecos: la paridad total resuelve ambos
            d ^= _en_anillo(lon[cand], lat[cand], r)
        zona[cand[d]], alc[cand[d]] = z, nombre
    return zona, alc


def chk_zona_vial(df):
    """Zona vial reportada (Z1-Z5) vs. polígono de zonas_viales_sct.kmz donde cae la coordenada."""
    zonas = cargar_zonas()
    vacio = pd.DataFrame(columns=["folio", "error"])
    if not zonas:
        log("  (no se encontró catalogos/zonas_viales_sct.kmz: se omite la validación de zona vial)")
        return vacio
    col = next((c for c in COLUMNAS_ZONA_VIAL if c in df.columns), None)
    if col is None:
        log("  (la base no trae una columna de zona vial: " + ", ".join(COLUMNAS_ZONA_VIAL) + ")")
        return vacio
    ok = get_coords(df)["ok"]
    zona, alc = zona_de_coordenadas(ok["lon"].values, ok["lat"].values)
    rep_raw = ok[col].map(_prep)
    rep = rep_raw.str.extract(r"(\d+)", expand=False)
    esp = pd.Series(zona, index=ok.index).str.extract(r"(\d+)", expand=False)
    fuera = pd.Series(zona, index=ok.index).isna()
    mal = (~fuera) & (rep != esp)
    out = ok.loc[mal, ["folio"]].copy()
    out["coordenada_x"], out["coordenada_y"] = ok.loc[mal, "coordenada_x"], ok.loc[mal, "coordenada_y"]
    out["zona_reportada"] = rep_raw[mal]
    out["zona_real"] = ("Z" + esp[mal]).values
    out["alcaldia_zona"] = pd.Series(alc, index=ok.index)[mal].values
    out["error"] = ("La zona vial reportada (" + out["zona_reportada"].replace("", "vacío") +
                    ") no corresponde a la coordenada. Zona real: " + out["zona_real"] +
                    " (" + out["alcaldia_zona"] + ")")
    sin = ok.loc[fuera, ["folio"]].copy()
    sin["coordenada_x"], sin["coordenada_y"] = ok.loc[fuera, "coordenada_x"], ok.loc[fuera, "coordenada_y"]
    sin["zona_reportada"], sin["zona_real"], sin["alcaldia_zona"] = rep_raw[fuera], "", ""
    sin["error"] = "La coordenada cae fuera de las zonas viales SCT"
    return pd.concat([out, sin], ignore_index=True)


def chk_vialidad(df):
    log("  Cargando catálogos de vialidades...")
    NUNCA = regex_palabras([])

    def prep_cat(nuc, extras=()):
        return (nuc, regex_palabras([n for n in nuc if " " in n]), regex_palabras(extras))

    # La clase de cada vialidad sale de su CLASIF_15 (no del archivo en que viene): p. ej. Av. Gran Canal es
    # "Vía secundaria" aunque esté en el archivo de primarias; Insurgentes Norte es VAC.
    por_clase = {k: set() for k in range(5)}
    for arch, cols, defecto, ejes_ in (("via_acceso_controlado.xlsx", ["Nombre", "nom_ruta"], 0, False),
                                       ("ejes_viales.xlsx", ["Nombre", "nom_ruta"], 2, True),
                                       ("vias_primarias.xlsx", ["Nombre", "nom_ruta"], 3, False)):
        for k, nombres in nucleos_por_clasif(arch, cols, defecto, ejes_).items():
            por_clase[k] |= nombres
    # Listas auxiliares: nombres de la red anular/viaducto (únicos de la ciudad) valen siempre; los nombres de destino
    # (TOLUCA, PUEBLA...) solo con prefijo de carretera, y el resto solo cuando el registro NO tiene coordenada,
    # para no confundir calles homónimas (Calle Puebla, Av. Toluca) con carreteras.
    VAC_FIJOS = ["PERIFERICO", "ANILLO PERIFERICO", "CIRCUITO INTERIOR", "VIADUCTO"]
    rx_vac_fuerte = regex_palabras(VAC_FIJOS)
    rx_vac_debil = regex_palabras([e for e in EXTRA_VAC if e not in VAC_FIJOS])
    rx_car_extra = regex_palabras(EXTRA_CARRETERA)
    rx_car_pref = re.compile(r"(?<![A-Z0-9])(CARR|CARRETERA|AUTOPISTA|AUT|FED|FEDERAL)(?![A-Z0-9])")
    rx_pri_debil = regex_palabras(EXTRA_PRIMARIA)
    vac = prep_cat(por_clase[0])
    eje = prep_cat(por_clase[2])
    car = prep_cat(nucleos_catalogo("acceso_carretero.xlsx", ["NOMVIAL"]))
    pri = prep_cat(por_clase[3])
    sec = dict(nucleos_secundarias())
    for nombre in por_clase[4]:          # vías "secundarias" que viajan en los archivos de vialidades mayores
        sec.setdefault(nombre, True)

    # PUNTO 1 es SIEMPRE la vialidad principal donde ocurrió el hecho
    cols = [c for c in VIALIDAD_COLUMNAS if c in df.columns]
    if not cols:
        raise ValueError("No se encontraron las columnas con el nombre de la vialidad "
                         f"({', '.join(VIALIDAD_COLUMNAS)}).")
    cols = cols[:1]
    log(f"  Columnas de vialidad usadas: {', '.join(cols)} (modo: {VIALIDAD_MODO})")

    c_rep = map_unique(df["clasificacion_de_la_vialidad"], _prep)
    c_clean = clean_txt(c_rep)

    def norm_clasif(c):
        c = c or ""
        if re.search(r"ACCESO CONTROLADO|VAC|ANULAR|RADIAL|VIADUCTO", c): return "VIA DE ACCESO CONTROLADO"
        if re.search(r"CARRETER|AUTOPISTA", c): return "ACCESO CARRETERO"
        if "EJE" in c: return "EJE VIAL"
        if re.search(r"PRIMAR|RIMAR", c): return "VIA PRIMARIA"
        if "SECUNDAR" in c: return "VIA SECUNDARIA"
        return c
    clasif_norm = map_unique(c_clean, norm_clasif)

    def hit(cat, v, n):
        s, rx_multi, rx_extra = cat
        return bool(n in s or rx_multi.search(n) or rx_extra.search(v) or rx_extra.search(n))

    # Nombres mal escritos ("CIRUNVALACION"): coincidencia aproximada contra los nombres de vialidades mayores
    core_clases = {}
    for k_, cat_ in enumerate((vac, car, eje, pri)):
        for nombre_ in cat_[0]:
            core_clases.setdefault(nombre_, []).append(k_)
    cores_lista = sorted(core_clases)

    def difuso(n):
        import difflib
        cand = [c for c in cores_lista if c[:1] == n[:1] and abs(len(c) - len(n)) <= 2]
        m = difflib.get_close_matches(n, cand, n=1, cutoff=0.9)
        return (m[0], core_clases[m[0]]) if m else None

    def flags(v):
        """(exacto VAC, carretero, eje, primaria) + (amplio x4) + (secundaria ok, 'CALLE X' en secundarias, prefijo EJE).
        Exacto = el núcleo del nombre es IGUAL al del catálogo; amplio = además contenido en el nombre o lista auxiliar."""
        v = v or ""
        tipo, n = tipo_y_nucleo(v)
        cats = (vac, car, eje, pri)
        eje_pref = v.startswith("EJE ")
        carr_pref = bool(rx_car_pref.search(v))

        def clases(n):
            ex = [n in c[0] for c in cats]
            am = [bool(n in c[0] or c[1].search(n)) for c in cats]
            am[0] = am[0] or bool(rx_vac_fuerte.search(v))
            am[1] = am[1] or bool(carr_pref and (rx_car_extra.search(v) or rx_car_extra.search(n)))
            ex[2], am[2] = ex[2] or eje_pref, am[2] or eje_pref
            return ex, am
        exacto, amplio = clases(n)
        debil = [bool(rx_vac_debil.search(v) or rx_vac_debil.search(n)), bool(not carr_pref and rx_car_extra.search(n)),
                 False, bool(rx_pri_debil.search(v) or rx_pri_debil.search(n))]
        if len(n) >= 7 and not any(amplio) and not any(debil) and n not in sec:
            corr = difuso(n)                       # (nombre corregido, clases) o None
            if corr:
                n = corr[0]
                exacto, amplio = clases(n)
        # "CALZ X" / "AV X" no se compara contra una "Calle X" del catálogo de secundarias
        sec_ok = (n in sec) and (tipo != "mayor" or sec[n])
        menor_sec = (tipo == "menor")      # calle, cerrada, andador...: secundaria por nombre (la geometría puede subirla)
        return tuple(exacto) + tuple(amplio) + (sec_ok, menor_sec, eje_pref) + tuple(debil)

    CLASES = ["VIA DE ACCESO CONTROLADO", "ACCESO CARRETERO", "EJE VIAL", "VIA PRIMARIA", "VIA SECUNDARIA"]
    EQUIV = {"ACCESO CARRETERO", "VIA PRIMARIA"}

    c0 = cols[0]
    reportadas = {c0: map_unique(df[c0], _prep)}
    v_clean = clean_txt(reportadas[c0])
    hay_nombre = ~v_clean.isin(SIN_NOMBRE)
    fl = map_unique(v_clean, flags)

    def col(k):
        return fl.map(lambda t, k=k: t[k]).astype(bool).values

    n_df = len(df)
    Ex = np.column_stack([col(k) for k in range(4)]) if n_df else np.zeros((0, 4), bool)
    Am = np.column_stack([col(k) for k in range(4, 8)]) if n_df else np.zeros((0, 4), bool)
    sec_ok, menor_sec, eje_pref = (col(8), col(9), col(10)) if n_df else (np.zeros(0, bool),) * 3
    Wk = np.column_stack([col(k) for k in range(11, 15)]) if n_df else np.zeros((0, 4), bool)

    # --- Georreferencia: coordenadas del hecho vs. red vial ---------------------------------
    G, Gn = np.zeros((n_df, 5), bool), np.zeros((n_df, 5), bool)
    D, Dn = np.full((n_df, 5), np.inf), np.full((n_df, 5), np.inf)
    J = np.full((n_df, 5), -1, np.int64)
    PX, PY = np.full(n_df, np.nan), np.full(n_df, np.nan)
    hay_geo = np.zeros(n_df, bool)
    red = cargar_red() if USAR_GEORREFERENCIA else None
    if USAR_GEORREFERENCIA and red is None:
        log("  (no se encontró catalogos/red_vial.npz: la clasificación se valida solo por nombre)")
    if red is not None and {"coordenada_x", "coordenada_y"} <= set(df.columns):
        ok = get_coords(df)["ok"]
        pos = df.index.get_indexer(ok.index)
        nid_u = map_unique(v_clean, lambda v: red["vocab"].get(tipo_y_nucleo(v)[1], -2))
        mayor_u = map_unique(v_clean, lambda v: tipo_y_nucleo(v)[0] == "mayor")
        g_, gn_, d_, dn_, j_ = consultar_red(red, ok["lon"].values, ok["lat"].values,
                                         nid_u.values[pos].astype(np.int64),
                                         mayor_u.values[pos].astype(bool), GEO_TOL_M)
        D[pos], Dn[pos], J[pos] = d_, dn_, j_
        PX[pos], PY[pos] = a_utm14(ok["lon"].values, ok["lat"].values)
        hay_geo[pos] = True
        log(f"  Georreferencia: {int(hay_geo.sum())} de {n_df} registros con coordenada válida (tolerancia {GEO_TOL_M} m)")

    tol_arr = np.array([GEO_TOL_CLASE.get(k, GEO_TOL_M) for k in range(5)], float)
    G, Gn = D <= tol_arr[None, :], Dn <= tol_arr[None, :]

    # --- Anexo 2 del Reglamento de Tránsito: radiales (VAC) y avenidas primarias por nombre ---
    Ax, Aa = np.zeros((n_df, 4), bool), np.zeros((n_df, 4), bool)
    anexo = cargar_anexo2()
    if anexo and n_df:
        def f_anexo(v):
            tp, n = tipo_y_nucleo(v or "")
            if tp == "menor":      # "CALLE X" no es la avenida X
                return (False, False, False, False)
            ex = [n in cj for cj, rx in anexo]
            am = [bool(e or rx.search(n)) for e, (cj, rx) in zip(ex, anexo)]
            if len(n) >= 7 and not any(am):        # nombre mal escrito: coincidencia aproximada con el Anexo
                import difflib
                cand = [c for cj, _ in anexo for c in cj if c[:1] == n[:1] and abs(len(c) - len(n)) <= 2]
                mm = difflib.get_close_matches(n, cand, n=1, cutoff=0.9)
                if mm:
                    ex = [mm[0] in cj for cj, rx in anexo]
                    am = list(ex)
            return tuple(ex) + tuple(am)
        fa = map_unique(v_clean, f_anexo)
        for j, k in enumerate((0, 3)):          # radiales = VAC (clase 0); avenidas primarias = clase 3
            Ax[:, k] = fa.map(lambda t, j=j: t[j]).astype(bool).values
            Aa[:, k] = fa.map(lambda t, j=j: t[2 + j]).astype(bool).values

    # Nombre vs. coordenada (calles homónimas): con coordenada válida, el nombre cuenta como vialidad mayor
    # solo si hay un tramo de ese nombre cerca; los "EJE N ..." se identifican por su propio nombre.
    Am = Am | (Wk & ~hay_geo[:, None])          # listas auxiliares débiles: solo sin coordenada
    cerca = (Dn[:, 0] <= NOMBRE_CERCA_VAC_M) | (Dn[:, 1:4] <= NOMBRE_CERCA_M).any(axis=1)
    abierta = ~hay_geo | cerca                      # coincidencia de nombre EXACTA: ¿hay un tramo de ese nombre cerca?
    puerta = np.repeat(abierta[:, None], 4, axis=1)
    puerta[:, 2] |= eje_pref

    def combinar(exacto, amplio):
        """Coincidencia EXACTA (con la puerta de homónimos) tiene prioridad; si no hay, valen las coincidencias
        por contención/listas auxiliares (nombres como "ANILLO PERIFERICO" o "VIADUCTO")."""
        e = exacto & puerta
        return np.where(e.any(axis=1)[:, None], e, amplio & ~exacto)

    def armar(mayor):
        F5 = np.zeros((n_df, 5), bool)
        F5[:, :4] = mayor
        F5[:, 4] = sec_ok & ~mayor.any(axis=1)
        F5[menor_sec, :4] = False
        F5[menor_sec, 4] = True
        return F5

    # El Anexo 2 (texto legal del Reglamento) manda sobre las capas en cuanto al NOMBRE: si el nombre coincide EXACTO
    # con una ruta del Anexo, solo cuentan las clases del Anexo (p. ej. "EJERCITO NACIONAL MEXICANO" = primaria,
    # aunque otro tramo con ese nombre figure como VAC en las capas).
    Ax_ok = Ax & puerta
    usa_ax = Ax_ok.any(axis=1)
    # Rutas del Anexo con varias clases (p. ej. Insurgentes Norte: radial VAC y avenida primaria por tramos):
    # si hay un tramo de ESE nombre junto a la coordenada, vale la clase (clasif_15) de ese tramo.
    en_tramo = Ax_ok & Gn[:, :4]
    Ax_ok = np.where(en_tramo.any(axis=1)[:, None], en_tramo, Ax_ok)
    Fn_base = armar(combinar(Ex, Am))                                              # solo catálogos
    Fn = armar(np.where(usa_ax[:, None], Ax_ok, combinar(Ex, Am | Aa)))            # catálogos + Anexo 2

    # La secundaria del nombre (geometría) solo cuenta si no hay clase mayor y el prefijo es compatible
    tipo = map_unique(v_clean, lambda v: tipo_y_nucleo(v)[0]).values
    Gn = Gn.copy()
    Gn[:, 4] &= ~Gn[:, :4].any(axis=1)
    Gn[tipo == "menor", :4] = False

    # Un nombre con varias clases (p. ej. "CIRCUNVALACION": eje y primaria): si hay un tramo de ESE nombre junto a la
    # coordenada, valen las clases de ese tramo (el Anexo 2 conserva su propia regla).
    resto = Fn[:, :4] & Gn[:, :4]
    usar = resto.any(axis=1) & ~usa_ax
    Fn = Fn.copy()
    Fn[usar, :4] = resto[usar]

    # Decisión: clase de MAYOR JERARQUÍA entre la evidencia por nombre (catálogos/Anexo 2 + tramo del mismo nombre
    # cerca) y la geometría (vialidades a <= GEO_TOL_M de la coordenada).
    nombre_ev = Fn | Gn
    tiene_nombre = nombre_ev.any(axis=1)
    idx_nom = np.where(tiene_nombre, nombre_ev.argmax(axis=1), 9)
    tiene_geo = hay_geo & G.any(axis=1)
    idx_geo = np.where(tiene_geo, G.argmax(axis=1), 9)
    # Si el PUNTO 1 es una vialidad MAYOR por nombre (VAC, carretero, eje, primaria) manda su clase: un cruce con
    # otra vialidad no la cambia. Si el nombre es una calle secundaria o no se reconoce, la vialidad de mayor
    # jerarquía junto a la coordenada define la clase (p. ej. el punto entre el Eje 4 Norte y Av. Gran Canal).
    idx_fin = np.where(idx_nom <= 3, idx_nom, np.minimum(idx_nom, idx_geo))
    esperado = np.zeros((n_df, 5), bool)
    sel = idx_fin < 9
    esperado[np.flatnonzero(sel), idx_fin[sel]] = True
    criterio = np.where(idx_nom <= 3, "nombre del PUNTO 1",
                np.where(idx_geo < idx_nom, "georreferencia: vialidad de mayor jerarquía junto a la coordenada",
                         "nombre del PUNTO 1 y georreferencia"))
    criterio = np.where(tiene_nombre & usa_ax & (idx_nom <= idx_geo), "Anexo 2 del Reglamento de Tránsito (nombre del PUNTO 1)", criterio)
    # Reglas manuales por nombre: mandan sobre todo lo anterior
    reglas = cargar_reglas_nombre()
    con_regla = np.zeros(n_df, bool)
    if reglas:
        def regla(v):
            for rx, cl in reglas:
                if rx.search(v or ""):
                    return norm_clasif(cl)
            return ""
        rc = map_unique(v_clean, regla).values
        con_regla = rc != ""
        if con_regla.any():
            for k, clase in enumerate(CLASES):
                esperado[con_regla, k] = rc[con_regla] == clase
            criterio = np.where(con_regla, "regla manual (reglas_vialidad.csv)", criterio)
    hay_esperado = esperado.any(axis=1)

    cn = clasif_norm.values
    rep_ok = np.column_stack([
        (cn == clase) | ((cn == "ACCESO CARRETERO") | (cn == "VIA PRIMARIA")) & (clase in EQUIV)
        for clase in CLASES])
    soportada = (esperado & rep_ok).any(axis=1)

    # Sin evidencia: se acepta una clasificación válida; si está vacía/inválida se pide corregir
    invalida = (~clasif_norm.isin(CLASES_VALIDAS)).values
    inconsistente = (hay_esperado & ~soportada) | (~hay_esperado & hay_nombre.values & invalida)

    anexo_vac = (anexo[0][0] if anexo else set())

    inv_voc = None
    if red is not None and red.get("rot") is not None:
        inv_voc = {i_: n_ for n_, i_ in red["vocab"].items()}
        clase_idx = {k_: np.flatnonzero(red["clase"] == k_) for k_ in range(4)}

    def nombre_vac(p1):
        n1 = tipo_y_nucleo(p1)[1]
        if "VIADUCTO" in n1:
            return "VAC VIADUCTO"
        if "PERIFERICO" in n1 or "INTERIOR" in n1:
            return "VAC ANULAR"
        if n1 in anexo_vac:
            return "VAC RADIAL"
        return ""

    def tramo(i):
        """(etiqueta, subtipo VAC) del tramo de la clase sugerida cuyo nombre más se parece al PUNTO 1."""
        if inv_voc is None or not esperado[i].any() or not hay_geo[i]:
            return "", ""
        k = int(esperado[i].argmax())
        if k == 4:
            return "", ""
        idx_k = clase_idx[k]
        px, py = PX[i] - red["origen"][0], PY[i] - red["origen"][1]
        ax, ay = red["x1"][idx_k].astype(float), red["y1"][idx_k].astype(float)
        vx, vy = red["x2"][idx_k].astype(float) - ax, red["y2"][idx_k].astype(float) - ay
        t = np.clip(((px - ax) * vx + (py - ay) * vy) / (vx * vx + vy * vy + 1e-9), 0, 1)
        d = np.hypot(px - ax - t * vx, py - ay - t * vy)
        cand = np.flatnonzero(d <= tol_arr[k])
        if not len(cand):
            return "", ""
        toks = set(tipo_y_nucleo(v_clean.iloc[i])[1].split())
        mejor, mejor_clave = None, None
        for c_ in cand:
            j = idx_k[c_]
            nombres = " ".join(inv_voc.get(int(red[cc][j]), "") for cc in ("c1", "c2", "c3") if red[cc][j] >= 0)
            clave = (-len(toks & set(nombres.split())), d[c_])
            if mejor_clave is None or clave < mejor_clave:
                mejor, mejor_clave = j, clave
        j = mejor
        lab = red["etiquetas"][red["rot"][j]] if red["rot"][j] >= 0 else ""
        sub = ""
        if k == 0:
            sub = {1: "VAC ANULAR", 2: "VAC RADIAL", 3: "VAC VIADUCTO"}.get(int(red["sub"][j]), "") or nombre_vac(lab)
        return lab, sub

    def sugerir(i):
        vals = [CLASES[k] for k in range(5) if esperado[i, k]]
        if not vals:
            return "VIA SECUNDARIA"
        if esperado[i, 0]:
            return tramo(i)[1] or nombre_vac(v_clean.iloc[i]) or vals[0]
        return " / ".join(vals)

    def p1_sugerido(i):
        if idx_nom[i] <= 3 or con_regla[i]:        # el nombre del PUNTO 1 ya es una vialidad mayor: no se cambia
            return ""
        lab = tramo(i)[0]
        if lab and tipo_y_nucleo(lab)[1] != tipo_y_nucleo(v_clean.iloc[i])[1]:
            return lab
        return ""

    def cercanas(i):
        if not hay_geo[i]:
            return "(sin coordenada válida)"
        vals = [f"{CLASES[k]} ({D[i, k]:.0f} m)" for k in np.argsort(D[i]) if D[i, k] <= tol_arr[k]]
        return "; ".join(vals) if vals else f"(ninguna a <= {GEO_TOL_M} m)"

    idx_pos = np.flatnonzero(inconsistente)
    idx = df.index[idx_pos]
    out = pd.DataFrame({"folio": df.loc[idx, "folio"]})
    out[c0] = reportadas[c0][idx]
    out["clasificacion_reportada"] = c_rep[idx]
    out["clasificacion_sugerida"] = [sugerir(i) for i in idx_pos]
    out["punto_1_sugerido"] = [p1_sugerido(i) for i in idx_pos]
    for cc in ("coordenada_x", "coordenada_y"):          # coordenadas tal como vienen en la base
        if cc in df.columns:
            out[cc] = df.loc[idx, cc].values
    out["error"] = "La clasificación de la vialidad registrada no corresponde a la sugerida"
    return out


# ------------------------------------------------------------------------------
# 6. REGISTRO DE VALIDACIONES
# ------------------------------------------------------------------------------
VALIDACIONES = [
    ("Coordenadas", "Coordenadas vacías, no numéricas o fuera de CDMX",
     ["coordenada_x", "coordenada_y"], chk_coordenadas),
    ("Sector_Alcaldia_GIS", "Sector/alcaldía vs. polígono donde cae la coordenada",
     ["coordenada_x", "coordenada_y", "sector", "alcaldia"], chk_espacial),
    ("Alcaldia", "Alcaldía fuera del catálogo", ["alcaldia"], chk_alcaldia),
    ("Sector", "Sector fuera del catálogo de cuadrantes", ["sector"], chk_sector),
    ("Fecha_vs_Folio", "Fecha del evento vs. fecha contenida en el folio", ["fecha_evento"], chk_fecha),
    ("Prioridad_Baja", "Prioridad BAJA con occisos o trasladados",
     ["prioridad", "total_occisos", "trasladado_s_lesionado"], chk_prioridad_baja),
    ("Tipo_Vehiculo", "Tipo de vehículo fuera del catálogo",
     ["tipo_de_vehiculo_1", "tipo_de_vehiculo_2", "tipo_de_vehiculo_3"], chk_tipo_vehiculo),
    ("Marca_Vehiculo_2", "Marca del vehículo 2 inválida (objeto fijo / SD)",
     ["tipo_de_vehiculo_1", "tipo_de_vehiculo_2", "marca_de_vehiculo_1", "marca_de_vehiculo_2"], chk_marca2),
    ("Moto_vs_Vehiculo1", "Motociclista afectado y vehículo 1 no es motocicleta",
     ["motociclista_occiso", "motociclista_lesionado", "tipo_de_vehiculo_1"], chk_moto_vs_vehiculo),
    ("Lesionados_Usuario", "Total lesionados vs. suma por usuario de la vía",
     ["total_lesionados", "peaton_lesionado", "ciclista_lesionado", "motociclista_lesionado",
      "pasajero_lesionado", "conductor_lesionado"], chk_lesionados_usuario),
    ("Lesionados_Sexo", "Total lesionados vs. hombres + mujeres",
     ["total_lesionados"], chk_lesionados_sexo),
    ("Tipo_Evento", "Tipo de evento vs. lesionados/occisos",
     ["tipo_evento", "total_lesionados", "total_occisos"], chk_tipo_evento),
    ("Zona_Vial_GIS", "Zona vial SCT vs. polígono donde cae la coordenada",
     ["coordenada_x", "coordenada_y"], chk_zona_vial),
    ("Clasif_Vialidad", "Clasificación de la vialidad vs. catálogos oficiales",
     ["clasificacion_de_la_vialidad"], chk_vialidad),
    ("Folios_Duplicados", "Folios repetidos", ["folio"], chk_folios_duplicados),
]


# ------------------------------------------------------------------------------
# 7. ENTRADA / SALIDA
# ------------------------------------------------------------------------------
def elegir_base():
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    log("Selecciona la base a validar (.xlsx)...")
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ruta = filedialog.askopenfilename(title="Selecciona la base a validar",
                                          filetypes=[("Excel", "*.xlsx *.xlsm")])
        root.destroy()
    except Exception:
        ruta = input("Ruta completa del archivo .xlsx: ").strip().strip('"')
    if not ruta:
        raise SystemExit("No se seleccionó ningún archivo.")
    return ruta


def escribir_reporte(resumen, resultados, archivo):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(archivo, engine="openpyxl") as xw:
        resumen.to_excel(xw, sheet_name="Resumen", index=False)
        for nombre, d in resultados.items():
            d.to_excel(xw, sheet_name=nombre[:31], index=False)
        for ws in xw.book.worksheets:
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E78")
                c.alignment = Alignment(horizontal="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(ws.columns, start=1):
                ancho = max((len(str(c.value)) for c in list(col)[:500] if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(i)].width = min(max(ancho + 2, 10), 70)


def main():
    log("\n" + "=" * 62)
    log("  VALIDADOR DE BASES SSC")
    log("=" * 62)

    ruta_base = elegir_base()
    log(f"Base: {ruta_base}")

    ssc = pd.read_excel(ruta_base, engine="openpyxl")
    ssc.columns = [clean_name(c) for c in ssc.columns]
    if "folio" not in ssc.columns:
        raise ValueError("La base no tiene la columna 'folio'.")
    if "fecha_evento" in ssc.columns:
        ssc["fecha_evento"] = convertir_fecha(ssc["fecha_evento"])
    ssc = ssc[ssc["folio"].notna()].reset_index(drop=True)
    log(f"Registros a validar: {len(ssc)}")

    resultados, filas = {}, []
    for hoja, desc, cols, fn in VALIDACIONES:
        log(f"-> {hoja}")
        faltan = [c for c in cols if c not in ssc.columns]
        if faltan:
            log(f"   OMITIDA: faltan columnas: {', '.join(faltan)}")
            filas.append((hoja, desc, "OMITIDA (faltan: " + ", ".join(faltan) + ")", None))
            continue
        try:
            res = fn(ssc)
        except Exception as e:
            log(f"   FALLÓ: {e}")
            filas.append((hoja, desc, f"FALLÓ: {e}", None))
            continue
        n = 0 if res is None else len(res)
        filas.append((hoja, desc, "OK" if n == 0 else "CON ERRORES", n))
        if n > 0:
            resultados[hoja] = res.reset_index(drop=True)
        log(f"   {n} error(es)")

    resumen = pd.DataFrame(filas, columns=["Validacion", "Descripcion", "Estado", "Errores"])
    resumen["Errores"] = resumen["Errores"].astype("Int64")
    dir_salida = os.path.join(APP_DIR, "resultados")
    os.makedirs(dir_salida, exist_ok=True)
    base = re.sub(r"[^A-Za-z0-9_-]", "_", os.path.splitext(os.path.basename(ruta_base))[0])
    estampa = dt.datetime.now().strftime("%Y%m%d_%H%M")
    archivo = os.path.join(dir_salida, f"validacion_{base}_{estampa}.xlsx")
    escribir_reporte(resumen, resultados, archivo)

    if SEPARAR_ARCHIVOS:
        for nombre, d in resultados.items():
            d.to_excel(os.path.join(dir_salida, f"{nombre}_{base}_{estampa}.xlsx"), index=False)

    log("\n" + "=" * 62)
    log("  RESUMEN")
    log("=" * 62)
    log(resumen[["Validacion", "Estado", "Errores"]].to_string(index=False))
    log(f"\nReporte generado en:\n  {archivo}")
    if not resultados:
        log("\nNo se encontraron errores.")
    if ABRIR_CARPETA and os.name == "nt":
        os.startfile(dir_salida)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code:
            log(str(e.code))
    except Exception:
        log("\nERROR INESPERADO:")
        traceback.print_exc()
    if getattr(sys, "frozen", False):
        input("\nPresiona Enter para salir...")
