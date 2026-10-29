# -*- coding: utf-8 -*-
"""
CONVERTIR RED VIAL (se corre UNA sola vez, o cuando cambien los shapefiles)

Lee los shapefiles del Marco Geoestadístico SEMOVI/SSC y genera catalogos/red_vial.npz,
un archivo ligero con los segmentos de cada vialidad (en metros, UTM 14N) que usa
validador.py para validar la clasificación de la vialidad por georreferencia.
No necesita geopandas ni GDAL: solo numpy (lector de .shp/.dbf incluido).

Uso:
  python convertir_red_vial.py "C:\\ruta\\MarcoGeoestadistico_SEMOVI_SSC"
(la carpeta que contiene Acceso_carretero, Ejes_viales, VACs, Vias_primarias, Vias_secundarias)
"""
import os
import re
import sys
import struct

import zipfile
import xml.etree.ElementTree as ET

import numpy as np

import validador as V

# (clase, subcarpeta/archivo sin extensión, campos con nombres)
CAPAS = [
    (0, "VACs/via_acceso_controlado", ["Nombre", "nom_ruta"], True),
    (1, "Acceso_carretero/acceso_carretero", ["NOMVIAL"], False),
    (2, "Ejes_viales/ejes_viales", ["Nombre", "nom_ruta"], True),
    (3, "Vias_primarias/vias_primarias", ["Nombre", "nom_ruta"], False),
    (4, "Vias_secundarias/vias_secundarias", ["NOMVIAL", "TIPOVIAL"], False),
]
ORIGEN_X, ORIGEN_Y = 460000.0, 2110000.0   # para guardar en float32 sin perder precisión


def leer_dbf(ruta, campos):
    with open(ruta + ".dbf", "rb") as f:
        b = f.read()
    n, = struct.unpack("<I", b[4:8])
    hl, rl = struct.unpack("<HH", b[8:12])
    pos, i, offs = 1, 32, {}
    while b[i] != 0x0D:
        nombre = b[i:i + 11].split(b"\0")[0].decode("ascii", "ignore")
        largo = b[i + 16]
        offs[nombre] = (pos, largo)
        pos += largo
        i += 32
    cod = "utf-8"
    cpg = ruta + ".cpg"
    if os.path.exists(cpg):
        cod = open(cpg).read().strip() or "utf-8"
        cod = "utf-8" if cod.upper() in ("UTF-8", "UTF8") else cod
    salida = []
    for r in range(n):
        reg = b[hl + r * rl: hl + (r + 1) * rl]
        fila = {}
        for c in campos:
            if c in offs:
                p, l = offs[c]
                fila[c] = reg[p:p + l].decode(cod, "replace").strip()
        salida.append(fila)
    return salida


def leer_shp_lineas(ruta):
    """Devuelve, por registro, una lista de arrays (n,2) (una por parte)."""
    with open(ruta + ".shp", "rb") as f:
        b = f.read()
    pos, total = 100, len(b)
    registros = []
    while pos + 8 <= total:
        _, largo = struct.unpack(">ii", b[pos:pos + 8])
        c0, c1 = pos + 8, pos + 8 + largo * 2
        tipo, = struct.unpack("<i", b[c0:c0 + 4])
        partes = []
        if tipo in (3, 13, 23):
            npart, npts = struct.unpack("<ii", b[c0 + 36:c0 + 44])
            idx = np.frombuffer(b, dtype="<i4", count=npart, offset=c0 + 44)
            pts = np.frombuffer(b, dtype="<f8", count=npts * 2, offset=c0 + 44 + 4 * npart).reshape(-1, 2)
            lim = list(idx) + [npts]
            partes = [pts[lim[k]:lim[k + 1]] for k in range(npart)]
        registros.append(partes)
        pos = c1
    return registros


def leer_kmz_ejes(ruta):
    """Ejes extra en KMZ/KML (lon,lat): devuelve [(nombre, [array (n,2) en UTM 14N, ...]), ...]."""
    if ruta.lower().endswith(".kmz"):
        with zipfile.ZipFile(ruta) as z:
            kml = next(n for n in z.namelist() if n.lower().endswith(".kml"))
            raiz = ET.fromstring(z.read(kml))
    else:
        raiz = ET.parse(ruta).getroot()
    ns = "{http://www.opengis.net/kml/2.2}"
    salida = []
    for pm in raiz.iter(ns + "Placemark"):
        nombre = pm.findtext(ns + "name") or ""
        lineas = []
        for ls in pm.iter(ns + "LineString"):
            pts = np.array([[float(a) for a in q.split(",")[:2]] for q in ls.findtext(ns + "coordinates").split()])
            e, n = V.a_utm14(pts[:, 0], pts[:, 1])
            lineas.append(np.c_[e, n])
        salida.append((nombre, lineas))
    return salida


def main():
    if len(sys.argv) < 2 or not os.path.isdir(sys.argv[1]):
        sys.exit(__doc__)
    raiz = sys.argv[1]
    vocab, X1, Y1, X2, Y2, CL, C1, C2, C3, TM = {}, [], [], [], [], [], [], [], [], []
    ROT, SUB, labels = [], [], {}

    def etiqueta(clase_f, a):
        """Nombre con el que se sugiere corregir el PUNTO 1 (estilo de la base)."""
        nr, nb = V._f_clean(a.get("nom_ruta", "")), V._f_clean(a.get("Nombre", ""))
        nv = V._f_clean(a.get("NOMVIAL", ""))
        if clase_f == 2:
            lab = nr
            if not lab:
                m = V.RX_EJE_PREF.match(nb)
                lab = nb[:m.end()].strip() if m else nb
            for lg, ab in ((" NORTE", " NTE"), (" ORIENTE", " OTE"), (" PONIENTE", " PTE")):
                if lab.endswith(lg):
                    lab = lab[:-len(lg)] + ab
            return lab
        if clase_f == 0:
            if V._f_clean(a.get("tipo_via", "")) == "VAC ANULAR":
                return " ".join(nb.split()[:2])               # ANILLO PERIFERICO / CIRCUITO INTERIOR
            return nb
        if clase_f == 1:
            return nv
        return nb if clase_f == 3 else ""

    def lab_id(txt):
        return labels.setdefault(txt, len(labels)) if txt else -1

    def cid(core):
        if not core or len(core) < 3:
            return -1
        return vocab.setdefault(core, len(vocab))

    for clase, rel, campos, es_eje in CAPAS:
        ruta = os.path.join(raiz, *rel.split("/"))
        print(f"Leyendo {rel} ...", flush=True)
        attrs = leer_dbf(ruta, campos + ["clasif_15", "tipo_via"])
        geoms = leer_shp_lineas(ruta)
        assert len(attrs) == len(geoms), f"{rel}: dbf y shp no coinciden ({len(attrs)} vs {len(geoms)})"
        n_seg = 0
        for a, partes in zip(attrs, geoms):
            cores = []
            tmayor = V._f_clean(a.get("TIPOVIAL", "")) in V._TIPOVIAL_MAYOR
            clase_f = V.clase_desde_clasif15(a.get("clasif_15", ""), clase)   # clase del tramo según clasif_15
            rot_id = lab_id(etiqueta(clase_f, a))
            sub_id = {"VAC ANULAR": 1, "VAC RADIAL": 2, "VAC VIADUCTO": 3}.get(V._f_clean(a.get("tipo_via", "")), 0) if clase_f == 0 else 0
            for c in [c for c in campos if c not in ("TIPOVIAL", "clasif_15")]:
                t = V._f_clean(a.get(c, ""))
                if not t:
                    continue
                cores.append(V.tipo_y_nucleo(t)[1])
                if es_eje:
                    m = V.RX_EJE_PREF.match(t)
                    if m:
                        cores.append(V.tipo_y_nucleo(t[m.end():])[1])
            ids = []
            for k in cores:
                i = cid(k)
                if i >= 0 and i not in ids:
                    ids.append(i)
            ids = (ids + [-1, -1, -1])[:3]
            for p in partes:
                if len(p) < 2:
                    continue
                X1.append(p[:-1, 0]); Y1.append(p[:-1, 1]); X2.append(p[1:, 0]); Y2.append(p[1:, 1])
                k = len(p) - 1
                CL.append(np.full(k, clase_f, np.int8))
                C1.append(np.full(k, ids[0], np.int32)); C2.append(np.full(k, ids[1], np.int32))
                C3.append(np.full(k, ids[2], np.int32))
                TM.append(np.full(k, tmayor, bool))
                ROT.append(np.full(k, rot_id, np.int32)); SUB.append(np.full(k, sub_id, np.int8))
                n_seg += k
        print(f"   {len(attrs)} vialidades, {n_seg} segmentos")

    # Ejes extra (KMZ): se agregan como clase EJE VIAL
    kmz = os.path.join(V.APP_DIR, "catalogos", "ejes_viales_extra.kmz")
    if os.path.exists(kmz):
        print("Leyendo ejes_viales_extra.kmz ...", flush=True)
        n_seg = 0
        for nombre, lineas in leer_kmz_ejes(kmz):
            ids = ([cid(V.tipo_y_nucleo(V._f_clean(nombre))[1])] + [-1, -1, -1])[:3]
            for p in lineas:
                if len(p) < 2:
                    continue
                X1.append(p[:-1, 0]); Y1.append(p[:-1, 1]); X2.append(p[1:, 0]); Y2.append(p[1:, 1])
                k = len(p) - 1
                CL.append(np.full(k, 2, np.int8)); TM.append(np.full(k, False, bool))
                lab_k = V._f_clean(nombre)
                for lg, ab in ((" ORIENTE", " OTE"), (" PONIENTE", " PTE"), (" NORTE", " NTE")):
                    if lab_k.endswith(lg):
                        lab_k = lab_k[:-len(lg)] + ab
                ROT.append(np.full(k, lab_id(lab_k), np.int32)); SUB.append(np.zeros(k, np.int8))
                C1.append(np.full(k, ids[0], np.int32)); C2.append(np.full(k, ids[1], np.int32))
                C3.append(np.full(k, ids[2], np.int32))
                n_seg += k
        print(f"   {n_seg} segmentos")

    cat = np.concatenate
    vocab_arr = np.array(sorted(vocab, key=vocab.get), dtype=str)
    destino = os.path.join(V.APP_DIR, "catalogos", "red_vial.npz")
    np.savez_compressed(
        destino,
        x1=(cat(X1) - ORIGEN_X).astype(np.float32), y1=(cat(Y1) - ORIGEN_Y).astype(np.float32),
        x2=(cat(X2) - ORIGEN_X).astype(np.float32), y2=(cat(Y2) - ORIGEN_Y).astype(np.float32),
        clase=cat(CL), tmayor=cat(TM), rot=cat(ROT), sub=cat(SUB),
        etiquetas=np.array(sorted(labels, key=labels.get), dtype=str), c1=cat(C1), c2=cat(C2), c3=cat(C3), vocab=vocab_arr,
        origen=np.array([ORIGEN_X, ORIGEN_Y]))
    print(f"\nListo: {destino}  ({os.path.getsize(destino) / 1e6:.1f} MB, {len(cat(CL))} segmentos)")


if __name__ == "__main__":
    main()
