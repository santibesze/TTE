# ============================================================
# SELECCIÓN DE CONMUTADORES BAJO CARGA (CBC)
# ============================================================
#
# Flujo:
#   1. Lee entrada_01.json
#   2. Inyecta los datos en una COPIA de la plantilla
#      "Estrella gradual - Fina R1.xlsx" (Hoja1)
#   3. Recalcula las fórmulas (motor interno, sin LibreOffice)
#   4. Lee I_max, V_step, P_step, A, B, F calculados
#   5. Compara contra el catálogo "CBC (1).xlsx"
#   6. Genera PMA.xlsx con el modelo seleccionado
#
# Dependencias: pandas, openpyxl  (nada más)
# ============================================================

import copy
import json
import math
import re
import shutil
import sys
import unicodedata
import warnings
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

warnings.filterwarnings("ignore", module="openpyxl")


# ============================================================
# ARCHIVOS
# ============================================================

ARCHIVO_JSON = "entrada_02.json"
ARCHIVO_CBC = "CBC (1).xlsx"
ARCHIVO_ESTRELLA = "Estrella gradual - Fina R1.xlsx"
# La plantilla original NO se modifica: se trabaja sobre una copia.
ARCHIVO_ESTRELLA_CALC = "Estrella gradual - Fina R1 (calculado).xlsx"
ARCHIVO_PMA = "PMA.xlsx"

HOJA_ESTRELLA = "Hoja1"

# Mapeo JSON -> celda de Hoja1
MAPA_JSON_CELDAS = {
    "potencia_prim": "C4",
    "tension_prim": "C5",
    "prim_grupo": "C6",
    "frecuencia": "C7",
    "cbc_pasos_sup": "E8",
    "cbc_pasos_inf": "E9",
    "cbc_paso_porc": "E10",
    "bil_prim_fase": "E12",
    "bil_prim_neutro": "E13",
}

# Tensiones aplicadas (frecuencia industrial). Son OBLIGATORIAS en el JSON.
#
# CELDA_UWH1: la plantilla toma UWH1 de Hoja1!E14 ("Frecuencia Industrial
# (UWH1)") y Hoja2!F3 = Hoja1!E14 alimenta todas las fórmulas. L19 está
# vacía y nada depende de ella, por eso se usa E14. Si realmente querés
# otra celda, cambiá solo esta constante.
CELDA_UWH1 = "E14"
CELDA_UWH0 = "E15"

# nombre canónico -> (celda, claves aceptadas en el JSON)
MAPA_JSON_APLICADA = {
    "uwh_prim_fase": (CELDA_UWH1, ["ensayo_apli_prim_fase"]),
    "uwh_prim_neutro": (CELDA_UWH0, ["ensayo_apli_prim_neutro"]),
}


class ConexionNoSoportada(Exception):
    """La conexión pedida todavía no se puede calcular (Delta)."""


def validar_conexion_soportada(datos):
    """Por ahora solo se calcula conexión Estrella."""
    if normalizar_conexion(datos.get("prim_grupo")) == "D":
        raise ConexionNoSoportada("No puedo calcular para conexión Delta")


def buscar_clave_json(datos, alias):
    for k in alias:
        if k in datos:
            return k
    return None


# ============================================================
# UTILIDADES
# ============================================================

def convertir_numero(valor):
    """Devuelve float, o None si es vacío / '---' / no numérico."""
    if valor is None:
        return None

    if isinstance(valor, bool):
        return None

    if isinstance(valor, str):
        valor = valor.strip()
        if valor in ("", "---", "-", "--"):
            return None
        valor = valor.replace(",", ".")

    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None

    if math.isnan(numero):
        return None

    return numero


def fmt(valor):
    if valor is None:
        return "-"
    return f"{valor:.3f}"


def _sin_acentos(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c))


def normalizar_texto(texto):
    """
    Normalización flexible para comparar cabeceras:
    saltos de línea / espacios múltiples / mayúsculas / acentos.
    """
    texto = _sin_acentos(texto)
    texto = texto.replace("\r", " ").replace("\n", " ").replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip().lower()


def normalizar_conexion(grupo):
    """
    Conexión del JSON -> 'Y' o 'D'.   YN/Y/Estrella -> Y ; D/Δ/Triángulo -> D
    """
    if grupo is None:
        raise ValueError("La conexión no fue encontrada en el JSON.")

    g = _sin_acentos(str(grupo)).strip().upper().replace("Δ", "D")

    if g.startswith("Y") or g.startswith("ESTRELLA") or g.startswith("STAR"):
        return "Y"

    if g.startswith("D") or g.startswith("TRIANGULO") or g.startswith("DELTA"):
        return "D"

    raise ValueError(f"No se pudo interpretar la conexión: {grupo}")


def conexion_a_texto_excel(grupo):
    """Texto para la celda C6 de la plantilla."""
    return "Estrella" if normalizar_conexion(grupo) == "Y" else "Triángulo"


# ============================================================
# JSON
# ============================================================

def cargar_json(ruta_json=None):
    ruta = Path(ruta_json or ARCHIVO_JSON)

    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo JSON: {ruta}")

    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            return json.load(archivo)
    except json.JSONDecodeError as e:
        raise ValueError(f"El archivo JSON no es válido: {e}")


def mostrar_datos_json(datos):
    print("\nDATOS LEÍDOS DEL JSON")
    print("-" * 60)

    campos = [
        ("Potencia primaria [MVA]", "potencia_prim"),
        ("Tensión primaria [kV]", "tension_prim"),
        ("BIL primario fase [kV]", "bil_prim_fase"),
        ("BIL primario neutro [kV]", "bil_prim_neutro"),
        ("Grupo de conexión", "prim_grupo"),
        ("Frecuencia [Hz]", "frecuencia"),
        ("Cantidad de escalones superiores", "cbc_pasos_sup"),
        ("Cantidad de escalones inferiores", "cbc_pasos_inf"),
        ("Paso del CBC [%]", "cbc_paso_porc"),
    ]

    for nombre, clave in campos:
        print(f"{nombre}: {datos.get(clave, 'NO ENCONTRADO')}")

    for nombre, canon in (("Tensión aplicada fase UWH1 [kV]", "uwh_prim_fase"),
                          ("Tensión aplicada neutro UWH0 [kV]", "uwh_prim_neutro")):
        clave = buscar_clave_json(datos, MAPA_JSON_APLICADA[canon][1])
        print(f"{nombre}: {datos[clave] if clave else 'NO ENCONTRADO'}")


# ============================================================
# 1) INYECCIÓN DEL JSON EN LA PLANTILLA
# ============================================================

def inyectar_json_en_estrella(datos_json, ruta_excel):
    """
    Escribe cada variable del JSON en su celda de 'Hoja1'.

        potencia_prim   -> C4        tension_prim    -> C5
        prim_grupo      -> C6        frecuencia      -> C7
        cbc_pasos_sup   -> E8        cbc_pasos_inf   -> E9
        cbc_paso_porc   -> E10       bil_prim_fase   -> E12
        bil_prim_neutro -> E13

        ensayo_apli_prim_fase   -> E14 (UWH1)
        ensayo_apli_prim_neutro -> E15 (UWH0)

    Solo se tocan celdas de ENTRADA; las fórmulas quedan intactas.

    Guarda en la misma ruta indicada (pasar una copia de la plantilla).
    Nota: openpyxl guarda las fórmulas sin valores en caché; por eso
    después hay que llamar a recalcular_estrella().
    """
    ruta = Path(ruta_excel)
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró '{ruta_excel}'.")

    faltantes = [k for k in MAPA_JSON_CELDAS if k not in datos_json]
    if faltantes:
        raise ValueError(f"Faltan claves en el JSON: {', '.join(faltantes)}")

    wb = load_workbook(ruta)  # con fórmulas

    if HOJA_ESTRELLA not in wb.sheetnames:
        raise ValueError(f"No existe la hoja '{HOJA_ESTRELLA}' en {ruta.name}.")

    ws = wb[HOJA_ESTRELLA]

    for clave, celda in MAPA_JSON_CELDAS.items():
        valor = datos_json[clave]

        if clave == "prim_grupo":
            valor = conexion_a_texto_excel(valor)
        else:
            numero = convertir_numero(valor)
            if numero is None:
                raise ValueError(f"Valor inválido para '{clave}': {valor!r}")
            valor = int(numero) if numero.is_integer() else numero

        ws[celda].value = valor

    for canon, (celda, alias) in MAPA_JSON_APLICADA.items():
        clave = buscar_clave_json(datos_json, alias)
        numero = convertir_numero(datos_json[clave]) if clave else None
        if numero is None:
            raise ValueError(
                f"Falta (o es inválida) la tensión aplicada en el JSON "
                f"para {celda}. Claves aceptadas: {', '.join(alias)}"
            )
        ws[celda].value = int(numero) if numero.is_integer() else numero

    wb.save(ruta)
    return ruta


# ============================================================
# 2) RECÁLCULO (motor interno, sin LibreOffice)
# ============================================================
#
# Evalúa las fórmulas del libro leyéndolas con openpyxl y
# traduciéndolas a Python. Soporta lo que usa la plantilla:
# + - * / ^, paréntesis, SQRT, MIN, MAX, SUM, ABS, ROUND, IF,
# referencias 'Hoja'!A1 y rangos. Si una fórmula usa algo no
# soportado o da error (#REF!, /0), esa celda queda como error
# y solo falla si el programa la necesita.

class ErrorCelda(Exception):
    pass


class _Rango(list):
    pass


def _f_num(x):
    if isinstance(x, _Rango):
        raise ErrorCelda("rango usado como escalar")
    if x is None or x == "":
        return 0.0
    if isinstance(x, bool):
        return float(x)
    if isinstance(x, (int, float)):
        return float(x)
    n = convertir_numero(x)
    if n is None:
        raise ErrorCelda("#VALUE!")
    return n


def _aplanar(args):
    for a in args:
        if isinstance(a, _Rango):
            for v in a:
                yield v
        else:
            yield a


def _solo_numeros(args):
    return [float(v) for v in _aplanar(args)
            if isinstance(v, (int, float)) and not isinstance(v, bool)]


def _sqrt(x):
    x = _f_num(x)
    if x < 0:
        raise ErrorCelda("#NUM!")
    return math.sqrt(x)


_FUNCIONES = {
    "SQRT": _sqrt,
    "ABS": lambda x: abs(_f_num(x)),
    "MIN": lambda *a: min(_solo_numeros(a)) if _solo_numeros(a) else 0.0,
    "MAX": lambda *a: max(_solo_numeros(a)) if _solo_numeros(a) else 0.0,
    "SUM": lambda *a: sum(_solo_numeros(a)),
    "ROUND": lambda x, n=0: round(_f_num(x), int(_f_num(n))),
    "IF": lambda c, a=True, b=False: a if c else b,
    "PI": lambda: math.pi,
}

_RE_REF = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(?:(?:'(?P<h1>[^']+)'|(?P<h2>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)"
    r"(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?"
    r"(?![A-Za-z0-9_(])"
)


class _Evaluador:
    def __init__(self, wb):
        self.wb = wb
        self.cache = {}
        self.en_curso = set()

    def valor(self, hoja, celda):
        clave = (hoja, celda)

        if clave in self.cache:
            res = self.cache[clave]
            if isinstance(res, ErrorCelda):
                raise res
            return res

        if clave in self.en_curso:
            raise ErrorCelda("referencia circular")

        self.en_curso.add(clave)
        try:
            crudo = self.wb[hoja][celda].value
            if isinstance(crudo, str) and crudo.startswith("="):
                res = self._evaluar(hoja, crudo[1:])
            else:
                res = crudo
            self.cache[clave] = res
            return res
        except ErrorCelda as e:
            self.cache[clave] = e
            raise
        except ZeroDivisionError:
            e = ErrorCelda("#DIV/0!")
            self.cache[clave] = e
            raise e
        finally:
            self.en_curso.discard(clave)

    def _evaluar(self, hoja_actual, formula):
        if "#REF!" in formula.upper():
            raise ErrorCelda("#REF!")

        cuerpo = formula.replace("^", "**").replace("<>", "!=")
        cuerpo = re.sub(r"(?<![<>!=])=(?!=)", "==", cuerpo)

        def reemplazo(m):
            hoja = m.group("h1") or m.group("h2") or hoja_actual
            if hoja not in self.wb.sheetnames:
                raise ErrorCelda(f"hoja inexistente: {hoja}")
            c1, r1 = m.group("c1"), int(m.group("r1"))
            if m.group("c2"):
                c2, r2 = m.group("c2"), int(m.group("r2"))
                return f"_rng({hoja!r},{c1!r},{r1},{c2!r},{r2})"
            return f"_val({hoja!r},'{c1}{r1}')"

        cuerpo = _RE_REF.sub(reemplazo, cuerpo)

        def _val(hoja, celda):
            v = self.valor(hoja, celda)
            return 0.0 if v is None else v

        def _rng(hoja, c1, r1, c2, r2):
            out = _Rango()
            for r in range(min(r1, r2), max(r1, r2) + 1):
                for c in range(column_index_from_string(min(c1, c2, key=column_index_from_string)),
                               column_index_from_string(max(c1, c2, key=column_index_from_string)) + 1):
                    out.append(self.valor(hoja, f"{get_column_letter(c)}{r}"))
            return out

        entorno = {"__builtins__": {}, "_val": _val, "_rng": _rng}
        entorno.update({k.lower(): v for k, v in _FUNCIONES.items()})
        entorno.update(_FUNCIONES)

        try:
            return eval(cuerpo, entorno)  # noqa: S307 (entorno restringido)
        except NameError as e:
            raise ErrorCelda(f"función no soportada: {e}")
        except (SyntaxError, TypeError) as e:
            raise ErrorCelda(f"fórmula no soportada: {e}")


def recalcular_estrella(ruta_excel, celdas_hoja1):
    """
    Devuelve {celda: valor} de 'Hoja1' con las fórmulas evaluadas
    a partir de los valores de entrada actuales del archivo.
    Las celdas con error quedan como ErrorCelda (no se ocultan).
    """
    wb = load_workbook(ruta_excel)  # fórmulas
    ev = _Evaluador(wb)

    resultado = {}
    for celda in celdas_hoja1:
        try:
            resultado[celda] = ev.valor(HOJA_ESTRELLA, celda)
        except ErrorCelda as e:
            resultado[celda] = e
    return resultado


# ============================================================
# 3) LECTURA DE RESULTADOS DE ESTRELLA GRADUAL
# ============================================================

CELDAS_RESULTADO = [
    "I11", "I12", "I13",          # I_max, V_step, P_step
    "H28", "H29",                 # A  (50 Hz / impulso)
    "K28", "K29",                 # A1 (50 Hz / impulso)
    "J24", "J25",                 # B  (50 Hz / impulso)
    "B28", "C28", "B29", "C29",   # F  (fila 28: 50 Hz / fila 29: impulso)
]


def _exigir(resultados, celda):
    v = resultados[celda]
    if isinstance(v, ErrorCelda):
        raise ValueError(f"La celda {HOJA_ESTRELLA}!{celda} dio error: {v}")
    n = convertir_numero(v)
    if n is None:
        raise ValueError(f"La celda {HOJA_ESTRELLA}!{celda} no es numérica: {v!r}")
    return n


def leer_estrella(datos_json, ruta_original=ARCHIVO_ESTRELLA,
                  ruta_copia=ARCHIVO_ESTRELLA_CALC):
    """
    1. Copia la plantilla (la original queda intacta).
    2. Inyecta el JSON en la copia.
    3. Recalcula y lee los resultados YA actualizados.

    Devuelve el dict de parámetros usado por seleccionar_cbc().
    """
    validar_conexion_soportada(datos_json)

    origen = Path(ruta_original)
    if not origen.exists():
        raise FileNotFoundError(f"No se encontró el archivo '{ruta_original}'.")

    shutil.copyfile(origen, ruta_copia)
    inyectar_json_en_estrella(datos_json, ruta_copia)

    r = recalcular_estrella(ruta_copia, CELDAS_RESULTADO)

    a_50, a_imp = _exigir(r, "H28"), _exigir(r, "H29")
    a1_50, a1_imp = _exigir(r, "K28"), _exigir(r, "K29")
    b_50, b_imp = _exigir(r, "J24"), _exigir(r, "J25")
    f28 = [_exigir(r, "B28"), _exigir(r, "C28")]
    f29 = [_exigir(r, "B29"), _exigir(r, "C29")]

    return {
        "I_max": _exigir(r, "I11"),
        "V_step": _exigir(r, "I12"),
        "P_step": _exigir(r, "I13"),

        # A / B / F separados por tipo de ensayo
        "a_50hz": a_50, "a_imp": a_imp,       # H28 / H29
        "a1_50hz": a1_50, "a1_imp": a1_imp,   # K28 / K29
        "b_50hz": b_50, "b_imp": b_imp,       # J24 / J25
        "f_max_28": max(f28),                 # fila 28 -> 50 Hz
        "f_max_29": max(f29),                 # fila 29 -> impulso

        # valores individuales de F (para mostrar)
        "f1": f28[0], "f2": f28[1], "f3": f29[0], "f4": f29[1],

        "fuente": ruta_copia,
        "fila_resultados_excel": 29,
    }


# ============================================================
# 4) CATÁLOGO CBC
# ============================================================

# nombre lógico -> lista de cabeceras posibles (ya normalizadas)
ALIAS_COLUMNAS = {
    "item": ["item"],
    "marca": ["marca"],
    "modelo": ["modelo"],
    "fases": ["fases"],
    "corriente": ["corriente nominal [a]", "corriente nominal"],
    "conexion": ["conexion"],
    "vstep1": ["maxima tension por escalon 1. [v]", "maxima tension por escalon 1"],
    "vstep2": ["maxima tension por escalon 2. [v]", "maxima tension por escalon 2"],
    "capacidad": ["capacidad de contactos. [kva]", "capacidad de contactos"],
    "posiciones": ["maximo n° de posiciones de operacion.",
                   "maximo n° de posiciones de operacion",
                   "maximo n de posiciones de operacion"],
}

# familias de columnas eléctricas: prefijo + tipo de ensayo
FAMILIAS_ENSAYO = ["a", "a1", "b", "b1", "b2", "b3",
                   "c1", "c2", "c2_1", "c2_2", "f", "f_1", "f_2"]
ENSAYOS = {"imp": "1,2/50us [kv]", "50hz": "50hz 1min [kv]"}


def _clave_ensayo(texto_normalizado):
    """
    'a 1,2/50us [kv]' -> ('a', 'imp');  'f_2 50hz 1min [kv]' -> ('f_2', '50hz')
    Tolera espacios de más y ausencia de espacio tras el prefijo.
    """
    t = texto_normalizado.replace(" ", "")
    for fam in sorted(FAMILIAS_ENSAYO, key=len, reverse=True):
        for clave, sufijo in ENSAYOS.items():
            if t == fam + sufijo.replace(" ", ""):
                return fam, clave
    return None


def cargar_catalogo():
    ruta = Path(ARCHIVO_CBC)

    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo CBC: {ARCHIVO_CBC}")

    try:
        tabla = pd.read_excel(ruta, header=1)
    except Exception as e:
        raise RuntimeError(f"No se pudo abrir {ARCHIVO_CBC}: {e}")

    # Cabeceras: saltos de línea, espacios extra, mayúsculas
    tabla.columns = [
        re.sub(r"\s+", " ", str(c).replace("\n", " ")).strip()
        for c in tabla.columns
    ]

    # Se descartan filas sin datos y las notas al pie (sin corriente / conexión)
    tabla = tabla[tabla.iloc[:, 0].notna() & tabla.iloc[:, 4].notna()]
    return tabla.reset_index(drop=True)


def mapear_columnas(tabla):
    """
    Devuelve dict con:
      'basicas': {nombre_logico: nombre_real_de_columna}
      'ensayo' : {(familia, 'imp'|'50hz'): nombre_real_de_columna}
    Lanza error si faltan las columnas imprescindibles.
    """
    por_norm = {normalizar_texto(c): c for c in tabla.columns}

    basicas, faltantes = {}, []
    for logico, alias in ALIAS_COLUMNAS.items():
        real = next((por_norm[a] for a in alias if a in por_norm), None)
        if real is None:
            faltantes.append(logico)
        else:
            basicas[logico] = real

    ensayo = {}
    for norm, real in por_norm.items():
        clave = _clave_ensayo(norm)
        if clave:
            ensayo[clave] = real

    for req in [("a", "imp"), ("a", "50hz"), ("a1", "imp"), ("a1", "50hz"),
                ("b", "imp"), ("f", "imp")]:
        if req not in ensayo:
            faltantes.append(f"{req[0]} ({req[1]})")

    if faltantes:
        raise ValueError(
            "Faltan columnas necesarias en el Excel CBC:\n"
            + "\n".join(f"  - {c}" for c in faltantes)
            + "\nColumnas encontradas:\n"
            + "\n".join(f"  · {c!r}" for c in tabla.columns)
        )

    return {"basicas": basicas, "ensayo": ensayo}


def normalizar_conexion_catalogo(valor):
    v = normalizar_texto(valor)
    if v in ("y", "yn", "estrella"):
        return "Y"
    if v in ("d", "delta", "triangulo"):
        return "D"
    if v.startswith("mono"):
        return "MONO"
    return v.upper()


# ============================================================
# 5) SELECCIÓN
# ============================================================

def _cumple(fila, columna, requerido):
    """Vacío / '---' = no aplica -> cumple."""
    v = convertir_numero(fila[columna])
    return True if v is None else v >= requerido


def seleccionar_cbc(tabla, datos, parametros, columnas=None):
    """
    Compara los valores calculados en Estrella gradual contra el catálogo.

    Requeridos (Estrella)          Catálogo
    ---------------------          --------------------------------
    I_max (I11)                 -> Corriente nominal [A]
    P_step (I13)                -> Capacidad de contactos [kVA]
    V_step (I12)                -> Máx. tensión por escalón 1 y 2 [V]
    A: H28 (50 Hz) / H29 (imp.) -> a  50Hz 1min / a  1,2/50us
    A1: K28 (50 Hz) / K29 (imp.)-> a1 50Hz 1min / a1 1,2/50us
    B: J24 (50 Hz) / J25 (imp.) -> b  (Y)  o  b1,b2,b3 (D)
    F: máx. fila 28 (50 Hz)     -> f, f_1, f_2  50Hz 1min
       máx. fila 29 (impulso)   -> f, f_1, f_2  1,2/50us
    """
    if columnas is None:
        columnas = mapear_columnas(tabla)

    bas, ens = columnas["basicas"], columnas["ensayo"]
    conexion = normalizar_conexion(datos["prim_grupo"])

    posiciones = (
        convertir_numero(datos["cbc_pasos_inf"])
        + convertir_numero(datos["cbc_pasos_sup"])
        + 1
    )

    # familia -> (requerido 50 Hz, requerido impulso)
    familias_b = ["b"] if conexion == "Y" else ["b1", "b2", "b3"]
    requisitos = [("a", parametros["a_50hz"], parametros["a_imp"]),
                  ("a1", parametros["a1_50hz"], parametros["a1_imp"])]
    requisitos += [(f, parametros["b_50hz"], parametros["b_imp"]) for f in familias_b]
    requisitos += [(f, parametros["f_max_28"], parametros["f_max_29"])
                   for f in ("f", "f_1", "f_2")]

    candidatos = []

    for indice, fila in tabla.iterrows():

        if normalizar_conexion_catalogo(fila[bas["conexion"]]) != conexion:
            continue

        corriente = convertir_numero(fila[bas["corriente"]])
        if corriente is None or corriente < parametros["I_max"]:
            continue

        potencia = convertir_numero(fila[bas["capacidad"]])
        if potencia is None or potencia < parametros["P_step"]:
            continue

        max_pos = convertir_numero(fila[bas["posiciones"]])
        if max_pos is not None and max_pos < posiciones:
            continue

        if not _cumple(fila, bas["vstep1"], parametros["V_step"]):
            continue
        if not _cumple(fila, bas["vstep2"], parametros["V_step"]):
            continue

        ok = True
        for familia, req_50, req_imp in requisitos:
            for tipo, req in (("50hz", req_50), ("imp", req_imp)):
                col = ens.get((familia, tipo))
                if col is not None and not _cumple(fila, col, req):
                    ok = False
                    break
            if not ok:
                break
        if not ok:
            continue

        candidato = fila.copy()
        # header en fila 2 del Excel + índice 0-based => fila = índice + 3
        candidato["_fila_excel_cbc"] = indice + 3
        candidatos.append(candidato)

    candidatos.sort(key=lambda f: (
        convertir_numero(f[bas["corriente"]]),
        convertir_numero(f[bas["capacidad"]]) or 0,
    ))
    return candidatos


def diagnosticar_rechazos(tabla, datos, parametros, columnas):
    """Explica por qué cada modelo de la conexión pedida fue descartado."""
    bas, ens = columnas["basicas"], columnas["ensayo"]
    conexion = normalizar_conexion(datos["prim_grupo"])
    familias_b = ["b"] if conexion == "Y" else ["b1", "b2", "b3"]

    checks = [("I_max", bas["corriente"], parametros["I_max"]),
              ("P_step", bas["capacidad"], parametros["P_step"]),
              ("V_step esc.1", bas["vstep1"], parametros["V_step"]),
              ("V_step esc.2", bas["vstep2"], parametros["V_step"])]
    for fam, r50, rimp in ([("a", parametros["a_50hz"], parametros["a_imp"]),
                            ("a1", parametros["a1_50hz"], parametros["a1_imp"])]
                           + [(f, parametros["b_50hz"], parametros["b_imp"]) for f in familias_b]
                           + [(f, parametros["f_max_28"], parametros["f_max_29"])
                              for f in ("f", "f_1", "f_2")]):
        if (fam, "50hz") in ens:
            checks.append((f"{fam} 50Hz", ens[(fam, "50hz")], r50))
        if (fam, "imp") in ens:
            checks.append((f"{fam} 1,2/50", ens[(fam, "imp")], rimp))

    lineas = []
    for _, fila in tabla.iterrows():
        if normalizar_conexion_catalogo(fila[bas["conexion"]]) != conexion:
            continue
        motivos = []
        for nombre, col, req in checks:
            v = convertir_numero(fila[col])
            if v is not None and v < req:
                motivos.append(f"{nombre}: {v:g} < {req:.2f}")
        lineas.append(f"  ITEM {fila[bas['item']]:g} ({fila[bas['corriente']]:g} A): "
                      + ("; ".join(motivos) if motivos else "sin rechazo por valores"))
    return lineas


# ============================================================
# TIPO DE EQUIPO / PMA
# ============================================================

def obtener_tipo_equipo_desde_excel(fila, columnas):
    valor = normalizar_texto(fila[columnas["basicas"]["fases"]]).upper()

    if valor in ("I", "1", "MONOFASICO"):
        return "Monofásico"
    if valor in ("III", "3", "TRIFASICO"):
        return "Trifásico"
    return f"No identificado ({valor})"


def generar_pma(seleccionado, conexion, tension, columnas):
    wb = Workbook()
    ws = wb.active
    ws.title = "ACCESORIOS"

    ws["A1"] = "Modelo seleccionado"
    ws["B1"] = "Conexión"
    ws["C1"] = "Tensión [kV]"

    ws["A2"] = seleccionado[columnas["basicas"]["modelo"]]
    ws["B2"] = conexion
    ws["C2"] = tension

    wb.save(ARCHIVO_PMA)


# ============================================================
# MAIN
# ============================================================

def main(ruta_json=None):
    try:
        # 1. JSON  (uso: python main_cbc.py entrada_00.json)
        datos = cargar_json(ruta_json)
        mostrar_datos_json(datos)

        # Por ahora solo se calcula conexión Estrella
        validar_conexion_soportada(datos)

        # 2. Inyección + recálculo + lectura de resultados
        print("\nINYECTANDO JSON Y RECALCULANDO ESTRELLA GRADUAL FINA")
        parametros = leer_estrella(datos)

        print("\nVARIABLES CALCULADAS EN ESTRELLA GRADUAL FINA")
        print("-" * 60)
        print(f"I_max (I11)          : {fmt(parametros['I_max'])}")
        print(f"V_step (I12)         : {fmt(parametros['V_step'])}")
        print(f"P_step (I13)         : {fmt(parametros['P_step'])}")
        print(f"A  H28 (50 Hz)       : {fmt(parametros['a_50hz'])}")
        print(f"A  H29 (impulso)     : {fmt(parametros['a_imp'])}")
        print(f"A1 K28 (50 Hz)       : {fmt(parametros['a1_50hz'])}")
        print(f"A1 K29 (impulso)     : {fmt(parametros['a1_imp'])}")
        print(f"B  J24 (50 Hz)       : {fmt(parametros['b_50hz'])}")
        print(f"B  J25 (impulso)     : {fmt(parametros['b_imp'])}")
        print(f"F  B28={fmt(parametros['f1'])}  C28={fmt(parametros['f2'])}"
              f"  -> máx fila 28 (50 Hz)  : {fmt(parametros['f_max_28'])}")
        print(f"F  B29={fmt(parametros['f3'])}  C29={fmt(parametros['f4'])}"
              f"  -> máx fila 29 (impulso): {fmt(parametros['f_max_29'])}")
        print(f"Libro calculado: {parametros['fuente']}")

        # 3. Catálogo
        tabla = cargar_catalogo()
        columnas = mapear_columnas(tabla)

        # 4. Selección
        candidatos = seleccionar_cbc(tabla, datos, parametros, columnas)

        if not candidatos:
            print("\nNo se ha podido seleccionar un CBC.")
            print("No hay un modelo disponible para esta configuración.")
            print("\nMotivos de descarte (conexión solicitada):")
            for linea in diagnosticar_rechazos(tabla, datos, parametros, columnas):
                print(linea)
            return

        seleccionado = candidatos[0]
        conexion = normalizar_conexion(datos["prim_grupo"])
        bas = columnas["basicas"]

        print("\nCBC SELECCIONADO")
        print("-" * 60)
        print(f"Modelo: {seleccionado[bas['modelo']]}")
        print(f"Marca: {seleccionado[bas['marca']]}")
        print(f"Fases: {seleccionado[bas['fases']]}")
        print(f"Tipo de equipo: {obtener_tipo_equipo_desde_excel(seleccionado, columnas)}")
        print(f"Conexión: {seleccionado[bas['conexion']]}")
        print(f"Corriente nominal: {seleccionado[bas['corriente']]} A")
        print(f"Fila del Excel CBC utilizada: {seleccionado['_fila_excel_cbc']}")
        print(f"Candidatos que cumplen: {len(candidatos)}")

        generar_pma(seleccionado, conexion, datos["tension_prim"], columnas)
        print("\nArchivo PMA.xlsx generado correctamente.")

    except ConexionNoSoportada as e:
        print(f"\n{e}")

    except Exception as e:
        print("\nERROR:")
        print(str(e))


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
    main()