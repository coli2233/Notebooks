# -*- coding: utf-8 -*-
"""
================================================================================
 TECNOLOGIA DEL CONCRETO - PROGRAMA PRINCIPAL
================================================================================

Programa de consola que integra las dos herramientas del curso en un solo
menu:

  1) Diseno de mezclas de concreto por el metodo del Comite ACI 211
     (usa las TABLAS ACI 01 a 09 definidas como constantes al inicio).
  2) Calculadora de resistencia del concreto (modulo resistencia_concreto.py),
     que se invoca desde este menu principal.

Salidas
-------
  * Resumen de la dosificacion en pantalla.
  * Informe en Excel (.xlsx) con varias hojas (datos de entrada, dosificacion,
    paso a paso y tablas ACI), generado solo si el usuario lo solicita.

Librerias
---------
  pandas, numpy, openpyxl

Unidades de trabajo
-------------------
  Resistencias  : kg/cm2
  Pesos         : kg/m3 (dosificacion) y kg (cantidades)
  Volumenes     : m3 y litros (lt)
================================================================================
"""

import os
import re
import sys
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Modulo del INTEGRANTE 5 (Luis Angel Espinoza Farfan):
# calculadora de resistencia del concreto. Su desarrollo completo esta en el
# archivo resistencia_concreto.py y se invoca desde el menu principal.
# -----------------------------------------------------------------------------
import resistencia_concreto

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass


# =============================================================================
#  INTEGRANTE 1 : Xiomara Brissely Nuñez Iman
#  PARTE        : DATOS NORMATIVOS (constantes generales + Tablas ACI 01-09)
#  QUE DESARROLLA:
#    - Las constantes generales del programa (peso del saco de cemento, peso
#      especifico del agua, ancho del reporte, orden de los tamanos, etc.).
#    - Las TABLAS ACI 01 a 09 definidas como constantes: volumen de agua de
#      mezclado, aire atrapado, peso del agregado grueso (b/bo), relacion
#      agua/cemento, aire incorporado, condiciones de exposicion, porcentaje
#      de agregado fino y volumen de agua segun la forma del agregado.
#    - Los valores del ejemplo resuelto "corregido", usados como valores por
#      defecto en todas las preguntas del programa.
#  IMPORTANCIA : es la base de datos normativa; el motor de calculo (integrante
#    4) toma de aqui todos los valores de las tablas para dosificar la mezcla.
# =============================================================================
# =============================================================================
# 1) CONSTANTES GENERALES
# =============================================================================
PESO_SACO_CEMENTO = 42.5          # kg por saco de cemento (usado en el ejemplo)
PESO_ESPECIFICO_AGUA = 1.00       # kg/lt
RESPONSABLE_DEFECTO = "ING. FLAVIO CESAR TRUJILLO BAILON"
ANCHO = 78                        # ancho del reporte en consola

# Lista de tamanos maximos nominales usados en el diseno
TMN_DISENO = ['3/8"', '1/2"', '3/4"', '1"', '1 1/2"', '2"', '3"', '6"']
# Orden de presentacion de la Tabla 02 (aire atrapado)
ORDEN_AIRE_ATRAPADO = ['3/8"', '1/2"', '3/4"', '1"', '1 1/2"', '2"', '3"', '4"']


# =============================================================================
# 2) TABLAS ACI (tomadas del documento "yablas aci")
# =============================================================================

# --- TABLA 01: VOLUMEN UNITARIO DE AGUA (lt/m3) -------------------------------
TABLA_01_AGUA = {
    "sin_aire": {
        '1" a 2"': {'3/8"': 207, '1/2"': 199, '3/4"': 190, '1"': 179,
                    '1 1/2"': 166, '2"': 154, '3"': 130, '6"': 113},
        '3" a 4"': {'3/8"': 228, '1/2"': 216, '3/4"': 205, '1"': 193,
                    '1 1/2"': 181, '2"': 169, '3"': 145, '6"': 124},
        '6" a 7"': {'3/8"': 243, '1/2"': 228, '3/4"': 216, '1"': 202,
                    '1 1/2"': 190, '2"': 178, '3"': 160, '6"': None},
    },
    "con_aire": {
        '1" a 2"': {'3/8"': 181, '1/2"': 175, '3/4"': 168, '1"': 160,
                    '1 1/2"': 150, '2"': 142, '3"': 122, '6"': 107},
        '3" a 4"': {'3/8"': 202, '1/2"': 193, '3/4"': 184, '1"': 175,
                    '1 1/2"': 165, '2"': 157, '3"': 133, '6"': 119},
        '6" a 7"': {'3/8"': 216, '1/2"': 205, '3/4"': 197, '1"': 184,
                    '1 1/2"': 174, '2"': 166, '3"': 154, '6"': None},
    },
}

# --- TABLA 02: CONTENIDO DE AIRE ATRAPADO (%) --------------------------------
TABLA_02_AIRE_ATRAPADO = {
    '3/8"': 3.0, '1/2"': 2.5, '3/4"': 2.0, '1"': 1.5,
    '1 1/2"': 1.0, '2"': 0.5, '3"': 0.3, '4"': 0.2, '6"': 0.2,
}

# --- TABLA 03: MODULO DE FINEZA DE LA COMBINACION DE AGREGADOS ---------------
TABLA_03_MF_COMBINACION = {
    '3/8"':   {6: 3.96, 7: 4.04, 8: 4.11, 9: 4.19},
    '1/2"':   {6: 4.46, 7: 4.54, 8: 4.61, 9: 4.69},
    '3/4"':   {6: 4.96, 7: 5.04, 8: 5.11, 9: 5.19},
    '1"':     {6: 5.26, 7: 5.34, 8: 5.41, 9: 5.49},
    '1 1/2"': {6: 5.56, 7: 5.64, 8: 5.71, 9: 5.79},
    '2"':     {6: 5.86, 7: 5.94, 8: 6.01, 9: 6.09},
    '3"':     {6: 6.16, 7: 6.24, 8: 6.31, 9: 6.39},
}

# --- TABLA 04: PESO DEL AGREGADO GRUESO POR UNIDAD DE VOLUMEN (b/bo) ---------
TABLA_04_AG_GRUESO = {
    '3/8"':   {2.40: 0.50, 2.60: 0.48, 2.80: 0.46, 3.00: 0.44},
    '1/2"':   {2.40: 0.59, 2.60: 0.57, 2.80: 0.55, 3.00: 0.53},
    '3/4"':   {2.40: 0.66, 2.60: 0.64, 2.80: 0.62, 3.00: 0.60},
    '1"':     {2.40: 0.71, 2.60: 0.69, 2.80: 0.67, 3.00: 0.65},
    '1 1/2"': {2.40: 0.76, 2.60: 0.74, 2.80: 0.72, 3.00: 0.70},
    '2"':     {2.40: 0.78, 2.60: 0.76, 2.80: 0.74, 3.00: 0.72},
    '3"':     {2.40: 0.81, 2.60: 0.79, 2.80: 0.77, 3.00: 0.75},
    '6"':     {2.40: 0.87, 2.60: 0.85, 2.80: 0.83, 3.00: 0.81},
}

# --- TABLA 05: RELACION AGUA/CEMENTO POR RESISTENCIA -------------------------
#     f'c (kg/cm2) -> (sin aire incorporado, con aire incorporado)
TABLA_05_AC = {
    150: (0.80, 0.71),
    200: (0.70, 0.61),
    250: (0.62, 0.53),
    300: (0.55, 0.46),
    350: (0.48, 0.40),
    400: (0.43, None),
    450: (0.38, None),
}

# --- TABLA 06: CONTENIDO DE AIRE INCORPORADO Y TOTAL (%) ---------------------
#     TMN -> (exposicion suave, moderada, severa)
TABLA_06_AIRE_INCORPORADO = {
    '3/8"':   (4.5, 6.0, 7.5),
    '1/2"':   (4.0, 5.5, 7.0),
    '3/4"':   (3.5, 5.0, 6.5),
    '1"':     (3.0, 4.5, 6.0),
    '1 1/2"': (2.5, 4.5, 5.5),
    '2"':     (2.0, 4.0, 5.0),
    '3"':     (1.5, 3.5, 4.5),
    '6"':     (1.0, 3.0, 4.0),
}

# --- TABLA 07: CONDICIONES ESPECIALES DE EXPOSICION --------------------------
TABLA_07_EXPOSICION = [
    {"condicion": "Baja permeabilidad - expuesto a agua dulce",
     "w_c_max": 0.50, "f_c_min_liviano": 2.60},
    {"condicion": "Baja permeabilidad - agua de mar o aguas solubles",
     "w_c_max": 0.45, "f_c_min_liviano": None},
    {"condicion": "Baja permeabilidad - aguas cloacales",
     "w_c_max": 0.45, "f_c_min_liviano": None},
    {"condicion": "Congelacion y deshielo - bardineles, cunetas, secc. delgadas",
     "w_c_max": 0.45, "f_c_min_liviano": 300},
    {"condicion": "Congelacion y deshielo - otros elementos",
     "w_c_max": 0.50, "f_c_min_liviano": None},
    {"condicion": "Proteccion contra corrosion (agua de mar / salubres)",
     "w_c_max": 0.40, "f_c_min_liviano": 325},
    {"condicion": "Proteccion contra corrosion - recubrimiento +15 mm",
     "w_c_max": 0.45, "f_c_min_liviano": 300},
]
TABLA_07_NOTA = "La resistencia f'c no debera ser menor de 245 Kg/cm2 por razones de durabilidad"

# --- TABLA 08: PORCENTAJE DE AGREGADO FINO -----------------------------------
#     estructura: MF_parte -> forma -> TMN -> {factor sacos: %}
TABLA_08_PORCENTAJE_FINO = {
    "2.3-2.4": {
        "redondeado": {
            '3/8"': {5: 60, 6: 57, 7: 54, 8: 51}, '1/2"': {5: 49, 6: 46, 7: 43, 8: 40},
            '3/4"': {5: 41, 6: 38, 7: 35, 8: 33}, '1"': {5: 40, 6: 37, 7: 34, 8: 32},
            '1 1/2"': {5: 37, 6: 34, 7: 32, 8: 30}, '2"': {5: 36, 6: 33, 7: 31, 8: 29},
        },
        "angular": {
            '3/8"': {5: 69, 6: 65, 7: 61, 8: 58}, '1/2"': {5: 57, 6: 54, 7: 51, 8: 48},
            '3/4"': {5: 48, 6: 45, 7: 43, 8: 41}, '1"': {5: 47, 6: 44, 7: 42, 8: 40},
            '1 1/2"': {5: 44, 6: 41, 7: 39, 8: 37}, '2"': {5: 43, 6: 40, 7: 38, 8: 36},
        },
    },
    "2.6-2.7": {
        "redondeado": {
            '3/8"': {5: 66, 6: 62, 7: 59, 8: 56}, '1/2"': {5: 53, 6: 50, 7: 47, 8: 44},
            '3/4"': {5: 44, 6: 41, 7: 38, 8: 36}, '1"': {5: 42, 6: 39, 7: 37, 8: 35},
            '1 1/2"': {5: 40, 6: 37, 7: 35, 8: 33}, '2"': {5: 37, 6: 35, 7: 33, 8: 32},
        },
        "angular": {
            '3/8"': {5: 75, 6: 71, 7: 67, 8: 64}, '1/2"': {5: 61, 6: 58, 7: 55, 8: 53},
            '3/4"': {5: 51, 6: 48, 7: 46, 8: 44}, '1"': {5: 49, 6: 46, 7: 44, 8: 42},
            '1 1/2"': {5: 47, 6: 44, 7: 42, 8: 40}, '2"': {5: 45, 6: 42, 7: 40, 8: 38},
        },
    },
    "3.0-3.1": {
        "redondeado": {
            '3/8"': {5: 74, 6: 70, 7: 66, 8: 62}, '1/2"': {5: 59, 6: 56, 7: 53, 8: 50},
            '3/4"': {5: 49, 6: 46, 7: 43, 8: 40}, '1"': {5: 47, 6: 44, 7: 41, 8: 38},
            '1 1/2"': {5: 44, 6: 41, 7: 38, 8: 36}, '2"': {5: 42, 6: 38, 7: 36, 8: 34},
        },
        "angular": {
            '3/8"': {5: 84, 6: 80, 7: 76, 8: 73}, '1/2"': {5: 70, 6: 66, 7: 62, 8: 59},
            '3/4"': {5: 57, 6: 54, 7: 51, 8: 48}, '1"': {5: 55, 6: 52, 7: 49, 8: 46},
            '1 1/2"': {5: 52, 6: 49, 7: 46, 8: 44}, '2"': {5: 49, 6: 46, 7: 44, 8: 42},
        },
    },
}

# --- TABLA 09: VOLUMEN UNITARIO DE AGUA POR FORMA (lt/m3, sin aire) ----------
#     TMN -> (1-2 red, 1-2 ang, 3-4 red, 3-4 ang, 6-7 red, 6-7 ang)
TABLA_09_AGUA_POR_FORMA = {
    '3/8"':   (185, 212, 201, 227, 230, 250),
    '1/2"':   (182, 201, 197, 216, 219, 238),
    '3/4"':   (170, 189, 185, 204, 208, 227),
    '1"':     (163, 182, 178, 197, 197, 216),
    '1 1/2"': (155, 170, 170, 185, 185, 204),
    '2"':     (148, 163, 163, 178, 178, 197),
    '3"':     (136, 151, 151, 167, 163, 182),
}

# --- RESISTENCIA PROMEDIO f'cr (PASO 1) --------------------------------------
#     regla: (limite inferior, limite superior, adicion)
RESISTENCIA_PROMEDIO_REGLA = [
    (0.0, 209.999, 70.0, "F'c + 70"),
    (210.0, 350.0, 84.0, "F'c + 84"),
    (350.001, 99999.0, 98.0, "F'c + 98"),
]

# --- CONSISTENCIA / ASENTAMIENTO (PASO 2) ------------------------------------
TABLA_CONSISTENCIA = [
    ("Seca",     '0" a 2"',     '1" a 2"',  "0\" a 2\""),
    ("Plastica", '3" a 4"',     '3" a 4"',  "3\" a 4\""),
    ("Fluida",   '>= 5"',       '6" a 7"',  "6\" a 7\""),
]


# =============================================================================
# 2.b) VALORES POR DEFECTO = DATOS DEL EJEMPLO RESUELTO "CORREGIDO"
#      (Documento: CORREGIDO_diseno_mezclas_ACI211_completo.pdf)
#      Son los valores que el programa propone al presionar ENTER.
# =============================================================================
EJEMPLO_CORREGIDO = {
    "descripcion": "losa de concreto",
    "fc": 210,
    "responsable": RESPONSABLE_DEFECTO,
    "consistencia": "Plastica",
    "con_aire": False,
    "exposicion_idx": 0,
    "cemento_nombre": "Andino Tipo V",
    "pe_cemento": 3.12,
    # Agregado fino (arena)
    "arena_puss": 1580,
    "arena_pucs": 1710,
    "arena_pe": 2.65,
    "arena_humedad": 3.5,
    "arena_absorcion": 1.2,
    "arena_mf": 2.7,
    # Agregado grueso (piedra)
    "piedra_puss": 1550,
    "piedra_pucs": 1660,
    "piedra_pe": 2.62,
    "piedra_humedad": 0.4,
    "piedra_absorcion": 1.0,
    "tmn": '3/4"',
    # Dosificacion para probetas
    "probeta_d": 15,
    "probeta_h": 30,
    "n_probetas": 3,
    "desperdicio": 20,
}


# --------------------- fin de Xiomara Brissely Nuñez Iman ---------------------


# =============================================================================
#  INTEGRANTE 2 : Aymi Juzarit Escalante Ramos
#  PARTE        : UTILIDADES DE INTERFAZ Y DE CALCULO
#  QUE DESARROLLA:
#    - Utilidades de consola: limpiar pantalla, titulos/separadores, pausa,
#      lectura y validacion de numeros y textos, menus numerados y preguntas
#      si/no con valor por defecto.
#    - Utilidades numericas: interpolacion lineal, ubicacion del rango entre
#      claves de una tabla, formateo de numeros y redondeo "mitad hacia
#      arriba" (ROUND_HALF_UP) que uniformiza los resultados del diseno.
#  IMPORTANCIA : da soporte comun a la captura de datos (integrante 3) y al
#    motor de calculo (integrante 4), y define con que criterio se redondea.
# =============================================================================
# =============================================================================
# 3) UTILIDADES DE CONSOLA
# =============================================================================
def limpiar_pantalla():
    os.system("cls" if os.name == "nt" else "clear")


def linea(caracter="=", ancho=ANCHO):
    print(caracter * ancho)


def encabezado(texto):
    print()
    linea("=")
    print(texto.center(ANCHO))
    linea("=")


def pausa(mensaje="\nPresione ENTER para continuar..."):
    input(mensaje)


def pedir_texto(mensaje, defecto=""):
    valor = input(mensaje).strip()
    return valor if valor else defecto


def pedir_flotante(mensaje, defecto=None, minimo=None, maximo=None):
    while True:
        entrada = input(mensaje).strip().replace(",", ".")
        if not entrada and defecto is not None:
            return round(float(defecto), 3)
        if not entrada:
            print("   [!] Debe ingresar un valor.")
            continue
        try:
            valor = float(entrada)
        except ValueError:
            print("   [!] Valor no valido.")
            continue
        if minimo is not None and valor < minimo:
            print(f"   [!] Debe ser >= {minimo}.")
            continue
        if maximo is not None and valor > maximo:
            print(f"   [!] Debe ser <= {maximo}.")
            continue
        return round(valor, 3)   # se guarda con 3 decimales como maximo


def pedir_opcion(mensaje, opciones, defecto):
    """
    opciones: dict {'1': ('etiqueta', valor)}.
    Muestra las opciones numeradas ('N  valor') y devuelve el valor elegido.
    Si se presiona ENTER, devuelve la opcion por defecto.
    """
    for clave, (etiqueta, _) in opciones.items():
        marca = "   <- por defecto" if clave == str(defecto) else ""
        print(f"     {clave}  {etiqueta}{marca}")
    while True:
        entrada = input(mensaje).strip()
        if not entrada:
            return opciones[str(defecto)][1]
        if entrada in opciones:
            return opciones[entrada][1]
        print(f"   [!] Opcion no valida. Ingrese un numero del 1 al {len(opciones)}.")


def pedir_si_no(mensaje, defecto="n"):
    while True:
        entrada = input(mensaje).strip().lower() or defecto
        if entrada in ("s", "si", "sí", "y", "yes"):
            return True
        if entrada in ("n", "no"):
            return False
        print("   [!] Responda 's' o 'n'.")


# =============================================================================
# 4) UTILIDADES NUMERICAS
# =============================================================================
def interpolar(x, xs, ys):
    """Interpolacion lineal con numpy (extrapola con el valor extremo)."""
    return float(np.interp(x, np.array(xs, dtype=float), np.array(ys, dtype=float)))


def rango_entre_claves(claves, x):
    """Devuelve el par de claves consecutivas que encierran a x."""
    ks = sorted(claves)
    if x <= ks[0]:
        return ks[0], ks[0]
    if x >= ks[-1]:
        return ks[-1], ks[-1]
    for i in range(len(ks) - 1):
        if ks[i] <= x <= ks[i + 1]:
            return ks[i], ks[i + 1]
    return ks[-1], ks[-1]


def formatear(valor, decimales=2, vacio="---"):
    if valor is None:
        return vacio
    return f"{valor:.{decimales}f}"


def redondear(valor, decimales=2):
    """Redondeo 'mitad hacia arriba' (evita el sesgo de round/banker's)."""
    if valor is None:
        return None
    cuantia = Decimal(1).scaleb(-decimales)
    base = Decimal(str(round(valor, decimales + 4)))
    return float(base.quantize(cuantia, rounding=ROUND_HALF_UP))


# --------------------- fin de Aymi Juzarit Escalante Ramos --------------------


# =============================================================================
#  INTEGRANTE 3 : José Salomón Cunya Carhuapoma
#  PARTE        : CAPTURA / INGRESO DE DATOS DEL DISENO
#  QUE DESARROLLA:
#    - La funcion solicitar_datos(): guia al usuario con TODAS las preguntas
#      necesarias para el diseno (proyecto, f'c, responsable, consistencia,
#      aire, cemento, agregado fino, agregado grueso, agua y probetas).
#    - Los MENUS NUMERADOS para las elecciones (tamano maximo nominal,
#      consistencia, exposicion, etc.), escritos solo con el numero.
#    - El uso de los VALORES POR DEFECTO del ejemplo "corregido": si el usuario
#      presiona ENTER, se toma el dato por defecto y se pasa al siguiente.
#    - mostrar_resumen_datos(): tabla de verificacion antes de calcular.
#  IMPORTANCIA : es la puerta de entrada de los datos; alimenta al motor de
#    calculo (integrante 4).
# =============================================================================
# =============================================================================
# 5) CONSULTA DE DATOS
# =============================================================================
def solicitar_datos():
    EJ = EJEMPLO_CORREGIDO   # valores por defecto = ejemplo resuelto "corregido"
    encabezado("INGRESO DE DATOS - DISENO DE MEZCLA ACI 211")
    print("Los valores por defecto son los del ejemplo resuelto 'CORREGIDO'.")
    print("Presione ENTER para aceptar cada valor por defecto.\n")

    d = {}

    # ---------------- DISENO ----------------
    print("--- DATOS DE DISENO ---")
    d["descripcion"] = pedir_texto(
        f"Descripcion del elemento/proyecto [{EJ['descripcion']}]: ", EJ["descripcion"])
    d["fc"] = pedir_flotante(
        f"Resistencia especificada f'c [kg/cm2] [{EJ['fc']:.0f}]: ",
        defecto=EJ["fc"], minimo=1)
    d["responsable"] = pedir_texto(
        f"Responsable / docente [{EJ['responsable']}]: ", EJ["responsable"])

    print("\nConsistencia de la mezcla:")
    opciones_consist = {
        "1": ('Seca (0" a 2")', "Seca"),
        "2": ('Plastica (3" a 4")', "Plastica"),
        "3": ('Fluida (>= 5")', "Fluida"),
    }
    defecto_consist = next(k for k, v in opciones_consist.items() if v[1] == EJ["consistencia"])
    d["consistencia"] = pedir_opcion(
        f"   Consistencia [{defecto_consist}]: ", opciones_consist, defecto_consist)

    d["con_aire"] = pedir_si_no(
        f"\nConcreto con aire incorporado? (s/n) [{'s' if EJ['con_aire'] else 'n'}]: ",
        "s" if EJ["con_aire"] else "n")
    if d["con_aire"]:
        opciones_expo = {"1": ("Suave", 0), "2": ("Moderada", 1), "3": ("Severa", 2)}
        defecto_expo = str(EJ["exposicion_idx"] + 1)
        exposicion = pedir_opcion(
            f"   Tipo de exposicion [{defecto_expo}]: ", opciones_expo, defecto_expo)
        d["exposicion_idx"] = exposicion
        d["exposicion_nombre"] = ["Suave", "Moderada", "Severa"][exposicion]
    else:
        d["exposicion_idx"] = None
        d["exposicion_nombre"] = "-"

    # ---------------- CEMENTO ----------------
    print("\n--- CEMENTO ---")
    d["cemento_nombre"] = pedir_texto(
        f"Tipo de cemento [{EJ['cemento_nombre']}]: ", EJ["cemento_nombre"])
    d["pe_cemento"] = pedir_flotante(
        f"Peso especifico del cemento [{EJ['pe_cemento']}]: ",
        defecto=EJ["pe_cemento"], minimo=1)

    # ---------------- AGREGADO FINO (ARENA) ----------------
    print("\n--- AGREGADO FINO (ARENA) ---")
    d["arena_puss"] = pedir_flotante(
        f"Peso unitario suelto seco [{EJ['arena_puss']:.0f}]: ",
        defecto=EJ["arena_puss"], minimo=1)
    d["arena_pucs"] = pedir_flotante(
        f"Peso unitario compactado seco [{EJ['arena_pucs']:.0f}]: ",
        defecto=EJ["arena_pucs"], minimo=1)
    d["arena_pe"] = pedir_flotante(
        f"Peso especifico de masa [{EJ['arena_pe']}]: ",
        defecto=EJ["arena_pe"], minimo=1)
    d["arena_humedad"] = pedir_flotante(
        f"% de humedad [{EJ['arena_humedad']}]: ",
        defecto=EJ["arena_humedad"], minimo=0)
    d["arena_absorcion"] = pedir_flotante(
        f"% de absorcion [{EJ['arena_absorcion']}]: ",
        defecto=EJ["arena_absorcion"], minimo=0)
    d["arena_mf"] = pedir_flotante(
        f"Modulo de fineza [{EJ['arena_mf']}]: ",
        defecto=EJ["arena_mf"], minimo=1)

    # ---------------- AGREGADO GRUESO (PIEDRA) ----------------
    print("\n--- AGREGADO GRUESO (PIEDRA) ---")
    d["piedra_puss"] = pedir_flotante(
        f"Peso unitario suelto seco [{EJ['piedra_puss']:.0f}]: ",
        defecto=EJ["piedra_puss"], minimo=1)
    d["piedra_pucs"] = pedir_flotante(
        f"Peso unitario compactado seco [{EJ['piedra_pucs']:.0f}]: ",
        defecto=EJ["piedra_pucs"], minimo=1)
    d["piedra_pe"] = pedir_flotante(
        f"Peso especifico de masa [{EJ['piedra_pe']}]: ",
        defecto=EJ["piedra_pe"], minimo=1)
    d["piedra_humedad"] = pedir_flotante(
        f"% de humedad [{EJ['piedra_humedad']}]: ",
        defecto=EJ["piedra_humedad"], minimo=0)
    d["piedra_absorcion"] = pedir_flotante(
        f"% de absorcion [{EJ['piedra_absorcion']}]: ",
        defecto=EJ["piedra_absorcion"], minimo=0)

    # TMN: orden y numeracion segun el PDF "yablas" (Tablas 01 a 09)
    #   1=3/8"  2=1/2"  3=3/4"  4=1"  5=1 1/2"  6=2"  7=3"  8=6"
    print("\nTamano maximo nominal (TMN) del agregado grueso:")
    opciones_tmn = {str(i + 1): (t, t) for i, t in enumerate(TMN_DISENO)}
    defecto_tmn = str(TMN_DISENO.index(EJ["tmn"]) + 1)
    d["tmn"] = pedir_opcion(
        f"   Elija el TMN [{defecto_tmn} = {EJ['tmn']}]: ", opciones_tmn, defecto_tmn)

    # ---------------- DOSIFICACION PARA PROBETAS ----------------
    print("\n--- DOSIFICACION PARA PROBETAS ---")
    d["probeta_d"] = pedir_flotante(
        f"Diametro de probeta [cm] [{EJ['probeta_d']:.0f}]: ",
        defecto=EJ["probeta_d"], minimo=1)
    d["probeta_h"] = pedir_flotante(
        f"Altura de probeta [cm] [{EJ['probeta_h']:.0f}]: ",
        defecto=EJ["probeta_h"], minimo=1)
    d["n_probetas"] = int(pedir_flotante(
        f"Numero de probetas [{EJ['n_probetas']}]: ",
        defecto=EJ["n_probetas"], minimo=1))
    d["desperdicio"] = pedir_flotante(
        f"% de desperdicio [{EJ['desperdicio']:.0f}]: ",
        defecto=EJ["desperdicio"], minimo=0)

    return d


def mostrar_resumen_datos(d):
    """Muestra en pantalla todos los datos ingresados antes de calcular."""
    encabezado("RESUMEN DE DATOS INGRESADOS")
    filas = [
        ["Resistencia f'c (kg/cm2)", f"{d['fc']:.0f}"],
        ["Descripcion", d["descripcion"]],
        ["Responsable", d["responsable"]],
        ["Consistencia", d["consistencia"]],
        ["Con aire incorporado", "Si" if d["con_aire"] else "No"],
        ["Exposicion", d["exposicion_nombre"]],
        ["Cemento", d["cemento_nombre"]],
        ["P.E. cemento", f"{d['pe_cemento']:.2f}"],
        ["ARENA - PUSS (kg/m3)", f"{d['arena_puss']:.0f}"],
        ["ARENA - PUCS (kg/m3)", f"{d['arena_pucs']:.0f}"],
        ["ARENA - P.E.", f"{d['arena_pe']:.2f}"],
        ["ARENA - Humedad (%)", f"{d['arena_humedad']:.1f}"],
        ["ARENA - Absorcion (%)", f"{d['arena_absorcion']:.1f}"],
        ["ARENA - Modulo de fineza", f"{d['arena_mf']:.1f}"],
        ["PIEDRA - PUSS (kg/m3)", f"{d['piedra_puss']:.0f}"],
        ["PIEDRA - PUCS (kg/m3)", f"{d['piedra_pucs']:.0f}"],
        ["PIEDRA - P.E.", f"{d['piedra_pe']:.2f}"],
        ["PIEDRA - Humedad (%)", f"{d['piedra_humedad']:.1f}"],
        ["PIEDRA - Absorcion (%)", f"{d['piedra_absorcion']:.1f}"],
        ["PIEDRA - TMN", d["tmn"]],
        ["PROBETA - Diametro (cm)", f"{d['probeta_d']:.0f}"],
        ["PROBETA - Altura (cm)", f"{d['probeta_h']:.0f}"],
        ["PROBETA - Numero", f"{d['n_probetas']}"],
        ["PROBETA - Desperdicio (%)", f"{d['desperdicio']:.0f}"],
    ]
    print(pd.DataFrame(filas, columns=["DATO", "VALOR"]).to_string(index=False))


# ------------------- fin de José Salomón Cunya Carhuapoma ---------------------


# =============================================================================
#  INTEGRANTE 4 : Anderson Jhimy Flores Alfaro
#  PARTE        : MOTOR DE CALCULO - METODO ACI 211 (15 PASOS)
#  QUE DESARROLLA:
#    - La funcion disenar(): el corazon del programa. Aplica el metodo de
#      volumenes absolutos del ACI 211 paso a paso:
#        f'cr, consistencia/asentamiento, relacion a/c, agua de mezclado,
#        cemento, aire atrapado, peso del agregado grueso (b/bo), volumenes
#        absolutos, correccion por humedad y absorcion, agua efectiva,
#        cemento humedo, proporcion en peso, proporcion en volumen por saco
#        y cantidades para las probetas.
#    - La interpolacion de valores entre filas de las Tablas ACI (usando las
#      utilidades del integrante 2 y las tablas de la integrante 1).
#  IMPORTANCIA : convierte los datos de entrada en la dosificacion final y
#    devuelve los resultados (R) y el paso a paso (S).
# =============================================================================
# =============================================================================
# 6) MOTOR DE CALCULO - METODO ACI 211 (15 PASOS)
# =============================================================================
def disenar(d):
    """Ejecuta el metodo ACI 211 y devuelve (R, secciones)."""
    R = {}
    S = []  # lista de secciones del reporte

    def sec(titulo, *bloques):
        """Agrega una seccion del reporte.

        Los bloques se guardan en el MISMO orden en que se pasan y pueden ser:
          - un DataFrame                     -> tabla
          - un texto (str)                   -> nota
          - una lista/tupla de notas         -> varias notas (compatibilidad)
          - una tupla ("tabla"/"nota", valor)-> bloque explicito
        Asi el Excel reproduce paso a paso el documento de referencia "corregido".
        """
        items = []
        for b in bloques:
            if b is None:
                continue
            if isinstance(b, pd.DataFrame):
                items.append(("tabla", b))
            elif isinstance(b, str):
                items.append(("nota", b))
            elif isinstance(b, (list, tuple)):
                for x in b:
                    if x is None:
                        continue
                    if isinstance(x, tuple) and len(x) == 2 and x[0] in ("tabla", "nota"):
                        items.append(x)
                    else:
                        items.append(("nota", str(x)))
            else:
                items.append(("nota", str(b)))
        S.append({
            "titulo": titulo,
            "bloques": items,
            "tabla": next((v for t, v in items if t == "tabla"), None),
            "notas": [v for t, v in items if t == "nota"],
        })

    # ------------------------------------------------------------------
    # ENCABEZADO / PROBLEMA
    # ------------------------------------------------------------------
    R["titulo"] = "DISENO DE MEZCLAS - METODO COMITE ACI 211"
    problema = (
        f"Se desea disenar una mezcla de concreto para ser utilizado en la "
        f"construccion de una {d['descripcion']}. Como dato de diseno se sabe "
        f"que la resistencia especificada por el proyectista es f'c = {d['fc']:.0f} Kg/cm2. "
        f"La mezcla es de consistencia {d['consistencia'].lower()}. "
        f"El cemento Portland a emplearse es {d['cemento_nombre']} con un peso "
        f"especifico de {d['pe_cemento']:.2f} gr/cm3."
    )
    sec("PROBLEMA", None, [problema])

    caracteristicas = pd.DataFrame(
        [
            ["Peso Unitario Suelto Seco", f"{d['arena_puss']:.0f} Kg/m3", f"{d['piedra_puss']:.0f} Kg/m3"],
            ["Peso Unitario Compactado Seco", f"{d['arena_pucs']:.0f} Kg/m3", f"{d['piedra_pucs']:.0f} Kg/m3"],
            ["Peso Especifico de Masa", f"{d['arena_pe']:.2f}", f"{d['piedra_pe']:.2f}"],
            ["Contenido de Humedad", f"{d['arena_humedad']:.1f} %", f"{d['piedra_humedad']:.1f} %"],
            ["Porcentaje de Absorcion", f"{d['arena_absorcion']:.1f} %", f"{d['piedra_absorcion']:.1f} %"],
            ["Modulo de Fineza", f"{d['arena_mf']:.1f}", "------"],
            ["Tamano Maximo del Agregado", "------", d["tmn"]],
        ],
        columns=["DESCRIPCION", "ARENA", "PIEDRA"],
    )
    sec("CARACTERISTICAS DE LOS AGREGADOS", caracteristicas)

    # ------------------------------------------------------------------
    # PASO 1: RESISTENCIA PROMEDIO f'cr
    # ------------------------------------------------------------------
    fc = d["fc"]
    for lo, hi, adic, regla in RESISTENCIA_PROMEDIO_REGLA:
        if lo <= fc <= hi:
            fcr = fc + adic
            regla_usada = regla
            break
    R["fc"] = fc
    R["fcr"] = fcr

    tabla_fcr = pd.DataFrame(
        [["Menos de 210", "F'c + 70"],
         ["210 a 350", "F'c + 84"],
         ["Sobre 350", "F'c + 98"]],
        columns=["F'c (Kg/cm2)", "F'cr"],
    )
    sec("PASO 1°: DETERMINAR LA RESISTENCIA PROMEDIO PARA EL DISEÑO", tabla_fcr,
        [f"F'cr = {regla_usada} = {fc:.0f} + {adic:.0f}",
         f"F'cr = {fcr:.0f} Kg/cm²"])

    # ------------------------------------------------------------------
    # PASO 2: ASENTAMIENTO
    # ------------------------------------------------------------------
    tabla_asent = pd.DataFrame(
        [["Seca", '0" a 2"'], ["Plastica", '3" a 4"'], ["Fluida", '>= 5"']],
        columns=["CONSISTENCIA", "ASENTAMIENTO"],
    )
    slump_col = None
    for nombre, rango, col, _ in TABLA_CONSISTENCIA:
        if nombre == d["consistencia"]:
            slump_col = col
            slump_rango = rango
            break
    R["slump_col"] = slump_col
    R["slump_rango"] = slump_rango
    sec("PASO 2°: DETERMINAR EL ASENTAMIENTO", tabla_asent,
        [f"Como la mezcla es de consistencia {d['consistencia'].lower()}, "
         f"el asentamiento adecuado es {slump_rango}."])

    # ------------------------------------------------------------------
    # PASO 3: RELACION AGUA-CEMENTO
    # ------------------------------------------------------------------
    idx_aire = 1 if d["con_aire"] else 0
    filas_ac = [[f, formatear(v[0], 2), formatear(v[1], 2)]
                for f, v in sorted(TABLA_05_AC.items())]
    tabla_ac = pd.DataFrame(filas_ac, columns=["F'cr (Kg/cm2)", "Sin aire incorp.", "Con aire incorp."])

    xs = [f for f in sorted(TABLA_05_AC) if TABLA_05_AC[f][idx_aire] is not None]
    ys = [TABLA_05_AC[f][idx_aire] for f in xs]
    ac_exacto = interpolar(fcr, xs, ys)
    ac = round(ac_exacto, 2)
    R["ac_exacto"] = ac_exacto
    R["ac"] = ac

    k1, k2 = rango_entre_claves(xs, fcr)
    v1, v2 = TABLA_05_AC[k1][idx_aire], TABLA_05_AC[k2][idx_aire]
    nota_ac = (f"Con F'cr = {fcr:.0f} Kg/cm2, interpolando entre {k2} (a/c={v2:.2f}) "
               f"y {k1} (a/c={v1:.2f}):")
    if k1 != k2:
        delta = (fcr - k1) / (k2 - k1) * (v2 - v1)
        nota_ac += (f"  X - {v1:.2f} = [({fcr:.0f} - {k1}) / ({k2} - {k1})] x "
                    f"({v2:.2f} - {v1:.2f}) = {delta:.3f}")
        nota_ac += f"  ->  X = {ac_exacto:.3f}"
    sec("PASO 3°: RELACIÓN AGUA-CEMENTO (a/c)", tabla_ac,
        [nota_ac, f"R a/c = {ac:.2f}"])

    # ------------------------------------------------------------------
    # PASO 4: AGUA DE MEZCLADO Y AIRE
    # ------------------------------------------------------------------
    tipo = "con_aire" if d["con_aire"] else "sin_aire"
    filas_agua = []
    for rango in ['1" a 2"', '3" a 4"', '6" a 7"']:
        fila = [rango] + [formatear(TABLA_01_AGUA[tipo][rango][t], 0) for t in TMN_DISENO]
        filas_agua.append(fila)
    tabla_agua = pd.DataFrame(filas_agua, columns=["Asentamiento"] + TMN_DISENO)

    agua = TABLA_01_AGUA[tipo][slump_col][d["tmn"]]
    R["agua"] = float(agua)
    tipo_txt = "con aire incorporado" if d["con_aire"] else "sin aire incorporado"
    sec("PASO 4°: CÁLCULO DE LA CANTIDAD DE AGUA DE MEZCLADO Y AIRE", tabla_agua,
        [f"TMN = {d['tmn']}, slump {slump_rango}, concreto {tipo_txt} "
         f"-> Agua = {agua:.0f} lt/m³"])

    # ------------------------------------------------------------------
    # PASO 5: CEMENTO
    # ------------------------------------------------------------------
    cemento = agua / ac
    R["cemento_diseno"] = cemento
    sec("PASO 5°: DETERMINAR EL CEMENTO", None,
        ["Ra/c = a / c   ->   c = a / Ra/c",
         f"c = {agua:.0f} / {ac:.2f}",
         f"c = {cemento:.2f} Kg/m³"])

    # ------------------------------------------------------------------
    # PASO 6: AIRE
    # ------------------------------------------------------------------
    if d["con_aire"]:
        aire_pct = TABLA_06_AIRE_INCORPORADO[d["tmn"]][d["exposicion_idx"]]
        tabla_aire = pd.DataFrame(
            [[t, f"{TABLA_06_AIRE_INCORPORADO[t][0]:.1f} %",
              f"{TABLA_06_AIRE_INCORPORADO[t][1]:.1f} %",
              f"{TABLA_06_AIRE_INCORPORADO[t][2]:.1f} %"] for t in TMN_DISENO],
            columns=["TMN", "Suave", "Moderada", "Severa"])
        nota = (f"TMN = {d['tmn']}, exposicion {d['exposicion_nombre']} "
                f"-> Aire incorporado total = {aire_pct:.1f} %")
        titulo6 = "PASO 6°: CONTENIDO DE AIRE INCORPORADO Y TOTAL"
    else:
        aire_pct = TABLA_02_AIRE_ATRAPADO[d["tmn"]]
        tabla_aire = pd.DataFrame(
            [[t, f"{TABLA_02_AIRE_ATRAPADO[t]:.1f} %"] for t in ORDEN_AIRE_ATRAPADO],
            columns=["Tamano Maximo Nominal", "Aire Atrapado"])
        nota = f"TMN = {d['tmn']} -> Aire Atrapado = {aire_pct:.1f} %"
        titulo6 = "PASO 6°: CÁLCULO DEL AIRE ATRAPADO"
    R["aire_pct"] = aire_pct
    sec(titulo6, tabla_aire, [nota])

    # ------------------------------------------------------------------
    # PASO 7: AGREGADO GRUESO
    # ------------------------------------------------------------------
    filas_bbo = []
    for t in TMN_DISENO:
        filas_bbo.append([t] + [f"{TABLA_04_AG_GRUESO[t][mf]:.2f}"
                                for mf in (2.40, 2.60, 2.80, 3.00)])
    tabla_bbo = pd.DataFrame(filas_bbo, columns=["TMN Ag. Grueso", "MF = 2.40",
                                                 "MF = 2.60", "MF = 2.80", "MF = 3.00"])

    mf = d["arena_mf"]
    mfs = sorted(TABLA_04_AG_GRUESO[d["tmn"]])
    bbo = round(interpolar(mf, mfs, [TABLA_04_AG_GRUESO[d["tmn"]][m] for m in mfs]), 2)
    peso_grueso = bbo * d["piedra_pucs"]
    R["bbo"] = bbo
    R["peso_grueso"] = peso_grueso

    k1, k2 = rango_entre_claves(mfs, mf)
    b1, b2 = TABLA_04_AG_GRUESO[d["tmn"]][k1], TABLA_04_AG_GRUESO[d["tmn"]][k2]
    nota_bbo = f"Con MF = {mf:.1f} y TMN = {d['tmn']}, interpolando entre MF {k2:.2f} ({b2:.2f}) y MF {k1:.2f} ({b1:.2f}):"
    if k1 != k2:
        nota_bbo += f"  ->  b/bo = {bbo:.2f}"
    sec("PASO 7°: DETERMINAR LA CANTIDAD DE AGREGADO GRUESO", tabla_bbo,
        [nota_bbo,
         f"Peso Ag. Grueso = {bbo:.2f} x {d['piedra_pucs']:.0f} = {peso_grueso:.1f} Kg/m³"])

    # ------------------------------------------------------------------
    # PASO 8: VOLUMENES ABSOLUTOS - AGREGADO FINO
    # ------------------------------------------------------------------
    vol_cemento = round(cemento / (d["pe_cemento"] * 1000), 4)
    vol_grueso = round(peso_grueso / (d["piedra_pe"] * 1000), 4)
    vol_agua = round(agua / (PESO_ESPECIFICO_AGUA * 1000), 4)
    vol_aire = round(aire_pct / 100.0, 4)
    vol_total = round(vol_cemento + vol_grueso + vol_agua + vol_aire, 4)
    vol_fino = round(1.0 - vol_total, 4)
    peso_fino = vol_fino * d["arena_pe"] * 1000

    R.update(dict(vol_cemento=vol_cemento, vol_grueso=vol_grueso, vol_agua=vol_agua,
                  vol_aire=vol_aire, vol_total=vol_total, vol_fino=vol_fino,
                  peso_fino=peso_fino))

    tabla_vol = pd.DataFrame(
        [["Cemento", f"{cemento:.2f}", f"{d['pe_cemento']:.2f}", f"{vol_cemento:.3f}"],
         ["Ag. Grueso", f"{peso_grueso:.1f}", f"{d['piedra_pe']:.2f}", f"{vol_grueso:.3f}"],
         ["Agua", f"{agua:.0f}", f"{PESO_ESPECIFICO_AGUA:.2f}", f"{vol_agua:.3f}"],
         [f"Aire ({aire_pct:.1f}%)", "---", "---", f"{vol_aire:.3f}"],
         ["TOTAL", "---", "---", f"{vol_total:.3f}"]],
        columns=["Material", "Peso Seco (Kg)", "P.E.", "Vol. Absoluto (m³)"])
    tabla_pesos_secos = pd.DataFrame(
        [["Cemento", f"{cemento:.2f}"], ["Ag. Fino", f"{peso_fino:.3f}"],
         ["Ag. Grueso", f"{peso_grueso:.1f}"], ["Agua", f"{agua:.0f}"]],
        columns=["Material", "Peso Seco (Kg)"])
    sec("PASO 8°: VOLÚMENES ABSOLUTOS — HALLAR EL AGREGADO FINO",
        "Volumen Absoluto = Peso Seco / (P.E. x 1000)",
        tabla_vol,
        f"Vol. Absoluto Ag. Fino = 1 - {vol_total:.3f} = {vol_fino:.3f} m³",
        f"Peso Seco Ag. Fino = {vol_fino:.3f} x {d['arena_pe']:.2f} x 1000 = {peso_fino:.3f} Kg",
        tabla_pesos_secos)

    # ------------------------------------------------------------------
    # PASO 9: CORRECCION POR HUMEDAD
    # ------------------------------------------------------------------
    fino_humedo = peso_fino * (d["arena_humedad"] / 100.0 + 1)
    grueso_humedo = peso_grueso * (d["piedra_humedad"] / 100.0 + 1)
    fino_humedo_red = round(fino_humedo)
    grueso_humedo_red = round(grueso_humedo)
    R.update(dict(fino_humedo=fino_humedo, grueso_humedo=grueso_humedo,
                  fino_humedo_red=fino_humedo_red, grueso_humedo_red=grueso_humedo_red))
    sec("PASO 9°: CORRECCIÓN POR HUMEDAD", None,
        ["Peso húmedo = Peso Seco x (% Humedad / 100 + 1)",
         f"Ag. Fino   = {peso_fino:.3f} x ({d['arena_humedad']:.1f}/100 + 1) = {fino_humedo_red} Kg",
         f"Ag. Grueso = {peso_grueso:.1f} x ({d['piedra_humedad']:.1f}/100 + 1) = "
         f"{grueso_humedo:.2f} ≈ {grueso_humedo_red} Kg"])

    # ------------------------------------------------------------------
    # PASO 10: CORRECCION POR ABSORCION (AGUA LIBRE)
    # ------------------------------------------------------------------
    agua_libre_fino = peso_fino * (d["arena_absorcion"] - d["arena_humedad"]) / 100.0
    agua_libre_grueso = peso_grueso * (d["piedra_absorcion"] - d["piedra_humedad"]) / 100.0
    agua_libre_total = round(agua_libre_fino, 2) + round(agua_libre_grueso, 2)
    R.update(dict(agua_libre_fino=agua_libre_fino, agua_libre_grueso=agua_libre_grueso,
                  agua_libre_total=agua_libre_total))
    sec("PASO 10°: CORRECCIÓN POR ABSORCIÓN", None,
        ["Agua libre = Peso Seco x (% Absorción − % Humedad) / 100",
         f"Ag. Fino   = {peso_fino:.3f} x ({d['arena_absorcion']:.1f} − {d['arena_humedad']:.1f}) / 100 = "
         f"{agua_libre_fino:.2f} Lt",
         f"Ag. Grueso = {peso_grueso:.1f} x ({d['piedra_absorcion']:.1f} − {d['piedra_humedad']:.1f}) / 100 = "
         f"+{agua_libre_grueso:.2f} Lt",
         f"Σ Agua Libre = {round(agua_libre_fino,2):.2f} + {round(agua_libre_grueso,2):.2f} = "
         f"{agua_libre_total:.2f} Lt"])

    # ------------------------------------------------------------------
    # PASO 11: AGUA EFECTIVA
    # ------------------------------------------------------------------
    agua_efectiva = agua + agua_libre_total
    R["agua_efectiva"] = agua_efectiva
    sec("PASO 11°: AGUA EFECTIVA", None,
        ["Agua efectiva = Agua de diseño + Agua libre",
         f"Agua efectiva = {agua:.0f} + ({agua_libre_total:.2f}) = {agua_efectiva:.2f} lt/m³"])

    # ------------------------------------------------------------------
    # PASO 12: DISENO TEORICO HUMEDO (cemento corregido)
    # ------------------------------------------------------------------
    cemento_humedo = redondear(agua_efectiva / ac, 2)
    R["cemento_humedo"] = cemento_humedo
    sec("PASO 12°: DISEÑO TEÓRICO HÚMEDO", None,
        ["Ra/c = Agua efectiva / Cemento   ->   C = Agua efectiva / Ra/c",
         f"C = {agua_efectiva:.2f} / {ac:.2f} = {agua_efectiva / ac:.3f}",
         f"C = {cemento_humedo:.2f} Kg/m³"])

    # ------------------------------------------------------------------
    # PASO 13: PROPORCIONES EN PESO
    # ------------------------------------------------------------------
    prop_fino = fino_humedo_red / cemento_humedo
    prop_grueso = grueso_humedo_red / cemento_humedo
    prop_agua = agua_efectiva / cemento_humedo
    R.update(dict(prop_fino=prop_fino, prop_grueso=prop_grueso, prop_agua=prop_agua))

    tabla_prop = pd.DataFrame(
        [["Cemento", f"{cemento_humedo:.2f}", f"/ {cemento_humedo:.2f}", "1"],
         ["Ag. Fino", f"{fino_humedo_red}", f"/ {cemento_humedo:.2f}", f"{prop_fino:.2f}"],
         ["Ag. Grueso", f"{grueso_humedo_red}", f"/ {cemento_humedo:.2f}", f"{prop_grueso:.2f}"],
         ["Agua", f"{agua_efectiva:.2f}", f"/ {cemento_humedo:.2f}", f"Ra/c = {prop_agua:.2f}"]],
        columns=["Material", "Peso Húmedo (Kg)", "/ Cemento", "Proporción"])
    sec("PASO 13°: PROPORCIONES EN PESO",
        f"Dividiendo cada material entre el peso del cemento ({cemento_humedo:.2f} Kg):",
        tabla_prop,
        f"Proporción en peso: 1 : {prop_fino:.2f} : {prop_grueso:.2f} con Ra/c = {prop_agua:.2f}")

    # ------------------------------------------------------------------
    # PASO 14: PROPORCIONES EN VOLUMEN
    # ------------------------------------------------------------------
    vol_fino_saco = (prop_fino * PESO_SACO_CEMENTO) / d["arena_puss"]
    vol_grueso_saco = (prop_grueso * PESO_SACO_CEMENTO) / d["piedra_puss"]
    agua_saco = prop_agua * PESO_SACO_CEMENTO
    R.update(dict(vol_fino_saco=vol_fino_saco, vol_grueso_saco=vol_grueso_saco,
                  agua_saco=agua_saco))

    tabla_vol_saco = pd.DataFrame(
        [["Cemento", "Referencia", "1", f"saco ({PESO_SACO_CEMENTO} Kg)"],
         ["Ag. Fino", f"({prop_fino:.2f}x{PESO_SACO_CEMENTO})/{d['arena_puss']:.0f}",
          f"{vol_fino_saco:.3f}", "m3"],
         ["Ag. Grueso", f"({prop_grueso:.2f}x{PESO_SACO_CEMENTO})/{d['piedra_puss']:.0f}",
          f"{vol_grueso_saco:.3f}", "m3"],
         ["Agua", f"{prop_agua:.2f}x{PESO_SACO_CEMENTO}", f"{agua_saco:.1f}", "Lt"]],
        columns=["Material", "Cálculo", "Proporción", "Unidad"])
    sec("PASO 14°: PROPORCIONES EN VOLUMEN (m³)",
        "Se convierte la proporción en peso a volumen usando el Peso Unitario "
        "Suelto Seco (PUSS) de cada agregado. El cemento se toma como referencia: "
        f"1 saco = {PESO_SACO_CEMENTO} Kg.",
        f"Fórmula general: Vol (m³) = (Proporción x {PESO_SACO_CEMENTO} Kg/saco) / PUSS (Kg/m³)",
        "Cemento: 1 saco (referencia fija)",
        f"Ag. Fino : ({prop_fino:.2f} x {PESO_SACO_CEMENTO}) / {d['arena_puss']:.0f} = "
        f"{prop_fino * PESO_SACO_CEMENTO:.3f} / {d['arena_puss']:.0f} = {vol_fino_saco:.3f} m³",
        f"Ag. Grueso : ({prop_grueso:.2f} x {PESO_SACO_CEMENTO}) / {d['piedra_puss']:.0f} = "
        f"{prop_grueso * PESO_SACO_CEMENTO:.3f} / {d['piedra_puss']:.0f} = {vol_grueso_saco:.3f} m³",
        f"Agua : {prop_agua:.2f} x {PESO_SACO_CEMENTO} = {agua_saco:.1f} Lt por saco de cemento",
        tabla_vol_saco,
        f"Proporción en volumen: 1 saco : {vol_fino_saco:.3f} m³ : {vol_grueso_saco:.3f} m³ "
        f"con {agua_saco:.1f} Lt de agua",
        "Lectura en obra: por cada saco de cemento se usan "
        f"{vol_fino_saco:.3f} m³ de arena y {vol_grueso_saco:.3f} m³ de piedra, medidos en suelto.")

    # ------------------------------------------------------------------
    # PASO 15: DOSIFICACION PARA PROBETAS
    # ------------------------------------------------------------------
    D_m = d["probeta_d"] / 100.0
    H_m = d["probeta_h"] / 100.0
    v_probeta = (np.pi / 4.0) * D_m ** 2 * H_m
    v_probetas = d["n_probetas"] * v_probeta
    v_total = round(v_probetas * (1 + d["desperdicio"] / 100.0), 4)

    cant_cemento = cemento_humedo * v_total
    cant_fino = fino_humedo_red * v_total
    cant_grueso = grueso_humedo_red * v_total
    cant_agua = agua_efectiva * v_total
    num_sacos = cant_cemento / PESO_SACO_CEMENTO
    obra_fino = vol_fino_saco * v_total
    obra_grueso = vol_grueso_saco * v_total
    obra_agua = agua_saco * v_total

    R.update(dict(v_probeta=v_probeta, v_probetas=v_probetas, v_total=v_total,
                  cant_cemento=cant_cemento, cant_fino=cant_fino, cant_grueso=cant_grueso,
                  cant_agua=cant_agua, num_sacos=num_sacos, obra_fino=obra_fino,
                  obra_grueso=obra_grueso, obra_agua=obra_agua))

    tabla_cant = pd.DataFrame(
        [["Cemento (húmedo)", f"{cemento_humedo:.2f} Kg/m³", f"x {v_total:.3f}",
          f"{cant_cemento:.3f} Kg"],
         ["Ag. Fino (húmedo)", f"{fino_humedo_red:.2f} Kg/m³", f"x {v_total:.3f}",
          f"{cant_fino:.3f} Kg"],
         ["Ag. Grueso (húmedo)", f"{grueso_humedo_red:.1f} Kg/m³", f"x {v_total:.3f}",
          f"{cant_grueso:.3f} Kg"],
         ["Agua efectiva", f"{agua_efectiva:.2f} Lt/m³", f"x {v_total:.3f}",
          f"{cant_agua:.3f} Lt"]],
        columns=["Material", "Dosis (Kg o Lt / m³)", "x V", "Cantidad"])

    tabla_obra = pd.DataFrame(
        [["Cemento", f"{cant_cemento:.3f} / {PESO_SACO_CEMENTO}", f"{num_sacos:.3f}", "sacos"],
         ["Ag. Fino", f"{vol_fino_saco:.3f} x {v_total:.3f}", f"{obra_fino:.3f}", "m3"],
         ["Ag. Grueso", f"{vol_grueso_saco:.3f} x {v_total:.3f}", f"{obra_grueso:.3f}", "m3"],
         ["Agua efectiva", f"{agua_saco:.1f} x {v_total:.3f}", f"{obra_agua:.3f}", "Lt"]],
        columns=["Material", "Cálculo", "Cantidad", "Unidad"])

    sec(f"PASO 15°: DOSIFICACIÓN PARA PROBETAS DE {d['probeta_d']:.0f}x{d['probeta_h']:.0f} cm",
        f"Datos: diámetro = {d['probeta_d']:.0f} cm ({(D_m):.2f} m), "
        f"altura = {d['probeta_h']:.0f} cm ({H_m:.2f} m)",
        f"Número de probetas = {d['n_probetas']}    Desperdicio = {d['desperdicio']:.0f}%",
        f"1. Volumen de una probeta cilíndrica: V = (π/4) x D² x H = "
        f"{np.pi/4:.3f} x {D_m:.2f}² x {H_m:.2f} = {v_probeta:.3f} m³/probeta",
        f"2. Volumen total para {d['n_probetas']} probetas: "
        f"{d['n_probetas']} x {v_probeta:.3f} = {v_probetas:.3f} m³",
        f"3. Volumen con desperdicio del {d['desperdicio']:.0f}%: "
        f"{v_probetas:.3f} x (1 + {d['desperdicio']/100:.2f}) = {v_total:.3f} m³",
        f"   Volumen total adoptado: V = {v_total:.3f} m³",
        "4. Cantidad de cada material = Dosis por m³ x Vtotal:",
        tabla_cant,
        "5. Equivalente en volumen de obra (por saco de cemento):",
        f"Número de sacos ≈ {cant_cemento:.3f} Kg ÷ {PESO_SACO_CEMENTO} Kg/saco = "
        f"{num_sacos:.3f} sacos",
        f"Ag. Fino   = {vol_fino_saco:.3f} m³ de arena",
        f"Ag. Grueso = {vol_grueso_saco:.3f} m³ de piedra",
        f"Agua       = {round(agua_saco,1):.1f} Lt",
        tabla_obra)

    # ------------------------------------------------------------------
    # RESUMEN FINAL
    # ------------------------------------------------------------------
    tabla_resumen = pd.DataFrame(
        [["Cemento", f"{cant_cemento:.3f} Kg", f"{num_sacos:.3f}", "sacos"],
         ["Ag. Fino", f"{cant_fino:.3f} Kg", f"{obra_fino:.3f}", "m3"],
         ["Ag. Grueso", f"{cant_grueso:.3f} Kg", f"{obra_grueso:.3f}", "m3"],
         ["Agua efectiva", f"{cant_agua:.3f} Lt", f"{obra_agua:.3f}", "Lt"]],
        columns=["Material", "Peso húmedo", "Volumen obra", "Unidad vol."])
    sec(f"RESUMEN — DOSIFICACIÓN PARA {d['n_probetas']} PROBETAS "
        f"({d['probeta_d']:.0f}x{d['probeta_h']:.0f} cm) + {d['desperdicio']:.0f}% DESPERDICIO",
        tabla_resumen,
        f"VOLUMEN TOTAL V = {v_total:.3f} m³")

    return R, S


# ------------------- fin de Anderson Jhimy Flores Alfaro ---------------------


# =============================================================================
#  INTEGRANTE 6 : Christian Llerena Huaracha
#  PARTE        : SALIDA DE RESULTADOS E INTEGRACION (une todos los modulos)
#  QUE DESARROLLA:
#    - El reporte y la presentacion de resultados: reporte de texto, resumen
#      final de la dosificacion en consola (mostrar_resumen_final).
#    - La EXPORTACION del informe a Excel con formato (exportar_excel) y el
#      visor de Tablas ACI (ver_tablas_aci).
#    - El MENU PRINCIPAL (main) que integra y EJECUTA todas las partes
#      anteriores: enlaza el modulo de resistencia del concreto (integrante 5)
#      y coordina la captura, el calculo y la salida.
#  IMPORTANCIA : es el modulo que "cose" el trabajo de todos los integrantes y
#    permite ejecutar el programa completo de principio a fin.
# =============================================================================
# =============================================================================
# 7) IMPRESION / REPORTE
# =============================================================================
def render_secciones(d, R, S):
    """Construye el reporte completo como lista de lineas de texto."""
    L = []
    L.append("=" * ANCHO)
    L.append("TECNOLOGIA DEL CONCRETO".center(ANCHO))
    L.append(R["titulo"].center(ANCHO))
    L.append("=" * ANCHO)
    L.append(f"Fecha         : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    L.append(f"Responsable   : {d['responsable']}")
    L.append(f"Resistencia   : f'c = {d['fc']:.0f} kg/cm2    "
             f"f'cr = {R['fcr']:.0f} kg/cm2")
    L.append(f"Relacion a/c  : {R['ac']:.2f}")
    L.append("")

    for s in S:
        L.append("-" * ANCHO)
        L.append(s["titulo"])
        L.append("-" * ANCHO)
        for tipo, contenido in s["bloques"]:
            if tipo == "tabla":
                L.append(contenido.to_string(index=False))
            elif contenido:
                L.append(str(contenido))
        L.append("")

    L.append("=" * ANCHO)
    L.append("PROPORCION FINAL EN OBRA".center(ANCHO))
    L.append("=" * ANCHO)
    L.append(f"  1 saco de cemento ({PESO_SACO_CEMENTO} Kg) : "
             f"{R['vol_fino_saco']:.3f} m3 de arena : {R['vol_grueso_saco']:.3f} m3 de piedra")
    L.append(f"  Agua: {R['agua_saco']:.1f} lt por saco    (a/c = {R['ac']:.2f})")
    L.append(f"  Cemento: {R['cemento_humedo']:.2f} kg/m3   Agua efectiva: {R['agua_efectiva']:.2f} lt/m3")
    L.append("=" * ANCHO)
    L.append(d["responsable"].center(ANCHO))
    L.append("=" * ANCHO)
    return L


def imprimir_reporte(lineas):
    for l in lineas:
        print(l)


# =============================================================================
# 8) EXPORTACION A EXCEL (openpyxl) - reporte con formato
# =============================================================================
C_TITULO = "1F3864"      # azul oscuro
C_SECCION = "2E75B6"     # azul medio
C_CABECERA = "D9E1F2"    # azul claro
C_CLAVE = "FCE4D6"       # naranja claro
C_TOTAL = "FFF2CC"       # amarillo claro
C_DATO = "F2F2F2"        # gris claro
C_RESULTADO = "C6E0B4"   # verde claro (resultados importantes)


def _alto_necesario(texto, ancho):
    """Estima cuantas lineas ocupa un texto en un ancho de columna dado."""
    ancho = max(6, int(ancho * 0.95))
    lineas = 0
    for parte in str(texto).split("\n"):
        lineas += max(1, (len(parte) + ancho - 1) // ancho)
    return max(1, lineas)


MARCAS_TEXTO = (
    "Proporción en peso:", "Proporción en volumen:", "Volumen total adoptado",
    "VOLUMEN TOTAL", "Número de sacos",
)
MARCAS_NUMERO = (
    "F'cr = ", "R a/c = ", "b/bo = ", "Aire Atrapado = ", "Agua = ",
    "Peso Ag. Grueso = ", "Vol. Absoluto Ag. Fino = ", "Peso Seco Ag. Fino = ",
    "\u03a3 Agua Libre = ", "Agua efectiva = ", "C = ",
)
_PATRONES_RESULTADO = tuple(re.compile(re.escape(m) + r"\s*-?\d") for m in MARCAS_NUMERO)
_PATRON_C_MIN = re.compile(r"(?<![\w'])c = \s*-?\d")


def _es_resultado(texto):
    """True si la nota es el resultado numerico importante de un paso."""
    t = " ".join(str(texto).split())
    if any(marca in t for marca in MARCAS_TEXTO):
        return True
    if _PATRON_C_MIN.search(t):
        return True
    return any(patron.search(t) for patron in _PATRONES_RESULTADO)


def _borde():
    lado = Side(style="thin", color="B0B0B0")
    return Border(left=lado, right=lado, top=lado, bottom=lado)


def _escribir_titulo(ws, fila, texto, ncols):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(row=fila, column=1, value=texto)
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=C_TITULO)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ancho_total = sum((ws.column_dimensions[get_column_letter(i)].width or 8.43)
                      for i in range(1, ncols + 1))
    ws.row_dimensions[fila].height = max(26.0, _alto_necesario(texto, ancho_total) * 16.0 + 6)
    return fila + 1


def _escribir_seccion(ws, fila, texto, ncols):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(row=fila, column=1, value=texto)
    c.font = Font(bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=C_SECCION)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ancho_total = sum((ws.column_dimensions[get_column_letter(i)].width or 8.43)
                      for i in range(1, ncols + 1))
    ws.row_dimensions[fila].height = max(19.0, _alto_necesario(texto, ancho_total) * 14.0 + 4)
    return fila + 1


def _escribir_cabecera(ws, fila, encabezados):
    max_lineas = 1
    for i, texto in enumerate(encabezados, start=1):
        c = ws.cell(row=fila, column=i, value=texto)
        c.font = Font(bold=True, color=C_TITULO)
        c.fill = PatternFill("solid", fgColor=C_CABECERA)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _borde()
        ancho = ws.column_dimensions[get_column_letter(i)].width or 8.43
        max_lineas = max(max_lineas, _alto_necesario(texto, ancho))
    ws.row_dimensions[fila].height = max(20.0, max_lineas * 14.0 + 5)
    return fila + 1


def _escribir_fila(ws, fila, valores, negrita=False, relleno=None, centrar_desde=2):
    max_lineas = 1
    for i, v in enumerate(valores, start=1):
        c = ws.cell(row=fila, column=i, value=v)
        c.border = _borde()
        c.font = Font(bold=negrita)
        if relleno:
            c.fill = PatternFill("solid", fgColor=relleno)
        if i >= centrar_desde:
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.alignment = Alignment(vertical="center", wrap_text=True)
        ancho = ws.column_dimensions[get_column_letter(i)].width or 8.43
        max_lineas = max(max_lineas, _alto_necesario(v, ancho))
    if max_lineas > 1:
        ws.row_dimensions[fila].height = max_lineas * 14.0 + 2
    return fila + 1


def _escribir_nota(ws, fila, texto, ncols):
    importante = _es_resultado(texto)
    if importante:   # resaltar toda la celda del resultado (antes de fusionar)
        for i in range(1, ncols + 1):
            ws.cell(row=fila, column=i).fill = PatternFill("solid", fgColor=C_RESULTADO)
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(row=fila, column=1, value=texto)
    if importante:
        c.font = Font(bold=True, size=10, color="1F3864")
    else:
        c.font = Font(italic=True, size=10, color="595959")
    c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ancho_total = sum((ws.column_dimensions[get_column_letter(i)].width or 8.43)
                      for i in range(1, ncols + 1))
    ws.row_dimensions[fila].height = max(15.0, _alto_necesario(texto, ancho_total) * 14.0)
    return fila + 1


def _ajustar_anchos(ws, anchos):
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho


def exportar_excel(d, R, S, ruta):
    """Genera el reporte Excel con hojas formateadas por tema."""
    wb = Workbook()
    wb.remove(wb.active)

    # ==================================================================
    # HOJA 1: DOSIFICACION FINAL
    # ==================================================================
    ws = wb.create_sheet("Dosificacion Final")
    _ajustar_anchos(ws, [28, 16, 18, 18, 12])
    f = 1
    f = _escribir_titulo(ws, f, "DOSIFICACION DE CONCRETO - METODO ACI 211", 5)
    f = _escribir_fila(ws, f, ["Proyecto", d["descripcion"], "",
                               "Fecha", datetime.now().strftime("%Y-%m-%d %H:%M")])
    f = _escribir_fila(ws, f, ["Responsable", d["responsable"], "",
                               "Consistencia", d["consistencia"]])
    f = _escribir_fila(ws, f, ["Resistencia f'c (kg/cm2)", f"{d['fc']:.0f}",
                               "Resistencia f'cr (kg/cm2)", f"{R['fcr']:.0f}",
                               f"a/c = {R['ac']:.2f}"])
    f += 1

    f = _escribir_seccion(ws, f, "DOSIFICACION PARA EL VOLUMEN CALCULADO", 5)
    f = _escribir_cabecera(ws, f, ["MATERIAL", "DOSIS POR m3", "PESO (kg)",
                                   "VOLUMEN DE OBRA", "UNIDAD"])
    f = _escribir_fila(ws, f, ["Cemento", f"{R['cemento_humedo']:.2f}",
                               f"{R['cant_cemento']:.3f}", f"{R['num_sacos']:.3f}", "sacos"],
                       relleno=C_DATO)
    f = _escribir_fila(ws, f, ["Agregado fino (arena)", f"{R['fino_humedo_red']:.2f}",
                               f"{R['cant_fino']:.3f}", f"{R['obra_fino']:.3f}", "m3"],
                       relleno=C_DATO)
    f = _escribir_fila(ws, f, ["Agregado grueso (piedra)", f"{R['grueso_humedo_red']:.1f}",
                               f"{R['cant_grueso']:.3f}", f"{R['obra_grueso']:.3f}", "m3"],
                       relleno=C_DATO)
    f = _escribir_fila(ws, f, ["Agua efectiva", f"{R['agua_efectiva']:.2f}",
                               f"{R['cant_agua']:.3f}", f"{R['obra_agua']:.3f}", "Lt"],
                       relleno=C_DATO)
    f = _escribir_fila(ws, f, ["VOLUMEN TOTAL", "", "", f"{R['v_total']:.3f}", "m3"],
                       negrita=True, relleno=C_TOTAL)
    f += 1

    f = _escribir_seccion(ws, f, "PROPORCIONES DE LA MEZCLA", 5)
    f = _escribir_cabecera(ws, f, ["EN PESO (cem : fino : grueso)", "a/c",
                                   "EN VOLUMEN (por saco)", "AGUA POR SACO",
                                   "CEMENTO POR m3"])
    f = _escribir_fila(ws, f, [f"1 : {R['prop_fino']:.2f} : {R['prop_grueso']:.2f}",
                               f"{R['ac']:.2f}",
                               f"{R['vol_fino_saco']:.3f} : {R['vol_grueso_saco']:.3f} m3",
                               f"{R['agua_saco']:.1f} Lt",
                               f"{R['cemento_humedo']:.2f} kg"],
                       negrita=True, relleno=C_RESULTADO)
    f += 1
    f = _escribir_nota(ws, f,
                       f"Por cada saco de cemento ({PESO_SACO_CEMENTO} kg) se usan "
                       f"{R['vol_fino_saco']:.3f} m3 de arena y {R['vol_grueso_saco']:.3f} m3 "
                       f"de piedra, medidos en suelto.", 5)
    f += 1

    f = _escribir_seccion(ws, f, f"DOSIFICACION PARA LAS PROBETAS (V total = {R['v_total']:.3f} m3)", 5)
    f = _escribir_cabecera(ws, f, ["MATERIAL", "CANTIDAD", "UNIDAD",
                                   "EQUIVALENTE EN OBRA", "UNIDAD"])
    f = _escribir_fila(ws, f, ["Cemento", f"{R['cant_cemento']:.3f}", "kg",
                               f"{R['num_sacos']:.3f}", "sacos"], relleno=C_DATO)
    f = _escribir_fila(ws, f, ["Agregado fino (arena)", f"{R['cant_fino']:.3f}", "kg",
                               f"{R['obra_fino']:.3f}", "m3"], relleno=C_DATO)
    f = _escribir_fila(ws, f, ["Agregado grueso (piedra)", f"{R['cant_grueso']:.3f}", "kg",
                               f"{R['obra_grueso']:.3f}", "m3"], relleno=C_DATO)
    f = _escribir_fila(ws, f, ["Agua efectiva", f"{R['cant_agua']:.3f}", "Lt",
                               f"{R['obra_agua']:.3f}", "Lt"], relleno=C_DATO)

    # ==================================================================
    # HOJA 2: DATOS DE ENTRADA (separados por material)
    # ==================================================================
    ws = wb.create_sheet("Datos de Entrada")
    _ajustar_anchos(ws, [40, 20, 14])
    f = 1
    f = _escribir_titulo(ws, f, "DATOS DE ENTRADA - PROYECTO", 3)

    def bloque(titulo, filas):
        nonlocal f
        f = _escribir_seccion(ws, f, titulo, 3)
        f = _escribir_cabecera(ws, f, ["PARAMETRO", "VALOR", "UNIDAD"])
        for i, (param, valor, unidad) in enumerate(filas):
            f = _escribir_fila(ws, f, [param, valor, unidad],
                               relleno=(C_DATO if i % 2 == 0 else None))

    bloque("DATOS DEL DISENO", [
        ("Descripcion del elemento", d["descripcion"], "-"),
        ("Resistencia especificada f'c", f"{d['fc']:.0f}", "kg/cm2"),
        ("Resistencia promedio f'cr", f"{R['fcr']:.0f}", "kg/cm2"),
        ("Consistencia", d["consistencia"], "-"),
        ("Asentamiento", R["slump_rango"], "-"),
        ("Con aire incorporado", "Si" if d["con_aire"] else "No", "-"),
        ("Exposicion", d["exposicion_nombre"], "-"),
    ])
    f += 1
    bloque("DATOS DEL CEMENTO", [
        ("Tipo de cemento", d["cemento_nombre"], "-"),
        ("Peso especifico", f"{d['pe_cemento']:.2f}", "g/cm3"),
        ("Dosis de cemento (diseno)", f"{R['cemento_diseno']:.2f}", "kg/m3"),
        ("Dosis de cemento (humedo corregido)", f"{R['cemento_humedo']:.2f}", "kg/m3"),
    ])
    f += 1
    bloque("DATOS DEL AGREGADO FINO (ARENA)", [
        ("Peso unitario suelto seco (PUSS)", f"{d['arena_puss']:.0f}", "kg/m3"),
        ("Peso unitario compactado seco (PUCS)", f"{d['arena_pucs']:.0f}", "kg/m3"),
        ("Peso especifico de masa", f"{d['arena_pe']:.2f}", "-"),
        ("Contenido de humedad", f"{d['arena_humedad']:.1f}", "%"),
        ("Porcentaje de absorcion", f"{d['arena_absorcion']:.1f}", "%"),
        ("Modulo de fineza", f"{d['arena_mf']:.2f}", "-"),
        ("Peso seco en la mezcla", f"{R['peso_fino']:.3f}", "kg/m3"),
        ("Peso humedo corregido", f"{R['fino_humedo_red']:.0f}", "kg/m3"),
    ])
    f += 1
    bloque("DATOS DEL AGREGADO GRUESO (PIEDRA)", [
        ("Peso unitario suelto seco (PUSS)", f"{d['piedra_puss']:.0f}", "kg/m3"),
        ("Peso unitario compactado seco (PUCS)", f"{d['piedra_pucs']:.0f}", "kg/m3"),
        ("Peso especifico de masa", f"{d['piedra_pe']:.2f}", "-"),
        ("Contenido de humedad", f"{d['piedra_humedad']:.1f}", "%"),
        ("Porcentaje de absorcion", f"{d['piedra_absorcion']:.1f}", "%"),
        ("Tamano maximo nominal (TMN)", d["tmn"], "-"),
        ("Volumen de agregado grueso (b/bo)", f"{R['bbo']:.2f}", "-"),
        ("Peso seco en la mezcla", f"{R['peso_grueso']:.1f}", "kg/m3"),
        ("Peso humedo corregido", f"{R['grueso_humedo_red']:.0f}", "kg/m3"),
    ])
    f += 1
    bloque("DATOS DEL AGUA", [
        ("Agua de mezclado (diseno)", f"{R['agua']:.0f}", "lt/m3"),
        ("Agua libre del agregado fino", f"{R['agua_libre_fino']:.2f}", "lt/m3"),
        ("Agua libre del agregado grueso", f"+{R['agua_libre_grueso']:.2f}", "lt/m3"),
        ("Agua libre total", f"{R['agua_libre_total']:.2f}", "lt/m3"),
        ("Agua efectiva", f"{R['agua_efectiva']:.2f}", "lt/m3"),
        ("Relacion agua/cemento (a/c)", f"{R['ac']:.2f}", "-"),
        ("Agua por saco de cemento", f"{R['agua_saco']:.1f}", "lt/saco"),
    ])
    f += 1
    bloque("DATOS DE LAS PROBETAS", [
        ("Diametro", f"{d['probeta_d']:.0f}", "cm"),
        ("Altura", f"{d['probeta_h']:.0f}", "cm"),
        ("Numero de probetas", f"{d['n_probetas']}", "-"),
        ("Desperdicio", f"{d['desperdicio']:.0f}", "%"),
        ("Volumen total a preparar", f"{R['v_total']:.3f}", "m3"),
    ])

    # ==================================================================
    # HOJA 3: PASO A PASO
    # ==================================================================
    ws = wb.create_sheet("Paso a Paso")
    _ajustar_anchos(ws, [34, 16, 16, 16, 16, 16, 16, 16, 16])
    f = 1
    f = _escribir_titulo(ws, f, "DISENO PASO A PASO - METODO ACI 211", 9)
    for s in S:
        f = _escribir_seccion(ws, f, s["titulo"], 9)
        for tipo, contenido in s["bloques"]:
            if tipo == "tabla":
                tabla = contenido
                f = _escribir_cabecera(ws, f, [str(c) for c in tabla.columns])
                for _, fila in tabla.iterrows():
                    f = _escribir_fila(ws, f, [str(v) for v in fila.values])
            elif contenido:
                f = _escribir_nota(ws, f, str(contenido), 9)
        f += 1

    # ==================================================================
    # HOJA 4: TABLAS ACI (referencia)
    # ==================================================================
    ws = wb.create_sheet("Tablas ACI")
    _ajustar_anchos(ws, [16, 14, 14, 14, 14, 14, 14, 14, 14])
    f = 1
    f = _escribir_titulo(ws, f, "TABLAS ACI 211 - VALORES DE REFERENCIA", 9)

    f = _escribir_seccion(ws, f, "TABLA 01 - VOLUMEN UNITARIO DE AGUA (lt/m3)", 9)
    for tipo, etiqueta in (("sin_aire", "CONCRETO SIN AIRE INCORPORADO"),
                           ("con_aire", "CONCRETO CON AIRE INCORPORADO")):
        f = _escribir_nota(ws, f, etiqueta, 9)
        f = _escribir_cabecera(ws, f, ["Asentamiento"] + TMN_DISENO)
        for r in ['1" a 2"', '3" a 4"', '6" a 7"']:
            f = _escribir_fila(ws, f, [r] + [formatear(TABLA_01_AGUA[tipo][r][t], 0)
                                             for t in TMN_DISENO])
    f += 1

    f = _escribir_seccion(ws, f, "TABLA 04 - PESO DEL AGREGADO GRUESO (b/bo)", 9)
    f = _escribir_cabecera(ws, f, ["TMN"] + [f"MF {m:.2f}" for m in (2.40, 2.60, 2.80, 3.00)])
    for t in TMN_DISENO:
        f = _escribir_fila(ws, f, [t] + [f"{TABLA_04_AG_GRUESO[t][m]:.2f}"
                                         for m in (2.40, 2.60, 2.80, 3.00)])
    f += 1

    f = _escribir_seccion(ws, f, "TABLA 05 - RELACION AGUA/CEMENTO POR RESISTENCIA", 9)
    f = _escribir_cabecera(ws, f, ["f'cr (kg/cm2)", "Sin aire", "Con aire"])
    for fc_, v in sorted(TABLA_05_AC.items()):
        f = _escribir_fila(ws, f, [fc_, formatear(v[0], 2), formatear(v[1], 2)])
    f += 1

    f = _escribir_seccion(ws, f, "TABLA 06 - CONTENIDO DE AIRE INCORPORADO Y TOTAL (%)", 9)
    f = _escribir_cabecera(ws, f, ["TMN", "Suave", "Moderada", "Severa"])
    for t, v in TABLA_06_AIRE_INCORPORADO.items():
        f = _escribir_fila(ws, f, [t, f"{v[0]:.1f}", f"{v[1]:.1f}", f"{v[2]:.1f}"])

    # ==================================================================
    # HOJA 5: GRAFICOS
    # ==================================================================
    ws = wb.create_sheet("Graficos")
    _ajustar_anchos(ws, [30, 18, 14, 14, 14, 14, 14, 14, 14])
    f = 1
    f = _escribir_titulo(ws, f, "GRAFICOS DE LA DOSIFICACION", 9)
    f = _escribir_nota(ws, f, "Los graficos se generan con los resultados del diseno.", 9)
    f += 1

    # --- Datos grafico 1: dosis por m3 ---
    f = _escribir_seccion(ws, f, "DOSIS POR m3 (Kg o Lt)", 9)
    fila_dosis = f + 1
    f = _escribir_cabecera(ws, f, ["Material", "Dosis por m3"])
    for nom, val in (("Cemento", R["cemento_humedo"]),
                     ("Ag. Fino", R["fino_humedo_red"]),
                     ("Ag. Grueso", R["grueso_humedo_red"]),
                     ("Agua", R["agua_efectiva"])):
        f = _escribir_fila(ws, f, [nom, round(float(val), 2)], relleno=C_DATO)
    fin_dosis = f - 1
    f += 1

    # --- Datos grafico 2: proporcion en volumen por saco ---
    f = _escribir_seccion(ws, f, "PROPORCION EN VOLUMEN POR SACO (m3)", 9)
    fila_vol = f + 1
    f = _escribir_cabecera(ws, f, ["Material", "Volumen (m3)"])
    for nom, val in (("Ag. Fino", round(R["vol_fino_saco"], 3)),
                     ("Ag. Grueso", round(R["vol_grueso_saco"], 3))):
        f = _escribir_fila(ws, f, [nom, val], relleno=C_DATO)
    fin_vol = f - 1
    f += 1

    # --- Datos grafico 3: cantidades para las probetas ---
    f = _escribir_seccion(ws, f, "CANTIDADES PARA LAS PROBETAS", 9)
    fila_prob = f + 1
    f = _escribir_cabecera(ws, f, ["Material", "Cantidad (Kg o Lt)"])
    for nom, val in (("Cemento", R["cant_cemento"]),
                     ("Ag. Fino", R["cant_fino"]),
                     ("Ag. Grueso", R["cant_grueso"]),
                     ("Agua", R["cant_agua"])):
        f = _escribir_fila(ws, f, [nom, round(float(val), 3)], relleno=C_DATO)
    fin_prob = f - 1

    # --- Grafico 1: barras de la dosis por m3 ---
    g1 = BarChart()
    g1.type = "col"
    g1.style = 10
    g1.title = "Dosis por m3 de concreto"
    g1.y_axis.title = "Kg o Lt"
    g1.add_data(Reference(ws, min_col=2, min_row=fila_dosis - 1, max_row=fin_dosis),
                titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=fila_dosis, max_row=fin_dosis))
    g1.height = 8.5
    g1.width = 15
    ws.add_chart(g1, "D2")

    # --- Grafico 2: circular de la proporcion en volumen ---
    g2 = PieChart()
    g2.title = "Proporcion en volumen por saco (m3)"
    g2.add_data(Reference(ws, min_col=2, min_row=fila_vol - 1, max_row=fin_vol),
                titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=fila_vol, max_row=fin_vol))
    g2.height = 8.5
    g2.width = 13
    ws.add_chart(g2, "D19")

    # --- Grafico 3: barras horizontales de las cantidades para probetas ---
    g3 = BarChart()
    g3.type = "bar"
    g3.style = 12
    g3.title = "Cantidades para las probetas"
    g3.y_axis.title = "Kg o Lt"
    g3.add_data(Reference(ws, min_col=2, min_row=fila_prob - 1, max_row=fin_prob),
                titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=fila_prob, max_row=fin_prob))
    g3.height = 8.5
    g3.width = 15
    ws.add_chart(g3, "D36")

    # ==================================================================
    # ORDEN FINAL DE LAS HOJAS (primero el paso a paso, como el PDF)
    # ==================================================================
    orden = ["Paso a Paso", "Dosificacion Final", "Graficos",
             "Datos de Entrada", "Tablas ACI"]
    wb._sheets = [wb[nombre] for nombre in orden]

    wb.save(ruta)


# =============================================================================
# 9) VISOR DE TABLAS ACI
# =============================================================================
def ver_tablas_aci():
    encabezado("TABLAS ACI 211 - CONSTANTES DEL PROGRAMA")

    print("\nTABLA 01 - VOLUMEN UNITARIO DE AGUA (lt/m3)")
    for tipo, etiqueta in (("sin_aire", "SIN AIRE INCORPORADO"), ("con_aire", "CON AIRE INCORPORADO")):
        df = pd.DataFrame([[r] + [formatear(TABLA_01_AGUA[tipo][r][t], 0) for t in TMN_DISENO]
                           for r in ['1" a 2"', '3" a 4"', '6" a 7"']],
                          columns=["Asentamiento"] + TMN_DISENO)
        print(f"\n  {etiqueta}:")
        print(df.to_string(index=False))

    print("\nTABLA 02 - CONTENIDO DE AIRE ATRAPADO")
    print(pd.DataFrame([[t, f"{TABLA_02_AIRE_ATRAPADO[t]:.1f} %"]
                        for t in sorted(TABLA_02_AIRE_ATRAPADO)],
                       columns=["TMN", "Aire Atrapado"]).to_string(index=False))

    print("\nTABLA 04 - PESO DEL AGREGADO GRUESO POR UNIDAD DE VOLUMEN (b/bo)")
    print(pd.DataFrame([[t] + [f"{TABLA_04_AG_GRUESO[t][m]:.2f}" for m in (2.40, 2.60, 2.80, 3.00)]
                        for t in TMN_DISENO],
                       columns=["TMN", "MF 2.40", "MF 2.60", "MF 2.80", "MF 3.00"]).to_string(index=False))

    print("\nTABLA 05 - RELACION AGUA/CEMENTO POR RESISTENCIA")
    print(pd.DataFrame([[f, formatear(v[0], 2), formatear(v[1], 2)]
                        for f, v in sorted(TABLA_05_AC.items())],
                       columns=["F'cr (kg/cm2)", "Sin aire", "Con aire"]).to_string(index=False))

    print("\nTABLA 06 - CONTENIDO DE AIRE INCORPORADO Y TOTAL (%)")
    print(pd.DataFrame([[t, f"{v[0]:.1f}", f"{v[1]:.1f}", f"{v[2]:.1f}"]
                        for t, v in TABLA_06_AIRE_INCORPORADO.items()],
                       columns=["TMN", "Suave", "Moderada", "Severa"]).to_string(index=False))

    print("\nTABLA 07 - CONDICIONES ESPECIALES DE EXPOSICION")
    print(pd.DataFrame([[e["condicion"], formatear(e["w_c_max"], 2), formatear(e["f_c_min_liviano"], 2)]
                        for e in TABLA_07_EXPOSICION],
                       columns=["Condicion", "w/c max", "f'c min liviano"]).to_string(index=False))
    print("  *", TABLA_07_NOTA)

    print("\nTABLA 09 - VOLUMEN UNITARIO DE AGUA POR FORMA (lt/m3, sin aire)")
    print(pd.DataFrame([[t, v[0], v[1], v[2], v[3], v[4], v[5]] for t, v in TABLA_09_AGUA_POR_FORMA.items()],
                       columns=["TMN", "1-2 red", "1-2 ang", "3-4 red", "3-4 ang", "6-7 red", "6-7 ang"]
                       ).to_string(index=False))
    print("\n  (Las Tablas 03 y 08 tambien estan disponibles como constantes en el codigo.)")
    pausa()


# =============================================================================
# 10) PROGRAMA PRINCIPAL
# =============================================================================
def mostrar_resumen_final(d, R):
    """Resumen compacto y claro de la dosificacion (pantalla final)."""
    sacos = R["cemento_humedo"] / PESO_SACO_CEMENTO

    encabezado("RESUMEN DE LA DOSIFICACION - ACI 211")

    print(f"  Proyecto      : {d['descripcion']}")
    print(f"  Responsable   : {d['responsable']}")
    print(f"  Fecha         : {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    linea("-")
    print(f"  Resistencia f'c   : {d['fc']:.0f} kg/cm2")
    print(f"  Resistencia f'cr  : {R['fcr']:.0f} kg/cm2")
    print(f"  Relacion a/c      : {R['ac']:.2f}")
    print(f"  Consistencia      : {d['consistencia']}  ({R['slump_rango']})")

    print()
    linea("-")
    print("  DOSIS POR m3 DE CONCRETO (peso humedo corregido)")
    linea("-")
    print(f"    Cemento ............. {R['cemento_humedo']:8.2f} kg/m3   ({sacos:.2f} sacos)")
    print(f"    Agregado fino ....... {R['fino_humedo_red']:8.2f} kg/m3")
    print(f"    Agregado grueso ..... {R['grueso_humedo_red']:8.1f} kg/m3")
    print(f"    Agua efectiva ....... {R['agua_efectiva']:8.2f} lt/m3")

    print()
    linea("-")
    print("  PROPORCION EN PESO  (cemento : fino : grueso)")
    linea("-")
    print(f"    1 : {R['prop_fino']:.2f} : {R['prop_grueso']:.2f}      a/c = {R['ac']:.2f}")

    print()
    linea("-")
    print(f"  PROPORCION EN VOLUMEN POR SACO DE CEMENTO ({PESO_SACO_CEMENTO} kg)")
    linea("-")
    print(f"    Agregado fino ....... {R['vol_fino_saco']:.3f} m3")
    print(f"    Agregado grueso ..... {R['vol_grueso_saco']:.3f} m3")
    print(f"    Agua ................ {R['agua_saco']:.1f} lt")

    print()
    linea("-")
    print(f"  CANTIDADES PARA EL VOLUMEN CALCULADO (V total = {R['v_total']:.3f} m3)")
    linea("-")
    print(f"    Cemento ............. {R['cant_cemento']:8.3f} kg   ({R['num_sacos']:.3f} sacos)")
    print(f"    Agregado fino ....... {R['cant_fino']:8.3f} kg")
    print(f"    Agregado grueso ..... {R['cant_grueso']:8.3f} kg")
    print(f"    Agua efectiva ....... {R['cant_agua']:8.3f} lt")
    linea("=")


def main():
    carpeta = os.path.dirname(os.path.abspath(__file__))

    while True:
        limpiar_pantalla()
        encabezado("TECNOLOGIA DEL CONCRETO - MENU PRINCIPAL")
        print("  1) Diseno de mezclas de concreto (ACI 211)")
        print("  2) Calculadora de resistencia del concreto (ACI 318 / ASTM)")
        print("  3) Ver tablas ACI 211 (constantes)")
        print("  0) Salir")
        linea("=")
        opcion = input("Seleccione una opcion: ").strip()

        if opcion == "1":
            limpiar_pantalla()
            d = solicitar_datos()

            limpiar_pantalla()
            mostrar_resumen_datos(d)
            if not pedir_si_no("\nLos datos son correctos? (s/n) [s]: ", "s"):
                pausa("\nVolviendo al menu... Presione ENTER.")
                continue

            limpiar_pantalla()
            print("\nCalculando diseno de mezcla...")
            R, S = disenar(d)

            # -------- PANTALLA FINAL: RESUMEN DE LA DOSIFICACION --------
            limpiar_pantalla()
            mostrar_resumen_final(d, R)

            # -------- SOLICITUD DEL INFORME (por defecto: SI) --------
            print("\nDesea hacer el informe de la dosificacion?")
            hacer_informe = pedir_opcion(
                "  Opcion [1]: ",
                {"1": ("Si, generar el informe", True),
                 "2": ("No", False)},
                "1",
            )

            if hacer_informe:
                marca = datetime.now().strftime("%Y%m%d_%H%M%S")
                ruta_xlsx = os.path.join(carpeta, f"diseno_aci211_{marca}.xlsx")
                try:
                    exportar_excel(d, R, S, ruta_xlsx)
                    print(f"\n  [OK] Informe guardado en:\n       {ruta_xlsx}")
                except Exception as error:
                    print(f"\n  [!] No se pudo generar el informe: {error}")
            else:
                print("\n  Informe no generado.")

            pausa()

        elif opcion == "2":
            limpiar_pantalla()
            resistencia_concreto.main()

        elif opcion == "3":
            limpiar_pantalla()
            ver_tablas_aci()

        elif opcion == "0":
            return
        else:
            print("\n  [!] Opcion no valida.")
            pausa()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nPrograma interrumpido.")
    print("\nHasta luego!\n")


# -------------------- fin de Christian Llerena Huaracha ----------------------
# =============================================================================
#  FIN DEL PROGRAMA PRINCIPAL
#  TECNOLOGIA DEL CONCRETO - ACI 211 / ACI 318
# =============================================================================
