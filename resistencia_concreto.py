# -*- coding: utf-8 -*-
"""
================================================================
 CALCULADORA DE RESISTENCIA DEL CONCRETO
================================================================
Programa de consola para el cálculo y analisis de la resistencia
del concreto, conforme a criterios del ACI 318-19 y ensayos ASTM.

Modulos incluidos
-----------------
 1. Resistencia a compresion en probeta cilindrica   (ASTM C39)
 2. Resistencia a compresion en probeta cubica       (EN 12390-3)
 3. Resistencia promedio requerida f'cr              (ACI 318-19)
 4. Modulo de elasticidad Ec                         (ACI 318-19)
 5. Modulo de ruptura a flexion fr                   (ACI 318-19)
 6. Resistencia a tension por compresion diametral   (ASTM C496)
 7. Conversion de unidades de esfuerzo
 8. Analisis estadistico de un conjunto de probetas
 9. Generacion de reporte en archivo de texto

Unidades internas: fuerza en Newton (N) y longitud en milimetros (mm).
Los resultados de esfuerzo se entregan en MPa, kg/cm2 y psi.
================================================================
"""


# =============================================================================
#  INTEGRANTE 5 : Luis Angel Espinoza Farfan
#  PARTE        : MODULO DE RESISTENCIA DEL CONCRETO (archivo hermano)
#  QUE DESARROLLA:
#    - La calculadora de resistencia del concreto (ACI 318-19 y ASTM):
#      compresion en probeta cilindrica (C39) y cubica (EN 12390-3),
#      resistencia promedio requerida f'cr, modulo de elasticidad Ec, modulo
#      de ruptura fr, traccion por compresion diametral (C496), conversion de
#      unidades, analisis estadistico de probetas y generacion de reporte.
#    - Su propio submenu de consola: se ejecuta desde el MENU PRINCIPAL
#      (integrante 6) y tambien puede correr de forma independiente.
#  IMPORTANCIA : es el segundo modulo funcional del programa; complementa el
#    diseno de mezclas con el analisis de resistencia del concreto.
# =============================================================================

import math
import os
import sys
from datetime import datetime

# ------------------------------------------------------------------
# Compatibilidad de salida en consola (acentos y simbolos)
# ------------------------------------------------------------------
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


# ==================================================================
# CONSTANTES
# ==================================================================
FACTOR_MPA_A_KGCM2 = 10.19716213
FACTOR_MPA_A_PSI = 145.0377377
FACTOR_KGCM2_A_MPA = 1.0 / FACTOR_MPA_A_KGCM2
FACTOR_PSI_A_MPA = 1.0 / FACTOR_MPA_A_PSI

# Peso volumetrico de referencia del concreto normal (kg/m3) - ACI 318-19
PESO_VOL_CONCRETO_NORMAL = 2320.0

# Conversion de fuerzas a Newton
FUERZA_A_NEWTON = {
    "1": ("kN", 1000.0),
    "2": ("tf", 9806.65),
    "3": ("kgf", 9.80665),
    "4": ("lbf", 4.4482216152605),
    "5": ("N", 1.0),
}

# Conversion de longitudes a milimetros
LONGITUD_A_MM = {
    "1": ("mm", 1.0),
    "2": ("cm", 10.0),
    "3": ("m", 1000.0),
    "4": ("in", 25.4),
}

# Tabla de correccion por relacion L/D para cilindros - ASTM C39
TABLA_FACTOR_LD = [
    (2.00, 1.00),
    (1.75, 0.98),
    (1.50, 0.96),
    (1.25, 0.93),
    (1.00, 0.87),
]

# ------------------------------------------------------------------
# Reportero global: acumula los calculos realizados en la sesion
# ------------------------------------------------------------------
BLOQUES_REPORTE = []   # lista de (titulo, [lineas])


# ==================================================================
# UTILIDADES DE ENTRADA / SALIDA
# ==================================================================
def limpiar_pantalla():
    os.system("cls" if os.name == "nt" else "clear")


def separador(caracter="=", ancho=64):
    print(caracter * ancho)


def titulo(texto, caracter="=", ancho=64):
    print()
    print(caracter * ancho)
    print(texto.center(ancho))
    print(caracter * ancho)


def pausa(mensaje="\nPresione ENTER para continuar..."):
    input(mensaje)


def pedir_flotante(mensaje, minimo=None, maximo=None, defecto=None):
    """Solicita un numero decimal validado dentro de un rango opcional."""
    while True:
        entrada = input(mensaje).strip().replace(",", ".")
        if not entrada and defecto is not None:
            return float(defecto)
        if not entrada:
            print("  [!] Debe ingresar un valor.")
            continue
        try:
            valor = float(entrada)
        except ValueError:
            print("  [!] Valor no valido. Ingrese un numero (ejemplo: 25.5).")
            continue
        if minimo is not None and valor < minimo:
            print("  [!] El valor debe ser mayor o igual a {}.".format(minimo))
            continue
        if maximo is not None and valor > maximo:
            print("  [!] El valor debe ser menor o igual a {}.".format(maximo))
            continue
        return valor


def pedir_entero(mensaje, minimo=None, maximo=None):
    """Solicita un numero entero validado dentro de un rango opcional."""
    while True:
        entrada = input(mensaje).strip()
        try:
            valor = int(entrada)
        except ValueError:
            print("  [!] Valor no valido. Ingrese un numero entero.")
            continue
        if minimo is not None and valor < minimo:
            print("  [!] El valor debe ser mayor o igual a {}.".format(minimo))
            continue
        if maximo is not None and valor > maximo:
            print("  [!] El valor debe ser menor o igual a {}.".format(maximo))
            continue
        return valor


def pedir_si_no(mensaje, defecto="n"):
    """Pregunta de si/no. Retorna True o False."""
    while True:
        entrada = input(mensaje).strip().lower()
        if not entrada:
            entrada = defecto
        if entrada in ("s", "si", "sí", "y", "yes"):
            return True
        if entrada in ("n", "no"):
            return False
        print("  [!] Responda 's' o 'n'.")


def elegir_unidad(mensaje, tabla):
    """Muestra un menu de unidades y retorna (nombre, factor_a_base)."""
    print(mensaje)
    for clave, (nombre, _) in tabla.items():
        print("   {}) {}".format(clave, nombre))
    while True:
        opcion = input("   Opcion: ").strip()
        if opcion in tabla:
            return tabla[opcion]
        print("  [!] Opcion no valida.")


# ==================================================================
# CONVERSIONES
# ==================================================================
def mpa_a_kgcm2(valor):
    return valor * FACTOR_MPA_A_KGCM2


def mpa_a_psi(valor):
    return valor * FACTOR_MPA_A_PSI


def formato_esfuerzo(mpa):
    """Devuelve una cadena con el esfuerzo en las tres unidades usuales."""
    return "{:.2f} MPa  |  {:.1f} kg/cm2  |  {:.1f} psi".format(
        mpa, mpa_a_kgcm2(mpa), mpa_a_psi(mpa)
    )


# ==================================================================
# FUNCIONES DE CALCULO
# ==================================================================
def area_circulo(diametro):
    """Area de un circulo en mm2 a partir del diametro en mm."""
    return math.pi * diametro ** 2 / 4.0


def factor_correccion_ld(relacion_ld):
    """
    Factor de correccion por relacion longitud/diametro (ASTM C39).
    Interpola linealmente entre los valores tabulados.
    """
    if relacion_ld >= 2.00:
        return 1.00
    if relacion_ld <= 1.00:
        return 0.87
    for i in range(len(TABLA_FACTOR_LD) - 1):
        ld_alto, factor_alto = TABLA_FACTOR_LD[i]
        ld_bajo, factor_bajo = TABLA_FACTOR_LD[i + 1]
        if ld_bajo <= relacion_ld <= ld_alto:
            pendiente = (factor_alto - factor_bajo) / (ld_alto - ld_bajo)
            return factor_bajo + pendiente * (relacion_ld - ld_bajo)
    return 1.00


def resistencia_compresion(carga_n, area_mm2):
    """Esfuerzo a compresion en MPa. f = P / A."""
    return carga_n / area_mm2


def resistencia_promedio_requerida(fc, proporciona_datos=False, desviacion=None):
    """
    Resistencia promedio requerida f'cr segun ACI 318-19.

    Sin datos historicos (19.3.2.2):
        f'c <= 35 MPa -> f'cr = f'c + 7.0
        f'c  > 35 MPa -> f'cr = 1.10 f'c + 5.0

    Con datos historicos (19.3.2.1 y 19.3.2.2):
        f'cr1 = f'c + 1.34 s
        f'cr2 = f'c + 2.33 s - 3.45
        f'cr  = mayor de f'cr1, f'cr2 y el minimo de 19.3.2.2
    """
    minimo = fc + 7.0 if fc <= 35.0 else 1.10 * fc + 5.0

    if not proporciona_datos or desviacion is None:
        return minimo, minimo, minimo, None

    fcr1 = fc + 1.34 * desviacion
    fcr2 = fc + 2.33 * desviacion - 3.45
    calculado = max(fcr1, fcr2)
    fcr_final = max(calculado, minimo)
    return fcr_final, fcr1, fcr2, minimo


def modulo_elasticidad(fc, peso_vol=PESO_VOL_CONCRETO_NORMAL):
    """
    Modulo de elasticidad del concreto en MPa (ACI 318-19, 19.2.2.1):
        Ec = 0.043 * wc^1.5 * sqrt(f'c)      (wc en kg/m3)
    Para concreto de peso normal (wc = 2320 kg/m3):
        Ec = 4700 * sqrt(f'c)
    """
    return 0.043 * peso_vol ** 1.5 * math.sqrt(fc)


def modulo_ruptura(fc, lam=1.0):
    """
    Modulo de ruptura (flexion) en MPa (ACI 318-19, 19.2.3.1):
        fr = 0.62 * lambda * sqrt(f'c)
    lambda = 1.0 para concreto de peso normal.
    """
    return 0.62 * lam * math.sqrt(fc)


def tension_diametral(carga_n, longitud_mm, diametro_mm):
    """
    Resistencia a tension por compresion diametral en MPa (ASTM C496):
        fsp = 2P / (pi * L * d)
    """
    return 2.0 * carga_n / (math.pi * longitud_mm * diametro_mm)


def estadisticas(valores):
    """Calcula media, desviacion estandar muestral, coeficiente de variacion."""
    n = len(valores)
    if n == 0:
        return None
    media = sum(valores) / n
    if n > 1:
        varianza = sum((x - media) ** 2 for x in valores) / (n - 1)
        desviacion = math.sqrt(varianza)
    else:
        desviacion = 0.0
    cov = (desviacion / media * 100.0) if media else 0.0
    return {
        "n": n,
        "media": media,
        "desviacion": desviacion,
        "cov": cov,
        "minimo": min(valores),
        "maximo": max(valores),
    }


# ==================================================================
# REGISTRO EN EL REPORTE
# ==================================================================
def registrar(titulo_bloque, lineas):
    """Guarda un bloque de resultados para el reporte final."""
    BLOQUES_REPORTE.append((titulo_bloque, list(lineas)))


# ==================================================================
# MODULO 1: COMPRESION EN PROBETA CILINDRICA
# ==================================================================
def modulo_compresion_cilindro():
    titulo("1. RESISTENCIA A COMPRESION - PROBETA CILINDRICA")
    print("Ensayo ASTM C39. El esfuerzo se calcula como f = P / A.\n")

    desc = input("Descripcion / identificacion de la probeta [Probeta]: ").strip()
    desc = desc or "Probeta"

    nombre_fuerza, factor_fuerza = elegir_unidad(
        "\nUnidad de la carga de rotura:", FUERZA_A_NEWTON
    )
    carga = pedir_flotante("Carga de rotura P [{}]: ".format(nombre_fuerza), minimo=0.01)
    carga_n = carga * factor_fuerza

    nombre_long, factor_long = elegir_unidad(
        "\nUnidad de las dimensiones:", LONGITUD_A_MM
    )
    diametro = pedir_flotante("Diametro d [{}]: ".format(nombre_long), minimo=0.01) * factor_long
    longitud = pedir_flotante(
        "Longitud (altura) L [{}] (0 = no aplicar correccion): ".format(nombre_long),
        minimo=0.0,
    ) * factor_long

    area = area_circulo(diametro)
    esfuerzo_bruto = resistencia_compresion(carga_n, area)

    # Correccion por relacion L/D
    factor = 1.0
    relacion_ld = None
    if longitud > 0:
        relacion_ld = longitud / diametro
        factor = factor_correccion_ld(relacion_ld)
    esfuerzo_corr = esfuerzo_bruto * factor

    print("\n--- RESULTADOS ---")
    print("  Area de la seccion     : {:.2f} mm2".format(area))
    print("  Esfuerzo bruto P/A     : {}".format(formato_esfuerzo(esfuerzo_bruto)))
    if relacion_ld is not None:
        print("  Relacion L/D           : {:.3f}".format(relacion_ld))
        print("  Factor de correccion   : {:.3f}".format(factor))
    print("  RESISTENCIA f'c        : {}".format(formato_esfuerzo(esfuerzo_corr)))

    lineas = [
        "Descripcion          : {}".format(desc),
        "Tipo de probeta      : Cilindrica (ASTM C39)",
        "Carga de rotura P    : {:.2f} {}  ({:.1f} N)".format(carga, nombre_fuerza, carga_n),
        "Diametro d           : {:.2f} mm".format(diametro),
        "Longitud L           : {:.2f} mm".format(longitud) if longitud > 0 else "Longitud L           : no especificada",
        "Area A               : {:.2f} mm2".format(area),
    ]
    if relacion_ld is not None:
        lineas.append("Relacion L/D         : {:.3f}".format(relacion_ld))
        lineas.append("Factor de correccion : {:.3f} (ASTM C39)".format(factor))
    lineas.append("Esfuerzo bruto P/A   : {}".format(formato_esfuerzo(esfuerzo_bruto)))
    lineas.append("RESISTENCIA f'c      : {}".format(formato_esfuerzo(esfuerzo_corr)))

    registrar("1. COMPRESION - PROBETA CILINDRICA (ASTM C39)", lineas)
    pausa()


# ==================================================================
# MODULO 2: COMPRESION EN PROBETA CUBICA
# ==================================================================
def modulo_compresion_cubo():
    titulo("2. RESISTENCIA A COMPRESION - PROBETA CUBICA")
    print("Ensayo EN 12390-3. El esfuerzo se calcula como f = P / a^2.\n")

    desc = input("Descripcion / identificacion de la probeta [Cubo]: ").strip()
    desc = desc or "Cubo"

    nombre_fuerza, factor_fuerza = elegir_unidad(
        "\nUnidad de la carga de rotura:", FUERZA_A_NEWTON
    )
    carga = pedir_flotante("Carga de rotura P [{}]: ".format(nombre_fuerza), minimo=0.01)
    carga_n = carga * factor_fuerza

    nombre_long, factor_long = elegir_unidad(
        "\nUnidad de las dimensiones:", LONGITUD_A_MM
    )
    lado = pedir_flotante("Lado de la cara a [{}]: ".format(nombre_long), minimo=0.01) * factor_long

    area = lado ** 2
    esfuerzo = resistencia_compresion(carga_n, area)

    print("\n--- RESULTADOS ---")
    print("  Area de la cara        : {:.2f} mm2".format(area))
    print("  RESISTENCIA f'c        : {}".format(formato_esfuerzo(esfuerzo)))

    lineas = [
        "Descripcion          : {}".format(desc),
        "Tipo de probeta      : Cubica (EN 12390-3)",
        "Carga de rotura P    : {:.2f} {}  ({:.1f} N)".format(carga, nombre_fuerza, carga_n),
        "Lado a               : {:.2f} mm".format(lado),
        "Area A               : {:.2f} mm2".format(area),
        "RESISTENCIA f'c      : {}".format(formato_esfuerzo(esfuerzo)),
    ]

    registrar("2. COMPRESION - PROBETA CUBICA (EN 12390-3)", lineas)
    pausa()


# ==================================================================
# MODULO 3: RESISTENCIA PROMEDIO REQUERIDA f'cr
# ==================================================================
def modulo_fcr():
    titulo("3. RESISTENCIA PROMEDIO REQUERIDA f'cr (ACI 318-19)")

    fc = pedir_flotante("Resistencia especificada f'c [MPa]: ", minimo=0.1)

    print("\nSe dispone de datos historicos (desviacion estandar s)?")
    tiene_datos = pedir_si_no("   (s/n) [n]: ")
    desviacion = None
    if tiene_datos:
        desviacion = pedir_flotante(
            "   Desviacion estandar s [MPa] (>= 15 ensayos): ", minimo=0.0
        )

    fcr, fcr1, fcr2, minimo = resistencia_promedio_requerida(fc, tiene_datos, desviacion)

    print("\n--- RESULTADOS ---")
    if not tiene_datos:
        print("  Criterio aplicado      : sin datos historicos (19.3.2.2)")
        print("  f'cr requerido         : {}".format(formato_esfuerzo(fcr)))
    else:
        print("  f'cr1 = f'c + 1.34 s   : {:.2f} MPa".format(fcr1))
        print("  f'cr2 = f'c + 2.33 s - 3.45 : {:.2f} MPa".format(fcr2))
        print("  Minimo de la norma     : {:.2f} MPa".format(minimo))
        print("  f'cr requerido         : {}".format(formato_esfuerzo(fcr)))

    lineas = [
        "Resistencia especificada f'c : {:.2f} MPa".format(fc),
    ]
    if not tiene_datos:
        lineas.append("Criterio            : sin datos historicos (ACI 318-19, 19.3.2.2)")
        lineas.append("f'cr requerido      : {}".format(formato_esfuerzo(fcr)))
    else:
        lineas.append("Desviacion estandar s        : {:.3f} MPa".format(desviacion))
        lineas.append("f'cr1 = f'c + 1.34 s          : {:.2f} MPa".format(fcr1))
        lineas.append("f'cr2 = f'c + 2.33 s - 3.45   : {:.2f} MPa".format(fcr2))
        lineas.append("Minimo normativo             : {:.2f} MPa".format(minimo))
        lineas.append("f'cr requerido               : {}".format(formato_esfuerzo(fcr)))

    registrar("3. RESISTENCIA PROMEDIO REQUERIDA f'cr (ACI 318-19)", lineas)
    pausa()


# ==================================================================
# MODULO 4: MODULO DE ELASTICIDAD
# ==================================================================
def modulo_modulo_elasticidad():
    titulo("4. MODULO DE ELASTICIDAD Ec (ACI 318-19)")

    fc = pedir_flotante("Resistencia f'c [MPa]: ", minimo=0.1)
    print("\nPeso volumetrico del concreto wc [kg/m3]")
    print("   (concreto de peso normal = 2320)")
    wc = pedir_flotante("   wc [kg/m3] [2320]: ", minimo=1440.0, maximo=2560.0, defecto=2320.0)

    ec = modulo_elasticidad(fc, wc)
    ec_simple = 4700.0 * math.sqrt(fc)

    print("\n--- RESULTADOS ---")
    print("  Ec = 0.043 * wc^1.5 * raiz(f'c)")
    print("  MODULO DE ELASTICIDAD Ec : {:.0f} MPa  |  {:.0f} kg/cm2".format(
        ec, mpa_a_kgcm2(ec)))
    print("  Referencia (Ec = 4700*raiz(f'c)) : {:.0f} MPa".format(ec_simple))

    lineas = [
        "Resistencia f'c      : {:.2f} MPa".format(fc),
        "Peso volumetrico wc  : {:.0f} kg/m3".format(wc),
        "Formula              : Ec = 0.043 * wc^1.5 * sqrt(f'c)",
        "MODULO Ec            : {:.0f} MPa  ({:.0f} kg/cm2)".format(ec, mpa_a_kgcm2(ec)),
        "Referencia 4700*raiz(f'c) : {:.0f} MPa".format(ec_simple),
    ]

    registrar("4. MODULO DE ELASTICIDAD Ec (ACI 318-19, 19.2.2.1)", lineas)
    pausa()


# ==================================================================
# MODULO 5: MODULO DE RUPTURA (FLEXION)
# ==================================================================
def modulo_modulo_ruptura():
    titulo("5. MODULO DE RUPTURA A FLEXION fr (ACI 318-19)")

    fc = pedir_flotante("Resistencia f'c [MPa]: ", minimo=0.1)
    liviano = pedir_si_no("\nConcreto liviano? (s/n) [n]: ")
    lam = 0.75 if liviano else 1.0

    fr = modulo_ruptura(fc, lam)

    print("\n--- RESULTADOS ---")
    print("  fr = 0.62 * lambda * raiz(f'c)   con lambda = {:.2f}".format(lam))
    print("  MODULO DE RUPTURA fr : {}".format(formato_esfuerzo(fr)))

    lineas = [
        "Resistencia f'c      : {:.2f} MPa".format(fc),
        "Tipo de concreto     : {}".format("Liviano" if liviano else "Peso normal"),
        "Factor lambda        : {:.2f}".format(lam),
        "Formula              : fr = 0.62 * lambda * sqrt(f'c)",
        "MODULO DE RUPTURA fr : {}".format(formato_esfuerzo(fr)),
    ]

    registrar("5. MODULO DE RUPTURA - FLEXION (ACI 318-19, 19.2.3.1)", lineas)
    pausa()


# ==================================================================
# MODULO 6: TRACCION POR COMPRESION DIAMETRAL
# ==================================================================
def modulo_tension_diametral():
    titulo("6. TRACCION POR COMPRESION DIAMETRAL (ASTM C496)")

    nombre_fuerza, factor_fuerza = elegir_unidad(
        "\nUnidad de la carga de rotura:", FUERZA_A_NEWTON
    )
    carga = pedir_flotante("Carga maxima P [{}]: ".format(nombre_fuerza), minimo=0.01)
    carga_n = carga * factor_fuerza

    nombre_long, factor_long = elegir_unidad(
        "\nUnidad de las dimensiones:", LONGITUD_A_MM
    )
    diametro = pedir_flotante("Diametro d [{}]: ".format(nombre_long), minimo=0.01) * factor_long
    longitud = pedir_flotante("Longitud L [{}]: ".format(nombre_long), minimo=0.01) * factor_long

    fsp = tension_diametral(carga_n, longitud, diametro)

    print("\n--- RESULTADOS ---")
    print("  fsp = 2P / (pi * L * d)")
    print("  RESISTENCIA A TRACCION fsp : {}".format(formato_esfuerzo(fsp)))

    lineas = [
        "Tipo de ensayo       : Compresion diametral (ASTM C496)",
        "Carga maxima P       : {:.2f} {}  ({:.1f} N)".format(carga, nombre_fuerza, carga_n),
        "Diametro d           : {:.2f} mm".format(diametro),
        "Longitud L           : {:.2f} mm".format(longitud),
        "Formula              : fsp = 2P / (pi * L * d)",
        "RESISTENCIA fsp      : {}".format(formato_esfuerzo(fsp)),
    ]

    registrar("6. TRACCION POR COMPRESION DIAMETRAL (ASTM C496)", lineas)
    pausa()


# ==================================================================
# MODULO 7: CONVERSION DE UNIDADES
# ==================================================================
def modulo_conversion():
    titulo("7. CONVERSION DE UNIDADES DE ESFUERZO")

    print("Ingrese el esfuerzo y su unidad de origen:\n")
    opciones = {
        "1": ("MPa", 1.0),
        "2": ("kg/cm2", FACTOR_KGCM2_A_MPA),
        "3": ("psi", FACTOR_PSI_A_MPA),
    }
    nombre, factor = elegir_unidad("Unidad de origen:", opciones)
    valor = pedir_flotante("Valor ({}) : ".format(nombre), minimo=0.0)
    mpa = valor * factor

    print("\n--- RESULTADOS ---")
    print("  Valor en MPa    : {:.3f}".format(mpa))
    print("  Valor en kg/cm2 : {:.3f}".format(mpa_a_kgcm2(mpa)))
    print("  Valor en psi    : {:.3f}".format(mpa_a_psi(mpa)))

    lineas = [
        "Valor de entrada     : {:.3f} {}".format(valor, nombre),
        "En MPa               : {:.3f}".format(mpa),
        "En kg/cm2            : {:.3f}".format(mpa_a_kgcm2(mpa)),
        "En psi               : {:.3f}".format(mpa_a_psi(mpa)),
    ]

    registrar("7. CONVERSION DE UNIDADES DE ESFUERZO", lineas)
    pausa()


# ==================================================================
# MODULO 8: ANALISIS ESTADISTICO DE PROBETAS
# ==================================================================
def modulo_estadistico():
    titulo("8. ANALISIS ESTADISTICO DE UN CONJUNTO DE PROBETAS")
    print("Ingrese las resistencias individuales (f'c) en MPa.")
    print("Se calcularan media, desviacion estandar y coeficiente de variacion.\n")

    n = pedir_entero("Cantidad de probetas: ", minimo=1, maximo=200)
    valores = []
    for i in range(n):
        v = pedir_flotante("  f'c probeta {} [MPa]: ".format(i + 1), minimo=0.0)
        valores.append(v)

    est = estadisticas(valores)

    # Resistencia caracteristica (criterio del 5% - media - 1.645 s)
    fck = est["media"] - 1.645 * est["desviacion"]

    print("\n--- RESULTADOS ---")
    print("  Numero de probetas     : {}".format(est["n"]))
    print("  Resistencia media      : {:.2f} MPa".format(est["media"]))
    print("  Desviacion estandar s  : {:.3f} MPa".format(est["desviacion"]))
    print("  Coef. de variacion     : {:.2f} %".format(est["cov"]))
    print("  Valor minimo           : {:.2f} MPa".format(est["minimo"]))
    print("  Valor maximo           : {:.2f} MPa".format(est["maximo"]))
    print("  f'ck (media - 1.645 s) : {:.2f} MPa".format(fck))

    lineas = [
        "Numero de probetas        : {}".format(est["n"]),
        "Resistencias individuales : " + ", ".join("{:.2f}".format(v) for v in valores),
        "Resistencia media         : {:.2f} MPa".format(est["media"]),
        "Desviacion estandar s     : {:.3f} MPa".format(est["desviacion"]),
        "Coeficiente de variacion  : {:.2f} %".format(est["cov"]),
        "Valor minimo              : {:.2f} MPa".format(est["minimo"]),
        "Valor maximo              : {:.2f} MPa".format(est["maximo"]),
        "f'ck (media - 1.645 s)    : {:.2f} MPa".format(fck),
    ]

    registrar("8. ANALISIS ESTADISTICO DE PROBETAS", lineas)
    pausa()


# ==================================================================
# MODULO 9: GENERACION DE REPORTE
# ==================================================================
def generar_reporte():
    titulo("9. GENERAR REPORTE")

    if not BLOQUES_REPORTE:
        print("No se han realizado calculos en esta sesion.")
        print("Realice al menos un calculo antes de generar el reporte.")
        pausa()
        return

    print("Se incluiran {} bloque(s) de calculo.\n".format(len(BLOQUES_REPORTE)))
    proyecto = input("Nombre del proyecto [Sin titulo]: ").strip() or "Sin titulo"
    autor = input("Realizado por [Anonimo]: ").strip() or "Anonimo"

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ancho = 64
    lineas = []
    lineas.append("=" * ancho)
    lineas.append("REPORTE DE CALCULO - RESISTENCIA DEL CONCRETO".center(ancho))
    lineas.append("=" * ancho)
    lineas.append("Proyecto      : {}".format(proyecto))
    lineas.append("Realizado por : {}".format(autor))
    lineas.append("Fecha         : {}".format(fecha))
    lineas.append("Bloques       : {}".format(len(BLOQUES_REPORTE)))
    lineas.append("")

    for i, (titulo_bloque, contenido) in enumerate(BLOQUES_REPORTE, start=1):
        lineas.append("-" * ancho)
        lineas.append(titulo_bloque)
        lineas.append("-" * ancho)
        for linea in contenido:
            lineas.append("   " + linea)
        lineas.append("")

    lineas.append("=" * ancho)
    lineas.append("FIN DEL REPORTE".center(ancho))
    lineas.append("=" * ancho)

    texto = "\n".join(lineas)

    # Nombre de archivo con marca de tiempo
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = "reporte_concreto_{}.txt".format(marca)
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre_archivo)

    try:
        with open(ruta, "w", encoding="utf-8") as archivo:
            archivo.write(texto)
    except OSError as error:
        print("\n  [!] No se pudo guardar el reporte: {}".format(error))
        pausa()
        return

    print("\n--- REPORTE GENERADO ---")
    print("  Archivo: {}".format(ruta))
    print("\n--- VISTA PREVIA ---\n")
    print(texto)

    pausa()


# ==================================================================
# MENU PRINCIPAL
# ==================================================================
def mostrar_menu():
    limpiar_pantalla()
    print()
    separador("=")
    print(" CALCULADORA DE RESISTENCIA DEL CONCRETO ".center(64, " "))
    separador("=")
    print("  1) Compresion - probeta cilindrica   (ASTM C39)")
    print("  2) Compresion - probeta cubica       (EN 12390-3)")
    print("  3) Resistencia promedio requerida f'cr (ACI 318-19)")
    print("  4) Modulo de elasticidad Ec          (ACI 318-19)")
    print("  5) Modulo de ruptura a flexion fr    (ACI 318-19)")
    print("  6) Traccion por compresion diametral (ASTM C496)")
    print("  7) Conversion de unidades de esfuerzo")
    print("  8) Analisis estadistico de probetas")
    print("  9) Generar reporte")
    print("  0) Salir")
    separador("=")
    if BLOQUES_REPORTE:
        print("  Calculos en esta sesion: {}".format(len(BLOQUES_REPORTE)))


def main():
    while True:
        mostrar_menu()
        opcion = input("Seleccione una opcion: ").strip()

        if opcion == "1":
            limpiar_pantalla()
            modulo_compresion_cilindro()
        elif opcion == "2":
            limpiar_pantalla()
            modulo_compresion_cubo()
        elif opcion == "3":
            limpiar_pantalla()
            modulo_fcr()
        elif opcion == "4":
            limpiar_pantalla()
            modulo_modulo_elasticidad()
        elif opcion == "5":
            limpiar_pantalla()
            modulo_modulo_ruptura()
        elif opcion == "6":
            limpiar_pantalla()
            modulo_tension_diametral()
        elif opcion == "7":
            limpiar_pantalla()
            modulo_conversion()
        elif opcion == "8":
            limpiar_pantalla()
            modulo_estadistico()
        elif opcion == "9":
            limpiar_pantalla()
            generar_reporte()
        elif opcion == "0":
            return
        else:
            print("\n  [!] Opcion no valida. Intente de nuevo.")
            pausa()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nPrograma interrumpido por el usuario.")
    print("\nHasta luego! Gracias por usar la calculadora.\n")


# -------------------- fin de Luis Angel Espinoza Farfan ----------------------
# =============================================================================
#  FIN DEL MODULO DE RESISTENCIA DEL CONCRETO
# =============================================================================
