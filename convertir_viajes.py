# -*- coding: utf-8 -*-
"""
convertir_viajes.py — Lee TODOS los Excel de la carpeta de viajes nacionales
y genera data_viajes.js con la inteligencia de flota por placa:

  - Viajes reales = conteo de "Envio" (manifiesto) DISTINTO por placa
  - Fecha del último viaje y días sin cargar
  - Estado de flota:
      FIDELIZADA     cargó en los últimos 30 días
      EVENTUAL       cargó entre 31 y 90 días
      POR RECUPERAR  más de 90 días sin cargar
  - Rotación = viajes promedio por mes (desde su primer viaje)
  - Historial agrupado por ORIGEN → DESTINO → CLIENTE con conteo de viajes

USO:
  1. Edita RUTA_CARPETA_VIAJES con tu carpeta de OneDrive.
  2. Ejecuta:  python convertir_viajes.py
     (o simplemente actualizar.bat, que corre todo junto)

Requiere:  pip install openpyxl
"""

import json
import os
import sys
import glob
from datetime import datetime, date

# ============================================================
# ⚙️ CONFIGURACIÓN — CARPETAS DE LAS OPERACIONES
# Cada operación tiene su carpeta con los Excel (divididos por meses).
# El script lee TODAS. Si alguna carpeta no existe, la salta avisando.
# Pon la ruta de la CARPETA CONTENEDORA (la que tiene NACIONAL, CEDIS, etc.)
# ============================================================
RUTA_BASE = "C:/Users/jarias/OneDrive - TRACTOCAR LOGISTICS SAS/POWER BI JEFFER/ARCHIVOS"

CARPETAS_OPERACION = {
    "NACIONAL": RUTA_BASE + "/NACIONAL",
    "CEDIS":    RUTA_BASE + "/CEDIS",
    "IMPO":     RUTA_BASE + "/IMPU",
    "EXPO":     RUTA_BASE + "/EXPO",
}

# Cada operación tiene columnas con nombres distintos. Aquí definimos
# qué columna del Excel corresponde a cada dato que necesitamos.
# "fila_encabezado" indica en qué fila está el encabezado (IMPO/EXPO traen
# una fila de título arriba, así que su encabezado está en la fila 2).
MAPEO_OPERACION = {
    "NACIONAL": {
        "fila_encabezado": 1,
        "cols": {
            "envio": "Envio", "placa": "Placa", "origen": "Ciudad Origen",
            "destino": "Ciudad Destino", "tipologia": "Tipologia",
            "fecha": "Fecha Creacion", "cliente": "Cliente", "operacion": "Operacion",
            "contable": "Cuenta Contable",
            "proveedor": "Afiliado",
        },
    },
    "CEDIS": {
        "fila_encabezado": 1,
        "cols": {
            "envio": "Manifiesto", "placa": "Placa (Veh)", "origen": "Origen (Man)",
            "destino": "Destino (Man)", "tipologia": "Tipologia",
            "fecha": "Creacion (Man)", "cliente": "Cliente (Orden)", "operacion": "Operacion (Orden)",
            "contable": "Contable (Man)",
            "proveedor": "Nom Afiliado",
        },
    },
    "IMPO": {
        "fila_encabezado": 2,
        "cols": {
            "envio": "Envio(compra)", "placa": "Placa", "origen": "Ciudad Origen",
            "destino": "Ciudad Destino", "tipologia": "Tipologia",
            "fecha": "Fecha Creacion", "cliente": "Cliente Nombre", "operacion": "Operacion",
            "proveedor": "Proveedor Nombre",
        },
    },
    "EXPO": {
        "fila_encabezado": 2,
        "cols": {
            "envio": "Envio(compra)", "placa": "Placa", "origen": "Ciudad Origen",
            "destino": "Ciudad Destino", "tipologia": "Tipologia",
            "fecha": "Fecha Creacion", "cliente": "Cliente Nombre", "operacion": "Operacion",
            "proveedor": "Proveedor Nombre",
        },
    },
}

SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_viajes.js")

# Umbrales de estado (en días)
DIAS_FIDELIZADA = 30
DIAS_EVENTUAL = 90

# ============================================================
# 🔁 UNIFICACIÓN DE CLIENTES
# Algunos clientes vienen escritos de varias formas. Aquí defines
# el nombre OFICIAL y todas las variantes que deben convertirse en él.
#
# Formato:  "NOMBRE OFICIAL": ["variante 1", "variante 2", ...]
# (escribe todo en MAYÚSCULAS; el script compara en mayúsculas)
#
# Para agregar un grupo nuevo: copia una línea y edítala.
# ============================================================
UNIFICAR_CLIENTES = {
    "ALIMENTOS POLAR SAS": [
        "ALIMENTOS POLAR COLOMBIA SAS",
        "ALIMENTOS POLAR COLOMBIA S.A.S",
    ],
    "ESENTTIA": [
        "ESENTTIA BY PROPILCO",
        "ESENTTIA MASTERBATCH LTDA",
        "ESENTTIA MARTERBATCH CROSS",   # nota: "MARTERBATCH" es un error de tipeo en el origen
        "ESENTTIA MASTERBATCH CROSS",
    ],
    # "NOMBRE OFICIAL": ["VARIANTE A", "VARIANTE B"],
}

# Se construye un diccionario inverso variante -> oficial (en mayúsculas)
_MAPA_CLIENTES = {}
for _oficial, _variantes in UNIFICAR_CLIENTES.items():
    _MAPA_CLIENTES[_oficial.strip().upper()] = _oficial.strip().upper()
    for _v in _variantes:
        _MAPA_CLIENTES[_v.strip().upper()] = _oficial.strip().upper()


def unificar_cliente(nombre):
    """Devuelve el nombre oficial si el cliente está en la tabla; si no, lo deja igual."""
    return _MAPA_CLIENTES.get(nombre, nombre)


# ============================================================
# 🗺️ CIUDAD -> DEPARTAMENTO
# Mapa de las ciudades que aparecen en los viajes a su departamento.
# Si una ciudad no está aquí, su departamento queda "OTRO" (revísalo en la
# lista que imprime el script al final y agrégala).
# ============================================================
CIUDAD_DEPTO = {
    # --- AMAZONAS ---
    "LETICIA": "AMAZONAS",
    # --- ANTIOQUIA ---
    "ABEJORRAL": "ANTIOQUIA", "AMAGA": "ANTIOQUIA", "AMALFI": "ANTIOQUIA",
    "ANDES": "ANTIOQUIA", "APARTADO": "ANTIOQUIA", "ARBOLETES": "ANTIOQUIA",
    "BARBOSA": "ANTIOQUIA", "BELLAVISTA": "ANTIOQUIA", "BELLO": "ANTIOQUIA",
    "BELMIRA": "ANTIOQUIA", "BURITICA": "ANTIOQUIA", "CALDAS": "ANTIOQUIA",
    "CANASGORDAS": "ANTIOQUIA", "CAREPA": "ANTIOQUIA", "CARMEN DE VIBORAL": "ANTIOQUIA",
    "CAUCASIA": "ANTIOQUIA", "CHIGORODO": "ANTIOQUIA", "CISNEROS": "ANTIOQUIA",
    "CIUDAD BOLIVAR": "ANTIOQUIA", "CONCEPCION": "ANTIOQUIA", "COPACABANA": "ANTIOQUIA",
    "DABEIBA": "ANTIOQUIA", "EL BAGRE": "ANTIOQUIA", "EL CARMEN DE VIBORAL": "ANTIOQUIA",
    "EL RETIRO": "ANTIOQUIA", "EL SANTUARIO": "ANTIOQUIA", "ENTRERRIOS": "ANTIOQUIA",
    "ENVIGADO": "ANTIOQUIA", "FREDONIA": "ANTIOQUIA", "FRONTINO": "ANTIOQUIA",
    "GIRARDOTA": "ANTIOQUIA", "GUADALUPE": "ANTIOQUIA", "GUARNE": "ANTIOQUIA",
    "GUATAPE": "ANTIOQUIA", "ITAGUI": "ANTIOQUIA", "LA CEJA": "ANTIOQUIA",
    "LA ESTRELLA": "ANTIOQUIA", "LIBORINA": "ANTIOQUIA", "MARINILLA": "ANTIOQUIA",
    "MEDELLIN": "ANTIOQUIA", "MUTATA": "ANTIOQUIA", "NECHI": "ANTIOQUIA",
    "NECOCLI": "ANTIOQUIA", "PENOL": "ANTIOQUIA", "PUERTO BERRIO": "ANTIOQUIA",
    "PUERTO TRIUNFO": "ANTIOQUIA", "REMEDIOS": "ANTIOQUIA", "RETIRO": "ANTIOQUIA",
    "RIONEGRO": "ANTIOQUIA", "SABANETA": "ANTIOQUIA", "SAN ANTONIO DE PRADO": "ANTIOQUIA",
    "SAN CARLOS": "ANTIOQUIA", "SAN CRISTOBAL": "ANTIOQUIA", "SAN FELIX": "ANTIOQUIA",
    "SAN JERONIMO": "ANTIOQUIA", "SAN JUAN DE URABA": "ANTIOQUIA", "SAN LUIS": "ANTIOQUIA",
    "SAN PEDRO": "ANTIOQUIA", "SAN PEDRO DE URABA": "ANTIOQUIA", "SAN RAFAEL": "ANTIOQUIA",
    "SANTA BARBARA": "ANTIOQUIA", "SANTA FE DE ANTIOQUIA": "ANTIOQUIA", "SANTA ROSA DE OSOS": "ANTIOQUIA",
    "SANTAFE DE ANTIOQUIA": "ANTIOQUIA", "SANTUARIO": "ANTIOQUIA", "SEGOVIA": "ANTIOQUIA",
    "SONSON": "ANTIOQUIA", "SOPETRAN": "ANTIOQUIA", "TAMESIS": "ANTIOQUIA",
    "TARAZA": "ANTIOQUIA", "TURBO": "ANTIOQUIA", "URRAO": "ANTIOQUIA",
    "VENECIA": "ANTIOQUIA", "YALI": "ANTIOQUIA", "YARUMAL": "ANTIOQUIA",
    "ZARAGOZA": "ANTIOQUIA",
    # --- ARAUCA ---
    "ARAUCA": "ARAUCA", "ARAUQUITA": "ARAUCA", "SARAVENA": "ARAUCA",
    "TAME": "ARAUCA",
    # --- ATLANTICO ---
    "BARANOA": "ATLANTICO", "BARRANQUILLA": "ATLANTICO", "GALAPA": "ATLANTICO",
    "JUAN DE ACOSTA": "ATLANTICO", "LURUACO": "ATLANTICO", "MALAMBO": "ATLANTICO",
    "MANATI": "ATLANTICO", "PALMAR DE VARELA": "ATLANTICO", "PONEDERA": "ATLANTICO",
    "PUERTO COLOMBIA": "ATLANTICO", "REPELON": "ATLANTICO", "SABANAGRANDE": "ATLANTICO",
    "SABANALARGA": "ATLANTICO", "SOLEDAD": "ATLANTICO", "TUBARA": "ATLANTICO",
    # --- BOLIVAR ---
    "ARJONA": "BOLIVAR", "ARROYOHONDO": "BOLIVAR", "BAYUNCA": "BOLIVAR",
    "CARTAGENA": "BOLIVAR", "CORDOBA": "BOLIVAR", "EL CARMEN DE BOLIVAR": "BOLIVAR",
    "GAMBOTE": "BOLIVAR", "MAGANGUE": "BOLIVAR", "MOMPOS": "BOLIVAR",
    "PASACABALLOS": "BOLIVAR", "SANTA ROSA": "BOLIVAR", "SINCERIN": "BOLIVAR",
    "TURBACO": "BOLIVAR",
    # --- BOYACA ---
    "AQUITANIA": "BOYACA", "ARCABUCO": "BOYACA", "BELEN": "BOYACA",
    "BOAVITA": "BOYACA", "BOYACA": "BOYACA", "CERINZA": "BOYACA",
    "CHIQUINQUIRA": "BOYACA", "CIENEGA": "BOYACA", "COMBITA": "BOYACA",
    "DUITAMA": "BOYACA", "EL COCUY": "BOYACA", "GARAGOA": "BOYACA",
    "GUATEQUE": "BOYACA", "LA UVITA": "BOYACA", "MIRAFLORES": "BOYACA",
    "MONIQUIRA": "BOYACA", "NOBSA": "BOYACA", "OTANCHE": "BOYACA",
    "PAIPA": "BOYACA", "PUERTO BOYACA": "BOYACA", "RAMIRIQUI": "BOYACA",
    "SABOYA": "BOYACA", "SAMACA": "BOYACA", "SAN LUIS DE GACENO": "BOYACA",
    "SANTA ROSA DE VITERBO": "BOYACA", "SOATA": "BOYACA", "SOCHA": "BOYACA",
    "SOGAMOSO": "BOYACA", "SORACA": "BOYACA", "SOTAQUIRA": "BOYACA",
    "SUSACON": "BOYACA", "SUTAMARCHAN": "BOYACA", "TIBANA": "BOYACA",
    "TIBASOSA": "BOYACA", "TOCA": "BOYACA", "TUNJA": "BOYACA",
    "TUTA": "BOYACA", "VENTAQUEMADA": "BOYACA", "VILLA DE LEYVA": "BOYACA",
    "VIRACACHA": "BOYACA",
    # --- CALDAS_DEP ---
    "AGUADAS": "CALDAS_DEP", "ANSERMA": "CALDAS_DEP", "BELALCAZAR": "CALDAS_DEP",
    "CHINCHINA": "CALDAS_DEP", "LA DORADA": "CALDAS_DEP", "MANIZALES": "CALDAS_DEP",
    "NEIRA": "CALDAS_DEP", "RIOSUCIO": "CALDAS_DEP", "RISARALDA": "CALDAS_DEP",
    "SALAMINA": "CALDAS_DEP", "SUPIA": "CALDAS_DEP", "VILLAMARIA": "CALDAS_DEP",
    "VITERBO": "CALDAS_DEP",
    # --- CAQUETA ---
    "BELEN DE LOS ANDAQUIES": "CAQUETA", "CARTAGENA DEL CHAIRA": "CAQUETA", "EL DONCELLO": "CAQUETA",
    "EL PAUJIL": "CAQUETA", "FLORENCIA": "CAQUETA", "PUERTO RICO": "CAQUETA",
    "SAN JOSE DE LA FRAGUA": "CAQUETA", "SAN VICENTE DEL CAGUAN": "CAQUETA",
    # --- CASANARE ---
    "AGUAZUL": "CASANARE", "HATO COROZAL": "CASANARE", "MANI": "CASANARE",
    "MONTERREY": "CASANARE", "NUNCHIA": "CASANARE", "PAZ DE ARIPORO": "CASANARE",
    "PORE": "CASANARE", "TAURAMENA": "CASANARE", "YOPAL": "CASANARE",
    # --- CAUCA ---
    "CAJIBIO": "CAUCA", "CALOTO": "CAUCA", "CORINTO": "CAUCA",
    "EL BORDO": "CAUCA", "GUACHENE": "CAUCA", "MERCADERES": "CAUCA",
    "MIRANDA": "CAUCA", "MORALES": "CAUCA", "PATIA": "CAUCA",
    "PIENDAMO": "CAUCA", "POPAYAN": "CAUCA", "PUERTO TEJADA": "CAUCA",
    "SANTANDER DE QUILICHAO": "CAUCA", "SILVIA": "CAUCA", "TIMBIO": "CAUCA",
    "VILLA RICA": "CAUCA",
    # --- CESAR ---
    "AGUACHICA": "CESAR", "AGUSTIN CODAZZI": "CESAR", "BECERRIL": "CESAR",
    "BOSCONIA": "CESAR", "CHIMICHAGUA": "CESAR", "CODAZZI": "CESAR",
    "EL COPEY": "CESAR", "LA JAGUA DE IBIRICO": "CESAR", "LA PAZ": "CESAR",
    "PAILITAS": "CESAR", "SAN ALBERTO": "CESAR", "VALLEDUPAR": "CESAR",
    # --- CHOCO ---
    "ISTMINA": "CHOCO", "QUIBDO": "CHOCO", "SAN FRANCISCO DE QUIBDO": "CHOCO",
    # --- CORDOBA ---
    "AYAPEL": "CORDOBA", "CERETE": "CORDOBA", "CIENAGA DE ORO": "CORDOBA",
    "LORICA": "CORDOBA", "MONTELIBANO": "CORDOBA", "MONTERIA": "CORDOBA",
    "PLANETA RICA": "CORDOBA", "SAHAGUN": "CORDOBA", "SAN ANTERO": "CORDOBA",
    "SAN BERNARDO DEL VIENTO": "CORDOBA", "SANTA CRUZ DE LORICA": "CORDOBA",
    # --- CUNDINAMARCA ---
    "AGUA DE DIOS": "CUNDINAMARCA", "ANAPOIMA": "CUNDINAMARCA", "ANOLAIMA": "CUNDINAMARCA",
    "APOSENTO ALTO": "CUNDINAMARCA", "APULO": "CUNDINAMARCA", "ARBELAEZ": "CUNDINAMARCA",
    "BOGOTA": "CUNDINAMARCA", "BOGOTA D.C.": "CUNDINAMARCA", "BOGOTA DC": "CUNDINAMARCA",
    "BOGOTÁ": "CUNDINAMARCA", "BOJACA": "CUNDINAMARCA", "CACHIPAY": "CUNDINAMARCA",
    "CAJICA": "CUNDINAMARCA", "CAQUEZA": "CUNDINAMARCA", "CHIA": "CUNDINAMARCA",
    "CHOACHI": "CUNDINAMARCA", "CHOCONTA": "CUNDINAMARCA", "COGUA": "CUNDINAMARCA",
    "COTA": "CUNDINAMARCA", "EL COLEGIO": "CUNDINAMARCA", "EL ROSAL": "CUNDINAMARCA",
    "FACATATIVA": "CUNDINAMARCA", "FUNZA": "CUNDINAMARCA", "FUSAGASUGA": "CUNDINAMARCA",
    "GACHANCIPA": "CUNDINAMARCA", "GACHETA": "CUNDINAMARCA", "GIRARDOT": "CUNDINAMARCA",
    "GUACHETA": "CUNDINAMARCA", "GUADUAS": "CUNDINAMARCA", "GUASCA": "CUNDINAMARCA",
    "LA CALERA": "CUNDINAMARCA", "LA MESA": "CUNDINAMARCA", "LA RAYA": "CUNDINAMARCA",
    "LA VEGA": "CUNDINAMARCA", "LENGUAZAQUE": "CUNDINAMARCA", "MADRID": "CUNDINAMARCA",
    "MADRONAL": "CUNDINAMARCA", "MOSQUERA": "CUNDINAMARCA", "NEMOCON": "CUNDINAMARCA",
    "NOCAIMA": "CUNDINAMARCA", "PACHO": "CUNDINAMARCA", "PARATEBUENO": "CUNDINAMARCA",
    "PUENTE DE PIEDRA": "CUNDINAMARCA", "PUERTO SALGAR": "CUNDINAMARCA", "RICAURTE": "CUNDINAMARCA",
    "SAN ANTONIO DEL TEQUENDAMA": "CUNDINAMARCA", "SAN JUAN DE RIO SECO": "CUNDINAMARCA", "SASAIMA": "CUNDINAMARCA",
    "SESQUILE": "CUNDINAMARCA", "SIBATE": "CUNDINAMARCA", "SILVANIA": "CUNDINAMARCA",
    "SIMIJACA": "CUNDINAMARCA", "SOACHA": "CUNDINAMARCA", "SOPO": "CUNDINAMARCA",
    "SUBACHOQUE": "CUNDINAMARCA", "SUESCA": "CUNDINAMARCA", "TABIO": "CUNDINAMARCA",
    "TENJO": "CUNDINAMARCA", "TOCAIMA": "CUNDINAMARCA", "TOCANCIPA": "CUNDINAMARCA",
    "TUNJUELO": "CUNDINAMARCA", "UBATE": "CUNDINAMARCA", "VILLA DE SAN DIEGO DE UBATE": "CUNDINAMARCA",
    "VILLAPINZON": "CUNDINAMARCA", "VILLETA": "CUNDINAMARCA", "VIOTA": "CUNDINAMARCA",
    "ZIPACON": "CUNDINAMARCA", "ZIPAQUIRA": "CUNDINAMARCA",
    # --- GUAINIA ---
    "INIRIDA": "GUAINIA",
    # --- GUAVIARE ---
    "SAN JOSE DEL GUAVIARE": "GUAVIARE",
    # --- HUILA ---
    "ACEVEDO": "HUILA", "AGRADO": "HUILA", "AIPE": "HUILA",
    "ALGECIRAS": "HUILA", "BARAYA": "HUILA", "CAMPOALEGRE": "HUILA",
    "GARZON": "HUILA", "GIGANTE": "HUILA", "HOBO": "HUILA",
    "LA ARGENTINA": "HUILA", "LA PLATA": "HUILA", "NEIVA": "HUILA",
    "OPORAPA": "HUILA", "PAICOL": "HUILA", "PALERMO": "HUILA",
    "PALESTINA": "HUILA", "PITAL": "HUILA", "PITALITO": "HUILA",
    "RIVERA": "HUILA", "SALADOBLANCO": "HUILA", "SAN AGUSTIN": "HUILA",
    "SAN JOSE DE ISNOS": "HUILA", "SUAZA": "HUILA", "TARQUI": "HUILA",
    "TIMANA": "HUILA", "VILLAVIEJA": "HUILA", "YAGUARA": "HUILA",
    # --- LA GUAJIRA ---
    "ALBANIA": "LA GUAJIRA", "DIBULLA": "LA GUAJIRA", "FONSECA": "LA GUAJIRA",
    "MAICAO": "LA GUAJIRA", "MANAURE": "LA GUAJIRA", "PARAGUACHON": "LA GUAJIRA",
    "RIOHACHA": "LA GUAJIRA", "URIBIA": "LA GUAJIRA", "VILLANUEVA": "LA GUAJIRA",
    # --- MAGDALENA ---
    "ALGARROBO": "MAGDALENA", "CIENAGA": "MAGDALENA", "EL BANCO": "MAGDALENA",
    "FUNDACION": "MAGDALENA", "GUACHACA": "MAGDALENA", "PLATO": "MAGDALENA",
    "SANTA MARTA": "MAGDALENA", "SITIONUEVO": "MAGDALENA",
    # --- META ---
    "ACACIAS": "META", "CABUYARO": "META", "CASTILLA LA NUEVA": "META",
    "CUMARAL": "META", "CUMARALITO": "META", "EL CASTILLO": "META",
    "EL DORADO": "META", "FUENTE DE ORO": "META", "GRANADA": "META",
    "GUAMAL": "META", "LEJANIAS": "META", "MESETAS": "META",
    "PUERTO GAITAN": "META", "PUERTO LLERAS": "META", "PUERTO LOPEZ": "META",
    "RESTREPO": "META", "SAN CARLOS DE GUAROA": "META", "SAN JUAN DE ARAMA": "META",
    "SAN MARTIN": "META", "VILLAVICENCIO": "META", "VISTA HERMOSA": "META",
    # --- NARINO ---
    "CUASPUD": "NARINO", "CUMBAL": "NARINO", "IPIALES": "NARINO",
    "PASTO": "NARINO", "SAMANIEGO": "NARINO", "SAN JUAN DE PASTO": "NARINO",
    "SOTOMAYOR": "NARINO", "TUMACO": "NARINO", "TUQUERRES": "NARINO",
    # --- NORTE DE SANTANDER ---
    "ARBOLEDAS": "NORTE DE SANTANDER", "CUCUTA": "NORTE DE SANTANDER", "LOS PATIOS": "NORTE DE SANTANDER",
    "OCANA": "NORTE DE SANTANDER", "PAMPLONA": "NORTE DE SANTANDER", "VILLA DEL ROSARIO": "NORTE DE SANTANDER",
    # --- PUTUMAYO ---
    "MOCOA": "PUTUMAYO", "ORITO": "PUTUMAYO", "PUERTO ASIS": "PUTUMAYO",
    "SIBUNDOY": "PUTUMAYO", "VALLE DEL GUAMUEZ": "PUTUMAYO",
    # --- QUINDIO ---
    "ALASKA": "QUINDIO", "ARMENIA": "QUINDIO", "CALARCA": "QUINDIO",
    "CIRCASIA": "QUINDIO", "FILANDIA": "QUINDIO", "GENOVA": "QUINDIO",
    "LA TEBAIDA": "QUINDIO", "MONTENEGRO": "QUINDIO", "QUIMBAYA": "QUINDIO",
    # --- RISARALDA ---
    "BELEN DE UMBRIA": "RISARALDA", "DOSQUEBRADAS": "RISARALDA", "GUATICA": "RISARALDA",
    "LA VIRGINIA": "RISARALDA", "PEREIRA": "RISARALDA", "QUINCHIA": "RISARALDA",
    "SANTA ROSA DE CABAL": "RISARALDA",
    # --- SANTANDER ---
    "BARRANCABERMEJA": "SANTANDER", "BRUSELAS": "SANTANDER", "BUCARAMANGA": "SANTANDER",
    "CASABE": "SANTANDER", "FLORIDABLANCA": "SANTANDER", "GIRON": "SANTANDER",
    "LEBRIJA": "SANTANDER", "LIZAMA": "SANTANDER", "MALAGA": "SANTANDER",
    "PIEDECUESTA": "SANTANDER", "PINCHOTE": "SANTANDER", "PUERTO WILCHES": "SANTANDER",
    "SABANA DE TORRES": "SANTANDER", "SAN GIL": "SANTANDER", "SANTANDER": "SANTANDER",
    "SOCORRO": "SANTANDER",
    # --- SUCRE ---
    "COROZAL": "SUCRE", "EL ROBLE": "SUCRE", "GUARANDA": "SUCRE",
    "SAMPUES": "SUCRE", "SANTIAGO DE TOLU": "SUCRE", "SINCELEJO": "SUCRE",
    "TOLUVIEJO": "SUCRE",
    # --- TOLIMA ---
    "ALVARADO": "TOLIMA", "AMBALEMA": "TOLIMA", "ATACO": "TOLIMA",
    "CAJAMARCA": "TOLIMA", "CARMEN DE APICALA": "TOLIMA", "CHAPARRAL": "TOLIMA",
    "COYAIMA": "TOLIMA", "CUNDAY": "TOLIMA", "DOLORES": "TOLIMA",
    "ESPINAL": "TOLIMA", "FALAN": "TOLIMA", "FLANDES": "TOLIMA",
    "FRESNO": "TOLIMA", "GUAMO": "TOLIMA", "GUAYABAL": "TOLIMA",
    "HONDA": "TOLIMA", "IBAGUE": "TOLIMA", "ICONONZO": "TOLIMA",
    "LERIDA": "TOLIMA", "LIBANO": "TOLIMA", "MARIQUITA": "TOLIMA",
    "MELGAR": "TOLIMA", "NATAGAIMA": "TOLIMA", "ORTEGA": "TOLIMA",
    "PALOCABILDO": "TOLIMA", "PIEDRAS": "TOLIMA", "PLANADAS": "TOLIMA",
    "PRADO": "TOLIMA", "PURIFICACION": "TOLIMA", "RIOBLANCO": "TOLIMA",
    "ROVIRA": "TOLIMA", "SALDANA": "TOLIMA", "SAN ANTONIO": "TOLIMA",
    "SANTA ISABEL": "TOLIMA", "VALLE DE SAN JUAN": "TOLIMA", "VENADILLO": "TOLIMA",
    # --- VALLE DEL CAUCA ---
    "ALCALA": "VALLE DEL CAUCA", "ANDALUCIA": "VALLE DEL CAUCA", "ANSERMANUEVO": "VALLE DEL CAUCA",
    "ARGELIA": "VALLE DEL CAUCA", "BUENAVENTURA": "VALLE DEL CAUCA", "BUGA": "VALLE DEL CAUCA",
    "BUGALAGRANDE": "VALLE DEL CAUCA", "CAICEDONIA": "VALLE DEL CAUCA", "CALI": "VALLE DEL CAUCA",
    "CANDELARIA": "VALLE DEL CAUCA", "CARTAGO": "VALLE DEL CAUCA", "DAGUA": "VALLE DEL CAUCA",
    "DARIEN": "VALLE DEL CAUCA", "EL CERRITO": "VALLE DEL CAUCA", "FLORIDA": "VALLE DEL CAUCA",
    "GINEBRA": "VALLE DEL CAUCA", "GUACARI": "VALLE DEL CAUCA", "GUADALAJARA DE BUGA": "VALLE DEL CAUCA",
    "JAMUNDI": "VALLE DEL CAUCA", "LA CUMBRE": "VALLE DEL CAUCA", "LA PAILA": "VALLE DEL CAUCA",
    "LA VICTORIA": "VALLE DEL CAUCA", "OBANDO": "VALLE DEL CAUCA", "PALMIRA": "VALLE DEL CAUCA",
    "PRADERA": "VALLE DEL CAUCA", "ROLDANILLO": "VALLE DEL CAUCA", "SANTIAGO DE CALI": "VALLE DEL CAUCA",
    "SEVILLA": "VALLE DEL CAUCA", "TULUA": "VALLE DEL CAUCA", "VIJES": "VALLE DEL CAUCA",
    "YUMBO": "VALLE DEL CAUCA", "ZARZAL": "VALLE DEL CAUCA",
    # --- VICHADA ---
    "ACEITICO": "VICHADA",
}

# ============================================================
# 🌎 DEPARTAMENTO -> REGIÓN / ZONA
# ============================================================
DEPTO_REGION = {
    "ATLANTICO": "COSTA", "BOLIVAR": "COSTA", "MAGDALENA": "COSTA",
    "CORDOBA": "COSTA", "SUCRE": "COSTA", "CESAR": "COSTA", "LA GUAJIRA": "COSTA",
    "ANTIOQUIA": "ANTIOQUIA",
    "CUNDINAMARCA": "CUNDINAMARCA/BOGOTA",
    "VALLE DEL CAUCA": "VALLE", "CAUCA": "VALLE",
    "SANTANDER": "SANTANDERES", "NORTE DE SANTANDER": "SANTANDERES",
    "BOYACA": "BOYACA",
    "RISARALDA": "EJE/OTROS", "CALDAS_DEP": "EJE/OTROS", "QUINDIO": "EJE/OTROS",
    "TOLIMA": "EJE/OTROS", "HUILA": "EJE/OTROS", "META": "EJE/OTROS",
    "NARINO": "EJE/OTROS", "OTRO": "EJE/OTROS",
}


def depto_de(ciudad):
    return CIUDAD_DEPTO.get(ciudad, "OTRO")


def region_de(depto):
    return DEPTO_REGION.get(depto, "EJE/OTROS")


# Nombre del proveedor que identifica flota PROPIA
PROVEEDOR_PROPIO = "TRACTOCAR LOGISTICS"


def tipo_flota(proveedor):
    """Propia si el afiliado/proveedor es Tractocar; en cualquier otro caso Tercero.
    (Todas las operaciones traen afiliado o proveedor, así que ya no hay 'sin dato'.)"""
    p = limpio(proveedor).upper()
    if PROVEEDOR_PROPIO in p:
        return "PROPIA"
    return "TERCERO"


def es_alto_cubicaje(cuenta_contable):
    """Alto cubicaje si la cuenta contable contiene 'AC' como segmento (ej. CGN-AC-0018)."""
    c = limpio(cuenta_contable).upper()
    if not c:
        return False
    # 'AC' rodeado de guiones o como token, evita falsos positivos dentro de palabras
    import re as _re
    return bool(_re.search(r"(^|[-_ ])AC([-_ ]|$)", c))


def limpio(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in {"-", "N/D", "nan", "NaT"} else s


def quitar_prefijo(v):
    """TCL.LZO222 -> LZO222 ; TCL.2261640 -> 2261640"""
    s = limpio(v)
    return s.split(".", 1)[1] if s.startswith("TCL.") else s


def a_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = limpio(v)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


def main():
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("❌ Falta openpyxl. Instálala con:  pip install openpyxl")
        sys.exit(1)

    # ---- acumuladores ----
    flota = {}
    envios_globales = set()
    clientes_finales = set()
    operaciones_vistas = set()
    fecha_min, fecha_max = None, None
    filas_totales = 0
    resumen_ops = {}

    carpetas_existentes = {op: ruta for op, ruta in CARPETAS_OPERACION.items() if os.path.isdir(ruta)}
    if not carpetas_existentes:
        print("❌ No encontré ninguna carpeta de operaciones. Revisa RUTA_BASE y CARPETAS_OPERACION.")
        for op, ruta in CARPETAS_OPERACION.items():
            print(f"   {op}: {ruta}")
        sys.exit(1)

    for op, carpeta in carpetas_existentes.items():
        mapeo = MAPEO_OPERACION[op]
        cols = mapeo["cols"]
        fila_enc = mapeo["fila_encabezado"]

        archivos = sorted(
            f for f in glob.glob(os.path.join(carpeta, "*.xls*"))
            if not os.path.basename(f).startswith("~$")
        )
        if not archivos:
            print(f"📂 {op}: carpeta sin archivos .xlsx — la salto ({carpeta})")
            continue

        print(f"\n📂 Operación {op}: {len(archivos)} archivo(s) en {carpeta}")
        resumen_ops[op] = {"archivos": len(archivos), "filas": 0, "viajes": set()}

        for ruta_archivo in archivos:
            nombre = os.path.basename(ruta_archivo)
            try:
                wb = load_workbook(ruta_archivo, read_only=True, data_only=True)
            except Exception as e:
                print(f"   ⚠️ {nombre}: no se pudo abrir ({e}) — lo salto")
                continue

            # tomar la hoja con datos y leer su encabezado en la fila indicada
            ws = wb[wb.sheetnames[0]]
            todas = list(ws.iter_rows(values_only=True))
            if len(todas) < fila_enc + 1:
                print(f"   ⚠️ {nombre}: sin filas suficientes — lo salto")
                continue
            encabezados = list(todas[fila_enc - 1])
            idx = {n: i for i, n in enumerate(encabezados) if n}

            col_placa = cols["placa"]
            col_envio = cols["envio"]
            if col_placa not in idx or col_envio not in idx:
                print(f"   ⚠️ {nombre}: no tiene columnas '{col_placa}'/'{col_envio}' — lo salto")
                continue

            def valor(fila, nombre_col):
                i = idx.get(nombre_col)
                return fila[i] if i is not None and i < len(fila) else None

            n_filas = 0
            for fila in todas[fila_enc:]:
                placa = quitar_prefijo(valor(fila, col_placa))
                envio = quitar_prefijo(valor(fila, col_envio))
                if not placa or not envio:
                    continue
                n_filas += 1
                filas_totales += 1
                resumen_ops[op]["filas"] += 1

                fecha = a_fecha(valor(fila, cols["fecha"]))
                origen = limpio(valor(fila, cols["origen"])).upper()
                destino = limpio(valor(fila, cols["destino"])).upper()
                cliente = unificar_cliente(limpio(valor(fila, cols["cliente"])).upper())
                tip = limpio(valor(fila, cols["tipologia"])).upper()
                # operación: de la columna si existe, si no el nombre de la carpeta
                op_fila = limpio(valor(fila, cols.get("operacion", ""))) .upper() or op
                operaciones_vistas.add(op_fila)
                # tipo de flota (propia/tercero) y alto cubicaje
                flota_tipo = tipo_flota(valor(fila, cols.get("proveedor", "")))
                alto_cub = es_alto_cubicaje(valor(fila, cols.get("contable", "")))
                if cliente:
                    clientes_finales.add(cliente)

                if fecha:
                    fecha_min = fecha if not fecha_min or fecha < fecha_min else fecha_min
                    fecha_max = fecha if not fecha_max or fecha > fecha_max else fecha_max

                f = flota.setdefault(placa, {"envios": set(), "rutas": {}, "ult": None,
                                             "pri": None, "tipologia": {}, "ops": {},
                                             "flota": {}, "ac": 0, "noac": 0})
                f["envios"].add(envio)
                envios_globales.add(envio)
                resumen_ops[op]["viajes"].add(envio)
                f["ops"][op_fila] = f["ops"].get(op_fila, 0) + 1
                if flota_tipo:
                    f["flota"][flota_tipo] = f["flota"].get(flota_tipo, 0) + 1
                if alto_cub:
                    f["ac"] += 1
                else:
                    f["noac"] += 1
                if tip:
                    f["tipologia"][tip] = f["tipologia"].get(tip, 0) + 1
                if fecha:
                    f["ult"] = fecha if not f["ult"] or fecha > f["ult"] else f["ult"]
                    f["pri"] = fecha if not f["pri"] or fecha < f["pri"] else f["pri"]

                # la ruta incluye la operación para poder filtrar por corredor + operación
                clave = (origen, destino, cliente, op_fila)
                r = f["rutas"].setdefault(clave, {"envios": set(), "ult": None})
                r["envios"].add(envio)
                if fecha:
                    r["ult"] = fecha if not r["ult"] or fecha > r["ult"] else r["ult"]

            print(f"   ✅ {nombre}: {n_filas:,} filas")

    if not flota:
        print("❌ No se pudo leer ningún dato de ninguna operación.")
        sys.exit(1)

    # ---- consolidar ----
    hoy = date.today()
    placas_out = []
    for placa, f in flota.items():
        viajes = len(f["envios"])
        ult, pri = f["ult"], f["pri"]
        dias = (hoy - ult).days if ult else 9999
        if dias <= DIAS_FIDELIZADA:
            estado = "FIDELIZADA"
        elif dias <= DIAS_EVENTUAL:
            estado = "EVENTUAL"
        else:
            estado = "POR RECUPERAR"
        meses = max(((hoy - pri).days / 30.44), 1.0) if pri else 1.0
        tip = max(f["tipologia"], key=f["tipologia"].get) if f["tipologia"] else ""

        rutas = []
        for (o, d, c, op_r), r in f["rutas"].items():
            od, dd = depto_de(o), depto_de(d)
            rutas.append({
                "o": o, "d": d, "c": c, "op": op_r,
                "od": od, "dd": dd,           # departamento origen / destino
                "orr": region_de(od), "dr": region_de(dd),  # región origen / destino
                "n": len(r["envios"]),
                "f": r["ult"].strftime("%Y-%m-%d") if r["ult"] else "",
            })
        rutas.sort(key=lambda x: -x["n"])

        # operación dominante de la placa y lista de todas sus operaciones
        ops_placa = sorted(f["ops"], key=f["ops"].get, reverse=True)
        op_top = ops_placa[0] if ops_placa else ""

        # tipo de flota dominante (PROPIA/TERCERO/"" si sin dato)
        flota_tipo = ""
        if f["flota"]:
            flota_tipo = max(f["flota"], key=f["flota"].get)
        # alto cubicaje: "SI" si la placa tiene al menos un viaje de alto cubicaje
        ac = "SI" if f["ac"] > 0 else "NO"

        placas_out.append({
            "placa": placa,
            "tip": tip,
            "viajes": viajes,
            "ultima": ult.strftime("%Y-%m-%d") if ult else "",
            "primera": pri.strftime("%Y-%m-%d") if pri else "",
            "dias": dias,
            "estado": estado,
            "meses": round(meses, 2),
            "rot": round(viajes / meses, 2),
            "opTop": op_top,
            "ops": ops_placa,
            "flotaTipo": flota_tipo,
            "ac": ac,
            "rutas": rutas,
        })

    placas_out.sort(key=lambda x: -x["viajes"])

    paquete = {
        "actualizado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "desde": fecha_min.strftime("%Y-%m-%d") if fecha_min else "",
        "hasta": fecha_max.strftime("%Y-%m-%d") if fecha_max else "",
        "filas": filas_totales,
        "totalViajes": len(envios_globales),
        "operaciones": sorted(operaciones_vistas),
        "umbralFidelizada": DIAS_FIDELIZADA,
        "umbralEventual": DIAS_EVENTUAL,
        "placas": placas_out,
    }

    with open(SALIDA, "w", encoding="utf-8") as f:
        f.write("// Generado por convertir_viajes.py — NO editar a mano.\n")
        f.write("window.DATOS_VIAJES = ")
        json.dump(paquete, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")

    # Lista de manifiestos (envíos) válidos del periodo de viajes, para que
    # convertir_pnc.py solo cuente PNC cuyo manifiesto exista aquí.
    ruta_manif = os.path.join(os.path.dirname(os.path.abspath(__file__)), "manifiestos_validos.json")
    with open(ruta_manif, "w", encoding="utf-8") as f:
        json.dump(sorted(envios_globales), f, ensure_ascii=False, separators=(",", ":"))

    peso = os.path.getsize(SALIDA) / 1024 / 1024
    fid = sum(1 for p in placas_out if p["estado"] == "FIDELIZADA")
    eve = sum(1 for p in placas_out if p["estado"] == "EVENTUAL")
    rec = sum(1 for p in placas_out if p["estado"] == "POR RECUPERAR")
    print(f"\n✅ data_viajes.js generado — {peso:.1f} MB")
    print(f"   {len(placas_out):,} placas · {len(envios_globales):,} viajes · {fecha_min} → {fecha_max}")
    print(f"   Fidelizadas: {fid:,} · Eventuales: {eve:,} · Por recuperar: {rec:,}")

    # Desglose por operación
    print("\n🚚 Viajes por operación:")
    for op in sorted(resumen_ops):
        r = resumen_ops[op]
        print(f"   {op}: {len(r['viajes']):,} viajes · {r['filas']:,} filas · {r['archivos']} archivo(s)")

    # Lista de clientes finales (ya unificados) para que detectes variantes pendientes.
    print(f"\n📋 {len(clientes_finales)} clientes (ya unificados). Revisa si hay variantes que falte agrupar:")
    for c in sorted(clientes_finales):
        print("   - " + c)

    # Ciudades sin departamento asignado (para que las agregues a CIUDAD_DEPTO)
    ciudades = set()
    for p in placas_out:
        for r in p["rutas"]:
            if r["od"] == "OTRO" and r["o"]:
                ciudades.add(r["o"])
            if r["dd"] == "OTRO" and r["d"]:
                ciudades.add(r["d"])
    if ciudades:
        print(f"\n⚠️ {len(ciudades)} ciudad(es) SIN departamento asignado (quedaron como OTRO/EJE-OTROS).")
        print("   Agrégalas a la tabla CIUDAD_DEPTO en convertir_viajes.py:")
        for c in sorted(ciudades):
            print("   - " + c)
    else:
        print("\n✅ Todas las ciudades tienen departamento asignado.")


if __name__ == "__main__":
    main()
