import re
import os
from pathlib import Path
import pandas as pd
import numpy as np

# -----------------------
# 1) Ajustes básicos
# -----------------------
# Carpeta donde pondrás tus Excels crudos
INPUT_DIR = Path("input")
# Archivo de salida
OUTPUT_XLSX = Path("output/Resultados_Skinner.xlsx")

# Si quieres cambiar la regla de "nueva sesión", mira la función `calcular_eventos_y_momento()`

# -----------------------
# 2) Funciones auxiliares
# -----------------------
def parsear_nombre_archivo(nombre_sin_ext):
    """
    Extrae Grupo, M# y Día desde el nombre del archivo.
    Formatos tolerantes, por ejemplo:
      'SSM21 Día1', 'SS M21 Dia 1', 'GrupoX M7 Día12', etc.
    Regresa (grupo, raton_id, dia_int)
    """
    patron = re.compile(r"^\s*([A-Za-zÁÉÍÓÚÜÑ0-9]+)\s*(M\d+).*(?:D[ií]a)\s*(\d+)", re.IGNORECASE)
    m = patron.search(nombre_sin_ext)
    if not m:
        return "", "", None

    grupo_crudo = m.group(1)   # p.ej. 'SSM21' o 'SS'
    raton_id = m.group(2)      # 'M21'
    dia = int(m.group(3))      # 1

    # Si el grupo son solo las letras antes de 'M', recórtalo
    p = grupo_crudo.upper().find("M")
    if p > 0:
        grupo = grupo_crudo[:p]
    else:
        grupo = grupo_crudo

    return grupo, raton_id, dia


def leer_datos_crudos(path_excel):
    """
    Lee las primeras 5 columnas de la primera hoja con datos, de forma tolerante.
    - Soporta .xlsx/.xlsm y .csv
    - Si no hay encabezados en la fila 2, igualmente fuerza los nombres estándar
    """
    suffix = str(path_excel.suffix).lower()

    if suffix in [".xlsx", ".xlsm"]:
        # 1) Intento estándar: encabezados en la fila 2 (header=1)
        try:
            df = pd.read_excel(path_excel, header=1, engine="openpyxl")
        except Exception:
            # 2) Fallback: sin encabezados
            df = pd.read_excel(path_excel, header=None, engine="openpyxl")

    elif suffix == ".csv":
        # Algunos exportan CSV: intentamos con encabezados en fila 2 y fallback
        try:
            df = pd.read_csv(path_excel, header=1, encoding="utf-8", sep=",")
        except Exception:
            df = pd.read_csv(path_excel, header=None, encoding="utf-8", sep=",")
    else:
        raise ValueError(f"Formato no soportado: {suffix}. Usa .xlsx/.xlsm (o .csv).")

    # Nos quedamos con las PRIMERAS 5 columnas que existan
    if df.shape[1] < 5:
        # Puede que haya columnas vacías a la derecha: completamos con NaN
        faltan = 5 - df.shape[1]
        for _ in range(faltan):
            df[f"_col_dummy_{_}"] = np.nan

    df = df.iloc[:, :5]  # primeras 5

    # Renombramos a los nombres estándar
    df.columns = [
        "Datos crudos columna 1",
        "Datos crudos columna 2",
        "Datos crudos columna 3",
        "Datos crudos columna 4",
        "Datos crudos columna 5",
    ]

    # Quitamos filas totalmente vacías
    df = df.dropna(how="all")

    return df



def calcular_eventos_y_momento(df):
    """
    Calcula #Evento por sesión y los momentos (seg/horas) según tus reglas:
      - Nueva sesión si el Tiempo interensayo de la fila anterior es 0 o NaN.
      - Para i==0: Momento(seg) = Z(i).
      - Si #Evento(i) == 1 (nueva sesión): Momento(seg) = Z(i-1).
      - Si #Evento(i) > 1: Momento(seg) = AA(i-1) + Y(i-1) + Z(i).
    """
    if df.empty:
        return df

    df = df.copy().reset_index(drop=True)

    # Asegura numérico en columnas usadas
    ti = pd.to_numeric(df["Tiempo interensayo"], errors="coerce").fillna(0)  # Y
    z  = pd.to_numeric(df["Tiempo respuesta correspondiente al evento"], errors="coerce").fillna(0)  # Z

    n = len(df)
    eventos = [0] * n
    momento = [0.0] * n

    for i in range(n):
        if i == 0:
            eventos[i] = 1
            momento[i] = float(z.iat[i])
        else:
            nueva_sesion = (ti.iat_)

def calcular_eventos_y_momento(df):
    """
    Calcula #Evento por sesión (reinicia si interensayo previo = 0 o NaN)
    y los momentos en segundos y horas como en tus reglas.
    """
    if df.empty:
        return df

    df = df.copy().reset_index(drop=True)

    # Convierte las columnas usadas a numérico (por seguridad)
    df["Tiempo interensayo"] = pd.to_numeric(df["Tiempo interensayo"], errors="coerce").fillna(0)
    df["Tiempo respuesta correspondiente al evento"] = pd.to_numeric(
        df["Tiempo respuesta correspondiente al evento"], errors="coerce"
    ).fillna(0)

    ti = df["Tiempo interensayo"]  # Y
    z = df["Tiempo respuesta correspondiente al evento"]  # Z
    n = len(df)

    # Inicializa listas del mismo largo que el DF
    eventos = [0] * n
    momento = [0.0] * n

    # Recorre todas las filas
    for i in range(n):
        if i == 0:
            # Primera fila del archivo → evento 1
            eventos[i] = 1
            momento[i] = z.iat[i]
        else:
            inter_prev = ti.iat[i - 1]
            nueva_sesion = (inter_prev == 0)

            if nueva_sesion:
                eventos[i] = 1
                momento[i] = z.iat[i - 1]
            else:
                eventos[i] = eventos[i - 1] + 1
                momento[i] = momento[i - 1] + inter_prev + z.iat[i]

    # Agrega las columnas calculadas
    df["#Evento"] = eventos
    df["Momento (seg) considerando sesiones como eventos distintos"] = momento

    # Calcula en horas
    df["Momento (horas) considerando sesiones como eventos distintos"] = (
        pd.to_numeric(df["Momento (seg) considerando sesiones como eventos distintos"], errors="coerce").fillna(0)
        / 3600.0
    )
    df["Redondeo de horas (IGNORAR)2"] = df["Momento (horas) considerando sesiones como eventos distintos"].round(1)

    return df


def calcular_columnas_derivadas(df):
    """
    Calcula TODAS las columnas derivadas a partir de las 5 crudas.
    Asegura que las crudas sean numéricas (soporta comas decimales).
    """
    # 1) Asegura numérico en crudas
    for c in [
        "Datos crudos columna 1",
        "Datos crudos columna 2",
        "Datos crudos columna 3",
        "Datos crudos columna 4",
        "Datos crudos columna 5",
    ]:
        if df[c].dtype == "object":
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", ".", regex=False), errors="coerce")

    # Alias (más legible)
    raw1 = df["Datos crudos columna 1"]  # para Izq/Der
    raw2 = df["Datos crudos columna 2"]  # Actividad (0,1,2)
    raw3 = df["Datos crudos columna 3"]  # TR (D)
    raw4 = df["Datos crudos columna 4"]  # Interensayo (E)

    # 2) Columnas derivadas
    # Izq/Der := SI(raw1=0,2,1)
    df["Izq/Der"] = np.where(raw1.fillna(0) == 0, 2, 1).astype(int)

    # Actividad := raw2
    df["Actividad"] = raw2

    # Izquierdo / Derecho
    df["Izquierdo"] = np.where(df["Izq/Der"] == 1, 1, 0).astype(int)
    df["Derecho"]  = np.where(df["Izq/Der"] == 2, 1, 0).astype(int)

    # ACIERTOS / ERRORES / NULOS (Actividad 1 / 2 / 0)
    df["ACIERTOS"] = np.where(df["Actividad"] == 1, 1, 0).astype(int)
    df["ERRORES"]  = np.where(df["Actividad"] == 2, 1, 0).astype(int)
    df["NULOS"]    = np.where(df["Actividad"] == 0, 1, 0).astype(int)

    # Desgloses por lado
    df["ACIERTOS izquierdo"] = np.where(df["ACIERTOS"] == 1, df["Izquierdo"], 0).astype(int)
    df["ACIERTOS derecho"]   = np.where(df["ACIERTOS"] == 1, df["Derecho"], 0).astype(int)
    df["error izquierdo"]    = np.where(df["ERRORES"]  == 1, df["Izquierdo"], 0).astype(int)
    df["error derecho"]      = np.where(df["ERRORES"]  == 1, df["Derecho"], 0).astype(int)
    df["nulo izquierdo"]     = np.where(df["NULOS"]    == 1, df["Izquierdo"], 0).astype(int)
    df["nulo derecho"]       = np.where(df["NULOS"]    == 1, df["Derecho"], 0).astype(int)

    # TR / TRA / Interensayo / Respuesta correspondiente
    df["Tiempo de respuesta(TR)"] = raw3
    df["TR aciertos (TRA)"] = np.where(df["ACIERTOS"] == 1, df["Tiempo de respuesta(TR)"], 0)
    df["Tiempo interensayo"] = raw4
    df["Tiempo respuesta correspondiente al evento"] = raw3

    return df


def procesar_archivo(path_excel):
    """
    Procesa un archivo individual:
      - lee crudos
      - agrega columnas derivadas
      - calcula eventos y momento
      - añade Grupo, Ratón, Día
    Devuelve (df_resultado, nombre_base)
    """
    base = path_excel.stem  # nombre sin extensión
    grupo, raton_id, dia = parsear_nombre_archivo(base)

    df = leer_datos_crudos(path_excel)

    # Guarda columnas crudas tal cual al inicio (para que salgan en el resultado)
    df_proc = df.copy()

    # Derivadas
    df_proc = calcular_columnas_derivadas(df_proc)

    # Eventos + Momento
    df_proc = calcular_eventos_y_momento(df_proc)

    # Metadatos del archivo
    df_proc.insert(df_proc.columns.tolist().index("Izq/Der"), "Día", dia)
    df_proc.insert(df_proc.columns.tolist().index("Día"), "Identificador ratón", raton_id)
    df_proc.insert(df_proc.columns.tolist().index("Identificador ratón"), "Grupo experimental", grupo)
    df_proc["Archivo_origen"] = base

    # Orden de presentación similar al Excel:
    columnas_orden = [
        "Datos crudos columna 1",
        "Datos crudos columna 2",
        "Datos crudos columna 3",
        "Datos crudos columna 4",
        "Datos crudos columna 5",
        "Grupo experimental",
        "Identificador ratón",
        "Día",
        "#Evento",
        "Izq/Der",
        "Actividad",
        "Izquierdo",
        "Derecho",
        "ACIERTOS",
        "ERRORES",
        "NULOS",
        "ACIERTOS izquierdo",
        "ACIERTOS derecho",
        "error izquierdo",
        "error derecho",
        "nulo izquierdo",
        "nulo derecho",
        "Tiempo de respuesta(TR)",
        "TR aciertos (TRA)",
        "Tiempo interensayo",
        "Tiempo respuesta correspondiente al evento",
        "Momento (seg) considerando sesiones como eventos distintos",
        "Momento (horas) considerando sesiones como eventos distintos",
        "Redondeo de horas (IGNORAR)2",
        "Archivo_origen",
    ]
    # Asegura orden, ignorando si falta alguna por error de origen
    columnas_orden = [c for c in columnas_orden if c in df_proc.columns]
    df_proc = df_proc[columnas_orden]

    return df_proc, base


def procesar_todos_los_archivos():
    """
    Recorre INPUT_DIR, procesa todos los .xlsx/.xlsm/.xls y genera el Excel de salida
    con:
      - Hoja 'Consolidado' (todas las filas)
      - Una hoja por archivo procesado (opcional: se puede quitar)
    """
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"No existe la carpeta {INPUT_DIR}. Créala y coloca tus archivos dentro.")

    archivos = sorted([p for p in INPUT_DIR.iterdir() if p.suffix.lower() in [".xlsx", ".xlsm", ".xls"]])
    if not archivos:
        raise FileNotFoundError(f"No encontré archivos Excel en {INPUT_DIR}.")

    resultados = []
    hojas_individuales = {}

    for p in archivos:
        df_proc, base = procesar_archivo(p)
        resultados.append(df_proc)
        hojas_individuales[base[:31]] = df_proc  # Excel limita nombre de hoja a 31 caracteres

    # Consolidado
    consolidado = pd.concat(resultados, ignore_index=True)

    # Orden final sugerido
    if set(["Grupo experimental", "Identificador ratón", "Día"]).issubset(consolidado.columns):
        consolidado.sort_values(by=["Grupo experimental", "Identificador ratón", "Día", "Archivo_origen"], inplace=True)

    # Escribir salida
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        consolidado.to_excel(writer, index=False, sheet_name="Consolidado")
        # Hojas por archivo (si no las quieres, comenta este bloque)
        for nombre_hoja, df in hojas_individuales.items():
            df.to_excel(writer, index=False, sheet_name=nombre_hoja)

    return OUTPUT_XLSX


if __name__ == "__main__":
    out = procesar_todos_los_archivos()
    print(f"Listo. Archivo generado en: {out.resolve()}")
