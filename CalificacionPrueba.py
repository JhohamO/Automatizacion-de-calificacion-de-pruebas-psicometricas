import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib

#EL backend predterminado
matplotlib.use("TkAgg")

"""Código para califcar automáticamente la prueba MSCEIT"""

#Leer el archivo excel
res_brutos = pd.read_excel("MSCEIT- TEST DE INTELIGENCIA EMOCIONAL(1-73).xlsx")

#Extraigo las identificaciones
identificaciones = res_brutos["Documento de Identificación"]

#Borro todo lo que no sea respuesta
res_brutos = res_brutos.drop(res_brutos.columns[: 7], axis=1)

# Como las preguntas son obligatorias, cada columna tendrá el mismo número de valores
frecuencia_total = len(res_brutos)

# Iterar sobre todas las columnas y calcular la frecuencia relativa para cada respuesta
for columna in res_brutos.columns:

    #Hallo la fecuencia relativa de cada elemento que surge
    cuentas = res_brutos[columna].value_counts() / frecuencia_total

    #Voy a cambiar el elemento por su frecuencia relativa
    for indice, valor in res_brutos[columna].items():
        frecuencia_relativa = cuentas[valor]
        res_brutos.at[indice, columna] = frecuencia_relativa

#Fines de las secciones
seccion_A = res_brutos.columns.get_loc("Entusiasmo4") + 1
seccion_B = res_brutos.columns.get_loc("Enojo y desafío") + 1
seccion_C = res_brutos.columns.get_loc("Una mujer amaba a una persona y luego se sintió segura. ¿Qué pudo ocurrirle?") + 1
seccion_D = res_brutos.columns.get_loc("Acción 4: Juró que nunca volvería a conducir por esa autovía.") + 1
seccion_E = res_brutos.columns.get_loc("Asco10") + 1
seccion_F = res_brutos.columns.get_loc("Calmado") + 1
seccion_G = res_brutos.columns.get_loc("Tristeza y satisfacción son a veces parte del sentimiento de ____________________.") + 1
seccion_H = res_brutos.columns.get_loc("Respuesta 3: Esa noche Lisa compartió sus sentimientos con su marido. Poco después, decidió que la familia debería pasar más tiempo junta los fines de semana y hacer más actividades familiares par...") + 1

#Voy a sacar las calificaciones en bruto de las secciones o tareas
calificaciones_brutas = {}
calificaciones_brutas["Caras"] = res_brutos.iloc[:, :seccion_A].sum(axis=1)
calificaciones_brutas["Facilitacion"] = res_brutos.iloc[:, seccion_A: seccion_B].sum(axis=1)
calificaciones_brutas["Cambios"] = res_brutos.iloc[:, seccion_B:seccion_C].sum(axis=1)
calificaciones_brutas["Manejo emocional"] = res_brutos.iloc[:, seccion_C:seccion_D].sum(axis=1)
calificaciones_brutas["Dibujos"] = res_brutos.iloc[:, seccion_D:seccion_E].sum(axis=1)
calificaciones_brutas["Sensaciones"] = res_brutos.iloc[:, seccion_E:seccion_F].sum(axis=1)
calificaciones_brutas["Combinaciones"] = res_brutos.iloc[:, seccion_F:seccion_G].sum(axis=1)
calificaciones_brutas["Relaciones emocionales"] = res_brutos.iloc[:, seccion_G:seccion_H].sum(axis=1)

#Borro lo que no necesito más
del seccion_A, seccion_B, seccion_C, seccion_D, seccion_E, seccion_F, seccion_G, seccion_H

#Voy a pasar a las calificaciones en bruto a unas que se rijan a una distribución normal
cal_tareas_normal = {}

#Recorro el diccionario original
for tarea, vector in calificaciones_brutas.items():

    z_scores = (vector - vector.mean()) / vector.std()

    cal_tareas_normal[tarea] = z_scores * 15 + 100

#Borro lo no necesario
del z_scores, tarea, vector

#Voy a crear el diccionario de calificaciones brutas de las ramas
ramas_brutas = {}

#Sumo las tareas correspondientes que componen cada rama
ramas_brutas["CIEP"] = calificaciones_brutas["Caras"] + calificaciones_brutas["Dibujos"]
ramas_brutas["CIEF"] = calificaciones_brutas["Facilitacion"] + calificaciones_brutas["Sensaciones"]
ramas_brutas["CIEC"] = calificaciones_brutas["Cambios"] + calificaciones_brutas["Combinaciones"]
ramas_brutas["CIEM"] = calificaciones_brutas["Manejo emocional"] + calificaciones_brutas["Relaciones emocionales"]

#Voy a crear el diccionario para extraer la calificación normal de cada rama
ramas_normales = {}

#Recorro el diccionario original
for rama, vector in ramas_brutas.items():
    z_scores = (vector - vector.mean()) / vector.std()
    ramas_normales[rama] = z_scores * 15 + 100

#Borro lo no necesario
del rama, vector, z_scores

#Voy a crear el diccionario de calificaciones brutas de las áreas
areas_brutas = {}

#Lleno cada área con sus correspondientes ramas
areas_brutas["CIEX"] = ramas_brutas["CIEP"] + ramas_brutas["CIEF"]
areas_brutas["CIES"] = ramas_brutas["CIEC"] + ramas_brutas["CIEM"]

#Creo el diccionario con la calificicación normalizada de las areas
areas_normal = {}

#Lleno el diccionario anterior
for area, vector in areas_brutas.items():
    z_scores = (vector - vector.mean()) / vector.std()
    areas_normal[area] = z_scores * 15 + 100

del area, vector, z_scores

#Creo el diccionario vacío para la lista de
CIE = areas_brutas["CIEX"] + areas_brutas["CIES"]
z_scores = (CIE - CIE.mean()) / CIE.std()
CIE_normal = z_scores * 15 + 100

#Creo una matriz donde voy a guardar todos los datos
datos_excel = np.zeros([len(res_brutos), 7])

datos_excel[:, 0] = CIE_normal
datos_excel[:, 1] = areas_normal["CIEX"]
datos_excel[:, 2] = areas_normal["CIES"]
datos_excel[:, 3] = ramas_normales["CIEP"]
datos_excel[:, 4] = ramas_normales["CIEF"]
datos_excel[:, 5] = ramas_normales["CIEC"]
datos_excel[:, 6] = ramas_normales["CIEM"]

#Creo las columanas del excel
columnas = ["CIE", "CIEX", "CIES", "CIEP", "CIEF", "CIEC", "CIEM"]

excel_MSCEIT = pd.DataFrame(datos_excel, columns=columnas, index=identificaciones)
excel_MSCEIT.to_excel("ResultadosMSCEIT.xlsx")

##

#Voy a recorrer los documentos
for ind, dni in identificaciones.items():
    documento_graficar = dni
# #Meto por parámetro el documento a graficar
# documento_graficar = 1034320576

# #Busco el índice del documento en las identificaciones
# indice_id = (identificaciones == documento_graficar).idxmax()

    indice_id = ind

    #Extraigo calificaciones del sujeto
    calificaciones = [CIE_normal[indice_id], areas_normal["CIEX"][indice_id], areas_normal["CIES"][indice_id],
                      ramas_normales["CIEP"][indice_id], ramas_normales["CIEF"][indice_id], ramas_normales["CIEC"][indice_id],
                      ramas_normales["CIEM"][indice_id]]

    # Crear un gráfico de barras para la hoja de perfil centrada en 100
    plt.bar(columnas, calificaciones, color='skyblue')

    # Configurar el título y etiquetas
    plt.title('Hoja de Perfil')
    plt.xlabel('Ítems')
    plt.ylabel('Calificaciones')

    # Centrar la gráfica en 100
    plt.axhline(y=100, color='red', linestyle='--', linewidth=2)

    # Agregar etiquetas a cada barra
    for i, valor in enumerate(calificaciones):
        plt.text(i, valor + 0.5, str(int(valor)), ha='center', va='bottom', color='black')

    plt.savefig("Graficas_MSCEIT\\" + str(dni) + ".png")

    plt.clf()











