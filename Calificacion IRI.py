import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib

#Importo el backend default para gráficas estáticas
matplotlib.use("TkAgg")

#Leo el archivo excel
respuestas_iri = pd.read_excel("IRI PRÁCTICA MNS  (respuestas).xlsx")

"""Invierto las columnas que hay que invertir"""

#Saco las columas de las preguntas
preguntas = respuestas_iri.iloc[:, 2:]

#Especifico las preguntas a invertir. La columna correspondiente a la pregunta es una menos
columnas_invertir = [3, 15, 7, 12, 4, 13, 14, 18, 19]

#Recorro las columnas
for ncol, col in enumerate(preguntas.columns):

    #Invierto si es necesario
    if ncol + 1 in columnas_invertir:
        preguntas[col] = 6 - preguntas[col]

"""Selecciono las preguntas correspondientes a cada dimension"""

#Preguntas correspondientes a cada dimensión
dimension_tp = np.array([3, 8, 11, 15, 21, 25, 28]) - 1
dimension_f = np.array([1, 5, 7, 12, 16, 23, 26]) - 1
dimension_ec = np.array([2, 4, 9, 13, 14, 18, 20, 22]) - 1
dimension_pd = np.array([6, 10, 17, 19, 24, 27]) - 1

#Agrego las columnas de los particpantes
preguntas["Sujeto"] = respuestas_iri['Código de participante']

#Agrego las columnas de calificación al df de preguntas, que es solo sumando por dimensiones
preguntas["TP"] = preguntas.iloc[:, dimension_tp].sum(axis=1)
preguntas["F"] = preguntas.iloc[:, dimension_f].sum(axis=1)
preguntas["EC"] = preguntas.iloc[:, dimension_ec].sum(axis=1)
preguntas["PD"] = preguntas.iloc[:, dimension_ec].sum(axis=1)

#Saco los puntos de cambio
limite_medio = 9
limite_alto = 19

#Recorro los sujetos para graficar su info
for sujetos in range(len(preguntas)):
    fila_imprimir = preguntas.iloc[sujetos, -4:]
    graf = fila_imprimir.plot(kind="bar")
    graf.bar_label(graf.containers[0])
    plt.title(preguntas.iloc[sujetos, -5])
    plt.ylabel("Calificación")
    plt.axhline(limite_medio, color="red")
    plt.text(0.5, limite_medio + 0.1, "Medio", color="black", ha="center", va="bottom")
    plt.axhline(limite_alto, color="red")
    plt.text(0.5, limite_alto + 0.1, "Alto", color="black", ha="center", va="bottom")
    plt.savefig("Graficas_IRI\\" + str(preguntas.iloc[sujetos, -5]) + ".png")
    plt.clf()

#Voy a guardar el excel de las preguntas
preguntas.to_excel("Calificacion_IRI.xlsx", index=False)



