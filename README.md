# Automatización de la calificación de pruebas psicométricas

Este repositorio contiene scripts para la **calificación automática de pruebas psicométricas**, específicamente el **Inventario de Reactividad Interpersonal (IRI)** y el **MSCEIT (Mayer-Salovey-Caruso Emotional Intelligence Test)**.

Ambos scripts reciben como entrada archivos en formato **Excel**, correspondientes a las salidas de cuestionarios aplicados mediante **Microsoft Forms**, y generan automáticamente calificaciones numéricas y representaciones gráficas por participante.

---

## Descripción de los códigos

### 1. Calificacion IRI (`Calificacion_IRI.py`)


Este script automatiza la calificación del **Inventario de Reactividad Interpersonal (IRI)** a partir de un archivo Excel con las respuestas de los participantes.

#### Funcionalidades principales
- Lectura de las respuestas desde un archivo Excel.
- Inversión automática de los ítems formulados de manera inversa.
- Cálculo de las puntuaciones por cada dimensión del IRI:
  - **Toma de Perspectiva (TP)**
  - **Fantasía (F)**
  - **Preocupación Empática (EC)**
  - **Malestar Personal (PD)**
- Generación de **gráficos de barras individuales por participante**, con puntos de corte visuales (medio y alto).
- Exportación de resultados:
  - Archivo Excel con las calificaciones finales por sujeto.
  - Imagen PNG por participante con su perfil de resultados.

Los gráficos se almacenan automáticamente en una carpeta dedicada y las calificaciones consolidadas se guardan en un nuevo archivo Excel.

#### Librerías necesarias
- `pandas` — lectura y manipulación de datos en Excel  
- `numpy` — manejo de arreglos y operaciones numéricas  
- `matplotlib` — generación y guardado de gráficos  

---

### 2. Calificación MSCEIT (`CalificacionPrueba.py`)

Este script implementa la **calificación automática del MSCEIT** utilizando un enfoque de **consenso basado en frecuencias relativas**.

#### Flujo general
- Lectura de las respuestas crudas desde un archivo Excel.
- Conversión de cada respuesta en su **frecuencia relativa**, según la distribución de respuestas del grupo.
- Cálculo de puntuaciones brutas por:
  - **Tareas**
  - **Ramas**
  - **Áreas**
  - **Índice global de Inteligencia Emocional (CIE)**
- Normalización de las puntuaciones a una escala estándar  
  *(media = 100, desviación estándar = 15)*.
- Exportación de un archivo Excel con los resultados finales por participante.
- Generación automática de una **hoja de perfil gráfica por sujeto**, centrada en la puntuación 100, guardada como imagen PNG.

Este script permite obtener de forma reproducible tanto los resultados numéricos como representaciones gráficas individuales del desempeño emocional.

#### Librerías necesarias
- `pandas` — lectura, transformación y exportación de datos  
- `numpy` — cálculo de puntuaciones, *z-scores* y normalización  
- `matplotlib` — visualización y guardado de hojas de perfil  

---

## Requisitos generales

- Python 3.x
- Archivos de entrada en formato Excel (`.xlsx`)

