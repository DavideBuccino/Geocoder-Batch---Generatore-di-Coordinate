*Read this in other languages: [🇮🇹 Italiano](#-italiano) | [🇬🇧 English](#-english)*

---

# 🇮🇹 Italiano

# 📍 Geocoder GIS Batch: Imputation & Color Matrix (V7)

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data_Analysis-150458?style=for-the-badge&logo=pandas)
![Geopy](https://img.shields.io/badge/Geopy-Geocoding-red?style=for-the-badge)
![OpenStreetMap](https://img.shields.io/badge/OSM-Nominatim-7289DA?style=for-the-badge&logo=openstreetmap)

## 📌 Panoramica del Progetto
Questo repository contiene uno strumento avanzato in Python per la **geocodifica massiva (batch)** di indirizzi da file CSV o Excel. A differenza di un normale geocoder, questo script integra una logica di **Data Cleaning** e **Classificazione Geografica** automatica.

Lo script non si limita a trovare le coordinate (Lat/Lon), ma analizza la qualità del dato, ripara i campi mancanti (Città o CAP) e restituisce un report Excel formattato con una matrice di colori per un'ispezione visiva immediata della qualità del database.

---

## 🛠️ Caratteristiche Principali
* **Geocodifica Intelligente:** Utilizza l'API Nominatim (OpenStreetMap) con gestione dei timeout e dei tentativi automatici.
* **Data Imputation (Riparazione):** Se il database originale ha CAP o Città mancanti, lo script li "ripara" utilizzando i risultati ufficiali della geocodifica.
* **Classificazione Automatica:** Suddivide i record in tre categorie: Capoluoghi/Grandi Città, Piccoli Comuni o Estero.
* **Output Multi-formato:** Esporta i risultati in **Excel (formattato)**, **CSV** e **GeoJSON** (pronto per essere caricato su ArcGIS o QGIS).

---

## 🔬 Workflow Tecnico

1.  **Input flessibile:** Lo script accetta file Excel o CSV e permette all'utente di mappare dinamicamente le colonne (Indirizzo, CAP, Città).
2.  **Processo Batch:** Esegue le richieste rispettando i limiti di velocità (1 req/sec) per garantire la stabilità del servizio.
3.  **Matrice di Colori:** Durante l'esportazione in Excel, viene applicato uno stile condizionale per evidenziare la natura del dato.

### 🎨 Legenda Colori Excel
| Colore | Significato |
| :--- | :--- |
| 🟩 **Verde** | Capoluogo o Grande Città Italiana |
| 🟨 **Giallo** | Provincia o Piccolo Comune Italiano |
| 🟥 **Rosso** | Località Estera (Internazionale) |
| 🟦 **Azzurro** | **Dato Riparato:** Cella originariamente vuota riempita dallo script |

---

## 🚀 Come iniziare

1.  **Installazione librerie:**
    ```bash
    pip install pandas geopy openpyxl
    ```
2.  **Esecuzione:**
    Esegui lo script `geocoder_batch.py` (o apri il Notebook) e segui le istruzioni nel terminale per mappare le tue colonne.

---

<br><br>

---

# 🇬🇧 English

# 📍 Geocoder GIS Batch: Imputation & Color Matrix (V7)

## 📌 Project Overview
This repository features an advanced Python tool for **batch geocoding** addresses from CSV or Excel files. Unlike standard geocoders, this script integrates **Data Cleaning** and automatic **Geographic Classification** logic.

The script goes beyond finding coordinates (Lat/Lon); it analyzes data quality, repairs missing fields (City or ZIP code), and returns a formatted Excel report with a color matrix for immediate visual inspection of database quality.

---

## 🛠️ Key Features
* **Smart Geocoding:** Uses the Nominatim API (OpenStreetMap) with built-in timeout and retry management.
* **Data Imputation:** If the original database has missing ZIP codes or Cities, the script "repairs" them using official geocoding results.
* **Automatic Classification:** Categorizes records into three groups: Major Cities/Capitals, Small Towns, or International.
* **Multi-format Output:** Exports results to **formatted Excel**, **CSV**, and **GeoJSON** (ready for ArcGIS or QGIS).

---

## 🔬 Technical Workflow

1.  **Flexible Input:** Accepts Excel or CSV files and allows the user to dynamically map columns (Address, ZIP, City).
2.  **Batch Processing:** Executes requests respecting speed limits (1 req/sec) to ensure service stability.
3.  **Color Matrix:** During Excel export, conditional styling is applied to highlight the nature of the data.

### 🎨 Excel Color Legend
| Color | Meaning |
| :--- | :--- |
| 🟩 **Green** | Major Italian City or Provincial Capital |
| 🟨 **Yellow** | Small Town or Province |
| 🟥 **Red** | International Location |
| 🟦 **Blue** | **Repaired Data:** Originally empty cell filled by the script |

---

## 🚀 Getting Started

1.  **Install dependencies:**
    ```bash
    pip install pandas geopy openpyxl
    ```
2.  **Usage:**
    Run the `geocoder_batch.py` script (or open the Notebook) and follow the terminal instructions to map your columns.

---
> *Developed as a utility for GIS Portfolio Data Preparation.*
