# gj-pdf-converter

PDF to Excel Data Engine 
Part of the GJ Tools Productivity Suite

A robust Python-based engine designed to extract tabular data from PDF files and convert them into professionally formatted Excel spreadsheets (.xlsx). This tool is optimized for automation workflows and high-accuracy data recovery.

 Technical Features
Dual Extraction Modes: Supports both Lattice (for bordered tables) and Stream (for whitespace-based layouts) using the Tabula-py engine.

Automated Formatting: Uses Openpyxl to automatically apply professional styling.

Data Integrity: Built on top of Pandas to ensure data structures are preserved during conversion.

 Project Structure
converter.py: The core conversion engine (Class-based logic).

requirements.txt: List of necessary Python libraries.

main.py: Entry point for the application.

 Getting Started
Prerequisites
Python 3.8+

Java Runtime Environment (Required by Tabula-py)

Installation
Clone this repository: git clone https://github.com/YOUR_USERNAME/pdf-to-excel-converter.git

Install dependencies: pip install -r requirements.txt

Basic Usage
from converter import PDFDataEngine

Initialize the engine
Set use_lattice=True if the PDF has visible grid lines
engine = PDFDataEngine("input.pdf", "output.xlsx", use_lattice=True)

Run conversion
engine.run_conversion()

 Commercial Context
This repository contains the core logic for the PDF to Excel Pro desktop application. If you are looking for the compiled standalone version (.exe) for Windows, please visit my Gumroad store.

 Contact & Portfolio
Developer: Gustavo (GJ Tools)

Email: gustavosjwork@gmail.com

Services: Python Automation | Data Extraction | Electronic Engineering Student

Disclaimer: This tool is intended for professional and ethical data extraction.
