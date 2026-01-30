# gj-pdf-converter

PDF to Excel Data Engine 
Part of the GJ Tools Productivity Suite

A robust Python-based engine designed to extract tabular data from PDF files and convert them into professionally formatted Excel spreadsheets (.xlsx).

 Technical Features
 
Dual Extraction Modes: Supports both Lattice and Stream modes.

Automated Formatting: Uses Openpyxl to apply professional styling.

Data Integrity: Built on top of Pandas.

 Project Structure
 
converter.py: The core conversion engine.

requirements.txt: List of necessary Python libraries.

 Getting Started
 
Installation

1. Clone this repository:

git clone https://github.com/gustavosj7/gj-pdf-converter.git

2. Install dependencies:

pip install -r requirements.txt

Basic Usage

from converter import PDFDataEngine

engine = PDFDataEngine("input.pdf", "output.xlsx", use_lattice=True)
engine.run_conversion()
