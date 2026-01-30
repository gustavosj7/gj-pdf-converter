import os
import tabula
import pandas as pd
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

class PDFDataEngine:
    """
    GJ Tools PDF to Excel Conversion Engine.
    Handles extraction and automated professional formatting.
    """
    
    def __init__(self, input_path, output_path, use_lattice=False):
        self.input_path = input_path
        self.output_path = output_path
        # Lattice: for tables with lines | Stream: for whitespace-based layouts
        self.extraction_mode = "lattice" if use_lattice else "stream"

    def run_conversion(self):
        """Executes extraction and applies business formatting."""
        print(f"[*] Processing file using {self.extraction_mode} mode...")
        
        try:
            # 1. Extraction via Tabula
            tables = tabula.read_pdf(
                self.input_path, 
                pages="all", 
                lattice=(self.extraction_mode == "lattice"), 
                stream=(self.extraction_mode == "stream")
            )

            if not tables:
                print("[!] No tables found in the document.")
                return False

            # 2. Excel writing and formatting
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                for i, df in enumerate(tables):
                    sheet_name = f"Table_{i+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    self._apply_styling(writer.sheets[sheet_name])
            
            print(f"[+] Success: File saved to {self.output_path}")
            return True

        except Exception as e:
            print(f"[#] Critical error during conversion: {e}")
            return False

    def _apply_styling(self, ws):
        """Applies borders, header bolding, and auto-column width."""
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'), 
            top=Side(style='thin'), bottom=Side(style='thin')
        )

        for col in ws.columns:
            max_length = 0
            column_letter = get_column_letter(col[0].column)
            
            for cell in col:
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center', horizontal='left')
                
                # Header styling
                if cell.row == 1:
                    cell.font = Font(bold=True)
                
                # Calculate required width
                try:
                    cell_val_len = len(str(cell.value))
                    if cell_val_len > max_length:
                        max_length = cell_val_len
                except:
                    pass
            
            ws.column_dimensions[column_letter].width = max_length + 3

# --- Standard CLI Usage ---
if __name__ == "__main__":
    # Example setup
    target_file = "input_sample.pdf"
    if os.path.exists(target_file):
        converter = PDFDataEngine(target_file, "converted_output.xlsx", use_lattice=True)
        converter.run_conversion()
    else:
        print("[!] Target PDF not found. Place a PDF file in the root folder.")
