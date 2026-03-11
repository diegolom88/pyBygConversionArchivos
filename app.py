import sys
import traceback
import warnings
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime

def convert_files(ruta_base, carpetas):
    converted_files = []
    bad_line_warnings = []
    for carpeta in carpetas:
        carpeta_path = ruta_base / carpeta
        for file_path in carpeta_path.glob('*.csv'):
            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always")
                df = pd.read_csv(file_path, encoding='ISO-8859-1', on_bad_lines='warn')
            for w in caught:
                bad_line_warnings.append(f"{file_path.name}: {w.message}")
            df.to_excel(file_path.with_suffix('.xlsx'), index=False)
            print(f"Conversion de {file_path.name} completada")
            file_path.unlink()
            converted_files.append(file_path.name)

    return converted_files, bad_line_warnings

##### Main script
if __name__ == "__main__":
    try:
        ruta_base = Path("C:/Users/Administrador/OneDrive - Desarrollo y Construcciones Urbanas SA de CV/BI/BYGSA")
        carpetas = [
            "Compensaciones",
            "Facturacion y Cobranza/Cobranza",
            "Facturacion y Cobranza/Facturas",
            "Facturacion y Cobranza/Notas de Credito",
            "Catalogo/CatalogoSis",
            "Flujo/Flujos Anuales Sis",
            "Mayores/MayoresSis",
            "Produccion/Produccion",
            "Facturas Portal"
        ]
        logs_folder = Path("C:/Users/Administrador/ScriptsBI/Logs")

        log_path = logs_folder / f"bygConversionArchivos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

        # Create logs folder if it doesn't exist
        log_path.parent.mkdir(parents=True, exist_ok=True)

        converted_files, bad_line_warnings = convert_files(ruta_base, carpetas)

        with open(log_path, 'a') as log_file:
            if bad_line_warnings:
                log_file.write("Lineas con problemas:\n")
                for warning_msg in bad_line_warnings:
                    log_file.write(f"  {warning_msg}\n")
                log_file.write("\n")
            message = f"Conversion de archivos completada: {converted_files}\n"
            log_file.write(message)
            print(message)

    except Exception as e:
        error_message = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: {e}\n{traceback.format_exc()}\n"
        print(error_message, file=sys.stderr)
        try:
            with open(log_path, 'a') as log_file:
                log_file.write(error_message)
        except Exception:
            pass
