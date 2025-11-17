import argparse
import io
import logging

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(".logs/roadmap.log", mode="a", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def valid_choice(value):
    try:
        ivalue = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError(f"Invalid choice: {value}. Must be an integer.")
    return ivalue

def get_collaborators(
    synthese_file, sheet_name: str = "Gestion_Interfaces",
    min_row: int = 3,
    min_col: int = 2, max_col: int = 2) -> list[str]:
    """ Load collaborators CE """
    collabs = []
    synthese_wb = load_workbook(synthese_file, read_only=True)

    for row in synthese_wb[sheet_name].iter_rows(min_row=min_row, min_col=min_col, max_col=max_col):
        value = row[0].value
        if value is None or str(value).strip() == "":
            break
        collabs.append(str(value).strip())
    synthese_wb.close()

    return collabs

def build_interface(template_bytes, output_path, collab_name):
    # wb = load_workbook(template_path)
    wb = load_workbook(filename=io.BytesIO(template_bytes))

    ws_pointage = wb["POINTAGE"]

    # Write collaborator name
    ws_pointage["B1"].value = collab_name

    # Create validations
    dv_semaine = DataValidation(type="list", formula1="='POINTAGE'!$A$2:$A$2")
    dv_cle = DataValidation(type="list", formula1="='LC'!$B$3:$B$1000")
    dv_libelle = DataValidation(type="list", formula1="='LC'!$C$3:$C$1000")
    dv_fonction = DataValidation(type="list", formula1="='LC'!$D$3:$D$1000")

    validations = [
        (dv_semaine,  'D'),
        (dv_cle,      'E'),
        (dv_libelle,  'F'),
        (dv_fonction, 'G'),
    ]

    for dv, col in validations:
        ws_pointage.add_data_validation(dv)
        dv.ranges.add(f"{col}4:{col}1000")

    wb.save(output_path)
    wb.close()

def add_validation_list(
    wb,
    list_range: str,
    target_column: str,
    dropdown_sheet: str = "POINTAGE",
    list_sheet: str = "LC",
    end_row: int = 1000,
    start_row = 4):

    # Build source address for formula
    ws_list = wb.sheets[list_sheet]
    source_range = ws_list.range(list_range)

    # Build target range: column from start_row to end of Excel
    ws_dropdown = wb.sheets[dropdown_sheet]
    target_range = ws_dropdown.range(f"{target_column}{start_row}:{target_column}{end_row}")

    # Apply validation
    validation = target_range.api.Validation
    try:
        validation.Delete()
    except:
        pass

    validation.Add(
        Type=3,
        AlertStyle=1,
        Operator=1,
        Formula1=f"='{list_sheet}'!{source_range.get_address()}"
    )

    return wb
