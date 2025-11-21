import argparse
import io
import logging
import os
import shutil
import tempfile
from pathlib import Path

import xlwings as xw
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from tqdm import tqdm

os.makedirs(os.path.join(os.getcwd(), ".logs"), exist_ok=True)

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

def get_collaborators(synthese_file, sheet_name="Gestion_Interfaces",
                      min_row=3, min_col=2, max_col=2) -> list[str]:
    collabs = []
    synthese_file = Path(synthese_file)

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=synthese_file.suffix) as tmp:
            temp_path = Path(tmp.name)
        shutil.copy2(synthese_file, temp_path)
        synthese_wb = load_workbook(temp_path, read_only=True, data_only=True)

        for row in synthese_wb[sheet_name].iter_rows(
            min_row=min_row, min_col=min_col, max_col=max_col
        ):
            value = row[0].value
            if not value or str(value).strip() == "":
                break
            collabs.append(str(value).strip())

        synthese_wb.close()
        temp_path.unlink(missing_ok=True)

    except Exception as err:
        logger.error(f"Error while reading {synthese_file.name}: {err}")

    return collabs

def build_interface(template_bytes, output_path, collab_name):
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

def get_parser():
    parser = argparse.ArgumentParser(
        description="RoadMap CLI",
        usage="roadmap <command>",
    )

    # Global argument
    parser.add_argument(
        "--basedir",
        type=str,
        default="none",
        help="Base directory for all file operations (default: current directory)."
    )

    subparsers_action = parser.add_subparsers(dest="action", required=True)
    create_parser = subparsers_action.add_parser("create", help="Create interfaces (create user tools)")
    create_parser.add_argument(
        "--way",
        choices=['xlw', 'normal', 'para'],
        default='normal',
        help="Choose the processing mode for RM generation: 'normal' (standard sequential), 'para' (parallel for faster creation), or 'xlw' (optimized for interacting with Excel/VBA via xlwings)."
    )
    create_parser.add_argument(
        "--archive",
        action="store_true",
        help="Archive files instead")

    delete_parser = subparsers_action.add_parser("delete", help="Delete interfaces (Remove all user tools)")
    delete_parser.add_argument(
        "--archive",
        action="store_true",
        help="Archive files instead of deleting them ('yes' to archive, 'no' to delete permanently)")
    delete_parser.add_argument(
        "--force",
        action="store_true",
        help="Force deletion without confirmation prompt")

    pointage_parser = subparsers_action.add_parser("pointage", help="Pointage (Export time tracking)")
    pointage_parser.add_argument(
        "--choice",
        type=valid_choice,
        default=-1,
        help="Specify a number (int) for a specific file index (default '-1' for all files)")
    subparsers_action.add_parser("update", help="Update LC (Update conditional lists)")

    return parser
