"""
Helper functions for roadmap management operations.

This module provides utility functions for:
    - XML export/import
    - Reading collaborator lists from Excel
    - Building Excel interfaces with data validation
    - CLI argument parsing
    - Logging configuration

Author: Mustapha ELKAMILI
"""
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
import xml.etree.ElementTree as ET


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

def write_xml(rows: list, xml_output: Path) -> None:
    """
        Write data rows to XML file in format expected by VBA.

        Creates an XML file with a structure that VBA can easily parse. Each row
        becomes a <row> element containing <col1>, <col2>, etc. child elements.

        Args:
            rows (list): List of row data, where each row is a list of values.
                None values are converted to empty strings.
            xml_output (Path): Path where the XML file should be written.

        Returns:
            None

        Example:
            >>> rows = [["Alice", 100], ["Bob", 200]]
            >>> write_xml(rows, Path("output.xml"))
            Creates XML with two <row> elements.
    """
    root = ET.Element("rows")

    for r in rows:
        row_el = ET.SubElement(root, "row")
        for idx, val in enumerate(r, start=1):
            col = ET.SubElement(row_el, f"col{idx}")
            col.text = "" if val is None else str(val)

    tree = ET.ElementTree(root)
    tree.write(xml_output, encoding="utf-8", xml_declaration=True)

def valid_choice(value: str) -> int:
    """
        Validate and convert CLI argument to integer.

        Used as a type validator for argparse to ensure choice arguments are valid integers.

        Args:
            value (str): String value to convert to integer.

        Returns:
            int: The integer value of the input string.

        Raises:
            argparse.ArgumentTypeError: If value cannot be converted to integer.

        Example:
            >>> valid_choice("5")
            5
            >>> valid_choice("abc")
            ArgumentTypeError: Invalid choice: abc. Must be an integer.
    """
    try:
        ivalue = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError(f"Invalid choice: {value}. Must be an integer.")
    return ivalue

def get_collaborators(
    synthese_file: Path | str,
    sheet_name: str = "Gestion_Interfaces",
    min_row: int = 3,
    min_col: int = 2,
    max_col: int = 2) -> list[str]:
    """
        Extract collaborator names from synthesis Excel file.

        Reads collaborator names from a specified sheet and column in the
        synthesis file. Uses a temporary copy to avoid file locking issues.
        Stops reading when encountering an empty cell.

        Args:
            synthese_file (Path | str): Path to the synthesis Excel file.
            sheet_name (str, optional): Name of the sheet containing collaborator
                list. Defaults to "Gestion_Interfaces".
            min_row (int, optional): Starting row number. Defaults to 3.
            min_col (int, optional): Column number to read from (1-indexed).
                Defaults to 2 (column B).
            max_col (int, optional): Ending column number. Defaults to 2.

        Returns:
            list[str]: List of collaborator names, stripped of whitespace.
            Returns empty list if file cannot be read or sheet not found.

        Note:
            Creates a temporary copy of the file to avoid locking issues with open Excel files.
            The temporary file is deleted after reading.
    """
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

def build_interface(template_bytes: bytes, output_path: str, collab_name: str) -> None:
    """
        Build a single collaborator interface Excel file from template.

        Creates a new Excel file for a collaborator based on the template.
        Sets the collaborator name in cell B1 and adds data validation lists
        for pointage entry (week, key, label, function).

        Args:
            template_bytes (bytes): Binary content of the template Excel file.
            output_path (str): Path where the new interface file should be saved.
            collab_name (str): Name of the collaborator to set in the interface.

        Returns:
            None

        Note:
            This function is designed to be called in parallel processes.
            It uses bytes instead of file path to avoid file locking issues in parallel execution.

        Data Validation Lists:
            - Column D: Week (from POINTAGE!A2:A2)
            - Column E: Key (from LC!B3:B1000)
            - Column F: Label (from LC!C3:C1000)
            - Column G: Function (from LC!D3:D1000)
    """
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
    wb: xw.Book,
    list_range: str,
    target_column: str,
    dropdown_sheet: str = "POINTAGE",
    list_sheet: str = "LC",
    end_row: int = 1000,
    start_row: int = 4) -> xw.Book:
    """
        Add data validation list to Excel workbook using xlwings.

        Applies a dropdown list validation to a column range in Excel.
        The dropdown options come from a specified range in another sheet.

        Args:
            wb: xlwings Book object representing the Excel workbook.
            list_range (str): Excel range string (e.g., "B3:B1000") containing
                the list of valid values.
            target_column (str): Column letter (e.g., "E") where validation
                should be applied.
            dropdown_sheet (str, optional): Name of sheet containing target cells.
                Defaults to "POINTAGE".
            list_sheet (str, optional): Name of sheet containing source list.
                Defaults to "LC".
            end_row (int, optional): Last row number for validation range.
                Defaults to 1000.
            start_row (int, optional): First row number for validation range.
                Defaults to 4.

        Returns:
            Book: The modified xlwings Book object.

        Note:
            Removes any existing validation on the target range before adding new validation.
            Uses Excel's native Validation API via xlwings.
    """

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
    except Exception:
        pass

    validation.Add(
        Type=3,
        AlertStyle=1,
        Operator=1,
        Formula1=f"='{list_sheet}'!{source_range.get_address()}"
    )

    return wb

def get_parser() -> argparse.ArgumentParser:
    """
        Create and configure CLI argument parser.

        Sets up argument parser with subcommands for create, delete, pointage, and update operations.
        Each subcommand has its own specific options.

        Returns:
            argparse.ArgumentParser: Configured argument parser ready to parse command-line arguments.

        Commands:
            - create: Create collaborator interfaces
                Options: --way (normal/para/xlw), --archive
            - delete: Delete collaborator interfaces
                Options: --archive, --force
            - pointage: Export time tracking data
                Options: --delete
            - update: Update conditional lists

        Global Options:
            --basedir: Base directory for file operations
    """
    parser = argparse.ArgumentParser(
        description="Roadmap Management CLI - Automate CE VHST roadmap operations including time tracking, interface creation, and data synchronization.",
        usage="roadmap <command>",
    )

    # Global argument
    parser.add_argument(
        "--basedir",
        type=str,
        default="none",
        help="Specify the base directory path containing roadmap files. If not provided, uses platform-specific default or current directory."
    )

    subparsers_action = parser.add_subparsers(dest="action", required=True)
    create_parser = subparsers_action.add_parser("create", help="Generate Excel interface files for all collaborators listed in the synthesis file")
    create_parser.add_argument(
        "--way",
        choices=['xlw', 'normal', 'para'],
        default='normal',
        help="Processing mode: 'normal' for sequential processing (~50s), 'para' for parallel processing (~9s, fastest), or 'xlw' for xlwings-based processing (~3min, best for VBA integration)"
    )
    create_parser.add_argument(
        "--archive",
        action="store_true",
        help="Archive existing RM_Collaborateurs folder before creating new interfaces"
    )

    delete_parser = subparsers_action.add_parser("delete", help="Remove or archive all collaborator interface files from RM_Collaborateurs folder")
    delete_parser.add_argument(
        "--archive",
        action="store_true",
        help="Copy files to Archived folder before moving to Deleted folder. Without this flag, files are moved directly to Deleted folder"
    )

    delete_parser.add_argument(
        "--force",
        action="store_true",
        help="Required flag to confirm deletion operation. Without this flag, the operation will be aborted with a warning"
    )

    pointage_parser = subparsers_action.add_parser("pointage", help="Export time tracking data from collaborator Excel files to XML format for VBA import")
    pointage_parser.add_argument(
        "--delete",
        action="store_true",
        help="Archive the SYNTHESE sheet instead of exporting pointage data. Creates a timestamped archive file containing SYNTHESE and LC sheets"
    )
    subparsers_action.add_parser("update", help="Synchronize conditional lists (LC) from master synthesis file to template and all collaborator interface files")

    return parser
