"""
    Main script for CE VHST Roadmap automation Handles:
        1. Pointage (time tracking export)
        2. Updating conditional lists (LC)
        3. Creating user interfaces
        4. Deleting interfaces
"""
import argparse
import os
import platform
import shutil
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime
from pathlib import Path

import xlwings as xw
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from tqdm import tqdm

from .helpers import build_interface, get_collaborators, logger, valid_choice, add_validation_list

if platform.system() == "Windows":
    BASE_DIR = Path(r"C:\Users\Consultant\OneDrive - IKOSCONSULTING\test_RM\files")
else:
    BASE_DIR = Path("/mnt/c/Users/Consultant/OneDrive - IKOSCONSULTING/test_RM/files")

app = xw.App(visible=False)

class RoadmapManager:
    def __init__(self):
        self.base_path = BASE_DIR
        self.synthese_file = self.base_path / "Synthèse_RM_CE_VHST.xlsx"
        self.template_file = self.base_path / "RM_NOM Prénom.xlsx"
        self.rm_folder = self.base_path / "RM_Collaborateurs"
    
    def check_rm_archive(self):
        """Archive existing folder"""
        if self.rm_folder.exists():
            shutil.move(
                self.rm_folder,
                self.base_path / f"Archive_RM_Collaborateurs_{datetime.now():%d%m%Y_%H%M%S}"
            )
        self.rm_folder.mkdir(exist_ok=True)

    def create_interfaces_fast(self, max_workers=8):
        """Create user interfaces using openpyxl (parallel). estimated time: 9s (51files)"""
        logger.info("[CREATE_INTERFACES] Parallel interface creation")

        # Load collaborators
        collaborators = get_collaborators(self.synthese_file)
        if not collaborators:
            logger.info(
                "[CREATE_INTERFACES] the list of CE is empty."
                f" Please check 'Gestion_Interfaces' sheet in '{self.synthese_file}'"
            )
            return

        logger.info(f"[CREATE_INTERFACES] Found {len(collaborators)} collaborators")
        self.check_rm_archive()

        template_bytes = Path(self.template_file).read_bytes()

        futures = []
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            for collab in collaborators:
                output_path = str(self.rm_folder / f"RM_{collab}.xlsx")
                futures.append(
                    executor.submit(build_interface, template_bytes, output_path, collab)
                )

            for _ in tqdm(futures, desc="Creating interfaces (parallel)"):
                try:
                    _.result()
                except Exception as e:
                    logger.error(f"error: {e}")

        logger.info("[CREATE_INTERFACES] parallel creation complete.")

    def create_interfaces(self):
        """Create user interfaces using openpyxl. estimated time: 50s (51files)"""
        logger.info("[CREATE_INTERFACES] interface creation")

        # Load collaborators
        collaborators = get_collaborators(self.synthese_file)
        if not collaborators:
            logger.info(
                "[CREATE_INTERFACES] the list of CE is empty."
                f" Please check 'Gestion_Interfaces' sheet in '{self.synthese_file}'"
            )
            return

        logger.info(f"[CREATE_INTERFACES] Found {len(collaborators)} collaborators")

        # check if rm_folder exists
        self.check_rm_archive()

        for collab in tqdm(collaborators, desc="Creating interfaces", total=len(collaborators)):
            target = self.rm_folder / f"RM_{collab}.xlsx"

            # Clone template
            wb = load_workbook(self.template_file)
            ws_pointage = wb["POINTAGE"]
            ws_lc = wb["LC"]

            # Write collaborator name
            ws_pointage["B1"].value = collab

            # Create validations
            dv_semaine  = DataValidation(type="list", formula1="='POINTAGE'!$A$2:$A$2")
            dv_cle      = DataValidation(type="list", formula1="='LC'!$B$3:$B$1000")
            dv_libelle  = DataValidation(type="list", formula1="='LC'!$C$3:$C$1000")
            dv_fonction = DataValidation(type="list", formula1="='LC'!$D$3:$D$1000")

            # Assign validation ranges
            validations = [
                (dv_semaine,  'D'),
                (dv_cle,      'E'),
                (dv_libelle,  'F'),
                (dv_fonction, 'G'),
            ]

            for dv, col in validations:
                ws_pointage.add_data_validation(dv)
                dv.ranges.add(f"{col}3:{col}1000")

            wb.save(target)
            wb.close()

        logger.info("[CREATE_INTERFACES] creation done.")

    def create_interfaces_xlwings(self):
        """Create user interfaces using xlwings. estimated time: 3min4s (51files)"""
        logger.info("[CREATE_INTERFACES] Starting interface creation")

        # Load collaborators
        collaborators = get_collaborators(self.synthese_file)
        if not collaborators:
            logger.info(
                "[CREATE_INTERFACES] the list of CE is empty."
                f" Please check 'Gestion_Interfaces' sheet in '{self.synthese_file}'"
            )
            return

        logger.info(f"[CREATE_INTERFACES] Found {len(collaborators)} collaborators")

        self.check_rm_archive()

        for collab in tqdm(collaborators, desc="Creating interfaces", total=len(collaborators)):
            target_path = self.rm_folder / f"RM_{collab}.xlsx"
            shutil.copy2(self.template_file, target_path)

            wb = xw.Book(str(target_path))
            pointage_sheet = wb.sheets["POINTAGE"]
            pointage_sheet["B1"].value = collab

            # Semain
            wb = add_validation_list(
                wb,
                list_sheet="POINTAGE",
                list_range="A2:A2",
                target_column="D",
            )

            # Clef d'imputation
            wb = add_validation_list(
                wb,
                list_range="B3:B1000",
                target_column="E",
            )

            # Libellé
            wb = add_validation_list(
                wb,
                list_range="C3:C1000",
                target_column="F",
            )

            # Fonction
            wb = add_validation_list(
                wb,
                list_range="D3:D1000",
                target_column="G",
            )

            wb.save()
            wb.close()

        logger.info("[CREATE_INTERFACES] creation done.")

    def update_lc(self):
        """Function 2: Update conditional lists (LC) in all personal tools"""
        logger.info("[UPDATE_LC] Starting LC update process")

        synthese_wb = load_workbook(self.synthese_file)
        lc_sheet = synthese_wb["LC"]

        lc_data = []
        for row in lc_sheet.iter_rows(min_row=3, max_row=10000, min_col=2, max_col=234):
            row_data = [cell.value for cell in row]
            if all(cell is None for cell in row_data):
                break
            lc_data.append(row_data)

        synthese_wb.close()

        logger.info(f"[UPDATE_LC] Loaded {len(lc_data)} rows of LC data")

        self._update_lc_in_file(self.template_file, lc_data)

        if self.rm_folder.exists():
            for rm_file in self.rm_folder.glob("*.xlsx"):
                if rm_file.name.startswith("~$"):
                    continue
                self._update_lc_in_file(rm_file, lc_data)

        logger.info("[UPDATE_LC] LC update completed")

    def _update_lc_in_file(self, file_path, lc_data):
        logger.info(f"[UPDATE_LC] Updating {file_path.name}")

        wb = load_workbook(file_path)

        if "LC" not in wb.sheetnames:
            logger.warning(f"[UPDATE_LC] LC sheet not found in {file_path.name}")
            wb.close()
            return

        lc_sheet = wb["LC"]

        for row in lc_sheet.iter_rows(min_row=3, max_row=lc_sheet.max_row):
            for cell in row:
                if cell.column >= 2:
                    cell.value = None

        for row_idx, row_data in enumerate(lc_data, start=3):
            for col_idx, value in enumerate(row_data, start=2):
                lc_sheet.cell(row=row_idx, column=col_idx, value=value)

        wb.save(file_path)
        wb.close()

    def pointage(self, collaborator_file):
        """Export pointage data from collaborator file to synthesis"""
        logger.info(f"[POINTAGE] Processing {collaborator_file}")

        collab_wb = load_workbook(collaborator_file)
        pointage_sheet = collab_wb["POINTAGE"]

        data_to_export = []
        for row in pointage_sheet.iter_rows(min_row=4, max_row=1000, min_col=1, max_col=348):
            row_data = [cell.value for cell in row]
            if all(cell is None for cell in row_data):
                break
            data_to_export.append(row_data)

        collab_wb.close()

        if not data_to_export:
            logger.info("[POINTAGE] No data to export")
            return False

        synthese_wb = load_workbook(self.synthese_file)
        synthese_sheet = synthese_wb["SYNTHESE"]

        pointage_headers = data_to_export[0] if data_to_export else []
        synthese_headers = [cell.value for cell in synthese_sheet[1]]

        if pointage_headers != synthese_headers[:len(pointage_headers)]:
            logger.warning("[POINTAGE] Headers mismatch detected!")
            logger.warning(f"Collab headers: {pointage_headers[:5]}...")
            logger.warning(f"Synthese headers: {synthese_headers[:5]}...")

        last_row = synthese_sheet.max_row
        if last_row < 2:
            last_row = 2

        write_row = last_row + 1
        for row_data in data_to_export:
            for col_idx, value in enumerate(row_data, start=1):
                synthese_sheet.cell(row=write_row, column=col_idx, value=value)
            write_row += 1

        synthese_wb.save(self.synthese_file)
        synthese_wb.close()

        collab_wb = load_workbook(collaborator_file)
        pointage_sheet = collab_wb["POINTAGE"]

        if pointage_sheet.max_row >= 5:
            pointage_sheet.delete_rows(5, pointage_sheet.max_row - 4)

        collab_wb.save(collaborator_file)
        collab_wb.close()

        logger.info(f"[POINTAGE] Successfully exported {len(data_to_export)} rows")
        return True

    def delete_interfaces(self, archive=True):
        """Function 4: Delete all interfaces (optionally archive them)"""
        logger.info("[DELETE_INTERFACES] Starting interface deletion")

        if not self.rm_folder.exists():
            logger.warning("[DELETE_INTERFACES] RM_Collaborateurs folder does not exist")
            return

        deleted_count = 0
        for rm_file in self.rm_folder.glob("*.xlsx"):
            if rm_file.name.startswith("~$"):
                continue

            if archive:
                archive_name = f"Archive_{rm_file.name}"
                archive_path = self.rm_folder / archive_name
                logger.info(f"[DELETE_INTERFACES] Archiving {rm_file.name} -> {archive_name}")
                rm_file.rename(archive_path)
            else:
                logger.info(f"[DELETE_INTERFACES] Deleting {rm_file.name}")
                rm_file.unlink()

            deleted_count += 1

        logger.info(f"[DELETE_INTERFACES] Processed {deleted_count} files")


def main():
    """Main entry point with CLI argument parser"""
    parser = argparse.ArgumentParser(
        description="RoadMap CLI",
        usage="roadmap <command>",
    )

    subparsers = parser.add_subparsers(dest="action", required=True)
    subparsers.add_parser("create", help="Create interfaces (create user tools)")
    subparsers.add_parser("update", help="Update LC (Update conditional lists)")

    pointage_parser = subparsers.add_parser("pointage", help="Pointage (Export time tracking)")
    pointage_parser.add_argument(
        "--choice",
        type=valid_choice,
        default=-1,
        help="Specify a number (int) for a specific file index (default '-1' for all files)"
    )

    delete_parser = subparsers.add_parser("delete", help="Delete interfaces (Remove all user tools)")
    delete_parser.add_argument(
        "--archive",
        choices=["yes", "no"],
        required=True,
        help="Archive files instead of deleting them ('yes' to archive, 'no' to delete permanently)"
    )
    delete_parser.add_argument(
        "--force",
        action="store_true",
        help="Force deletion without confirmation prompt"
    )

    args = parser.parse_args()
    manager = RoadmapManager()

    if args.action == "create":
        # manager.create_interfaces()
        manager.create_interfaces_fast()
        # manager.create_interfaces_xlwings()
        return

    if args.action == "update":
        manager.update_lc()
        return

    if args.action == "pointage":
        if not manager.rm_folder.exists():
            logger.error("RM_Collaborateurs folder not found")
            return

        files = list(
            filepath for filepath in list(manager.rm_folder.glob("*.xlsx"))
            if not filepath.name.startswith("~$")
        )

        if not files:
            logger.warning("No collaborator files found")
            return

        if args.choice == -1:
            for path in files:
                manager.pointage(path)
            return

        try:
            for idx, path in enumerate(files, start=1):
                logger.info(f"  {idx}. {path.name}")

            idx = int(args.choice) - 1
            if 0 <= idx < len(files):
                manager.pointage(files[idx])
            else:
                logger.error("Invalid selection")

        except ValueError:
            logger.error("Invalid input")

        return

    if args.action == "delete":
        if not args.force:
            logger.warning("⚠️  Operation not confirmed. Use --force to proceed.")
            return

        manager.delete_interfaces(archive=args.archive == "yes")
        return

    if args.action == 'exit':
        return


def run():
    """Entry point for console script"""
    main()

if __name__ == "__main__":
    main()
