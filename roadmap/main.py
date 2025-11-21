"""
Description:
    Main script for CE VHST Roadmap automation Handles:
        1. Pointage (time tracking export)
        2. Updating conditional lists (LC)
        3. Creating user interfaces
        4. Deleting interfaces

Author: Mustapha ELKAMILI
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

from .helpers import (add_validation_list, build_interface, get_collaborators,
                      get_parser, logger, valid_choice)

app = xw.App(visible=False)


class RoadmapManager:
    def __init__(self, base_dir: str):
        self.base_path = base_dir

        # Define paths for configuration files
        self.synthese_file = self.base_path / "Synthèse_RM_CE.xlsm"
        self.template_file = self.base_path / "RM_template.xlsx"

        # Define and create necessary folders
        self.rm_folder = self.base_path / "RM_Collaborateurs"
        self.archived_folder = self.base_path / "Archived"
        self.deleted_folder = self.base_path / "Deleted"

        for folder in [self.rm_folder, self.archived_folder, self.deleted_folder]:
            folder.mkdir(exist_ok=True)

        # Check existence of essential files
        self.all_ok = all([
            self.synthese_file.exists(),
            self.template_file.exists()
        ])

        if not self.all_ok:
            logger.error("Required files 'Synthese_RM_CE.xlsm' or 'RM_template.xlsx' are missing. Please check the base directory.")

    def check_roadmap_archive(self) -> Path:
        """Archive existing folder"""
        if self.rm_folder.exists():
            rm_count = sum(1 for f in self.rm_folder.glob("*.xlsx") if f.is_file())
            path_rm = self.archived_folder / f"Archive_RM_Collaborateurs_{datetime.now():%d%m%Y_%H%M%S}"
            if rm_count != 0:
                shutil.move(self.rm_folder, path_rm)
            return path_rm

        self.rm_folder.mkdir(exist_ok=True)
        return self.rm_folder

    def create_interfaces_fast(self, archive: bool, max_workers=8):
        """Create user interfaces using openpyxl (parallel). estimated time: 9s (51files)"""
        if not self.all_ok:
            return

        logger.info("[CREATE_INTERFACES] Parallel processing mode interface creation")

        collaborators = get_collaborators(self.synthese_file)
        if not collaborators:
            logger.info(
                "[CREATE_INTERFACES] the list of CE is empty."
                f" Please check 'Gestion_Interfaces' sheet in '{self.synthese_file}'")
            return

        logger.info(f"[CREATE_INTERFACES] Found {len(collaborators)} collaborators")

        if archive:
            self.check_roadmap_archive()

        try:
            template_bytes = Path(self.template_file).read_bytes()
        except PermissionError:
            logger.error(f"'{self.template_file}' is opened. Please close the excel file")
            return

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

    def create_interfaces(self, archive: bool):
        """Create user interfaces using openpyxl. estimated time: 50s (51files)"""
        if not self.all_ok:
            return

        logger.info("[CREATE_INTERFACES] interface creation (Normal processing mode)")

        collaborators = get_collaborators(self.synthese_file)[:3]
        if not collaborators:
            logger.info(
                "[CREATE_INTERFACES] the list of CE is empty."
                f" Please check 'Gestion_Interfaces' sheet in '{self.synthese_file}'")
            return

        logger.info(f"[CREATE_INTERFACES] Found {len(collaborators)} collaborators, synthese file : {self.synthese_file.name}")

        if archive:
            self.check_roadmap_archive()

        for collab in tqdm(collaborators, desc="Creating interfaces", total=len(collaborators)):
            target = self.rm_folder / f"RM_{collab}.xlsx"

            try:
                wb = load_workbook(self.template_file)
            except PermissionError:
                logger.error(f"'{self.template_file}' is opened. Please close the excel file")
                return
            
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

    def create_interfaces_xlwings(self, archive: bool):
        """Create user interfaces using xlwings. estimated time: 3min4s (51files)"""
        if not self.all_ok:
            return

        logger.info("[CREATE_INTERFACES] interface creation (xlwings processing mode)")

        collaborators = get_collaborators(self.synthese_file)
        if not collaborators:
            logger.info(
                "[CREATE_INTERFACES] the list of CE is empty."
                f" Please check 'Gestion_Interfaces' sheet in '{self.synthese_file}'")
            return

        logger.info(f"[CREATE_INTERFACES] Found {len(collaborators)} collaborators")

        if archive:
            self.check_roadmap_archive()

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

    def delete_interfaces(self, archive: bool):
        """Move RM_Collaborateurs into Deleted/ with timestamp, or archive normally."""
        if not self.all_ok:
            return

        logger.info("[DELETE_INTERFACES] Starting interface deletion")

        rm_folder = self.rm_folder

        if not rm_folder.exists():
            logger.warning("[DELETE_INTERFACES] RM_Collaborateurs folder does not exist")
            return

        rm_count = sum(1 for f in rm_folder.glob("*.xlsx") if f.is_file())

        if archive:
            rm_folder = self.check_roadmap_archive()
            logger.info(f"[DELETE_INTERFACES] Archived {rm_count} interface file(s) to {rm_folder.name}")

        try:
            target_path = self.deleted_folder / f"Deleted_RM_Collaborateurs_{datetime.now().strftime('%d%m%Y_%H%M%S')}"
            if rm_count != 0:
                shutil.copytree(rm_folder, target_path)
                logger.info(f"[DELETE_INTERFACES] Deleted & Moved {rm_count} interface file(s) to {target_path.name}")

        except Exception as e:
            logger.error(f"[DELETE_INTERFACES] Error while moving folder: {e}")
            return

    def pointage(self, collaborator_file):
        """Export pointage data from collaborator file to synthesis"""
        if not self.all_ok:
            return

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

    def update_lc(self):
        """Function 2: Update conditional lists (LC) in all personal tools"""
        if not self.all_ok:
            return

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


def main():
    """Main entry point with CLI argument parser"""
    if platform.system() == "Windows":
        BASE_DIR = Path(r"C:\Users\Consultant\OneDrive - IKOSCONSULTING\test_RM\files")
    else:
        BASE_DIR = Path("/mnt/c/Users/Consultant/OneDrive - IKOSCONSULTING/test_RM/files")

    args = get_parser().parse_args()
    manager = RoadmapManager(base_dir=BASE_DIR if args.basedir == 'none' else Path(args.basedir))

    if args.action == "create":
        if args.way == 'normal':
            manager.create_interfaces(args.archive)
        elif args.way == 'para':
            manager.create_interfaces_fast(args.archive)
        elif args.way == 'xlw':
            manager.create_interfaces_xlwings(args.archive)
        else:
            logger.error(f"Unknown '--way' argument '{args.way}'. Valid choices are 'normal', 'para', and 'xlw'.")
        return

    if args.action == "delete":
        if not args.force:
            logger.warning("⚠️  Operation not confirmed. Use --force to proceed.")
            return

        manager.delete_interfaces(archive=args.archive)
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

    if args.action == "update":
        manager.update_lc()
        return

    if args.action == 'exit':
        return


def run():
    """Entry point for console script"""
    main()

if __name__ == "__main__":
    main()
