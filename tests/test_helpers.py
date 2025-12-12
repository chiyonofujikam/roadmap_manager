"""
Helper Functions Tests for Roadmap Manager.

Tests for utility functions including XML operations, Excel file manipulation,
data validation, and file system operations.
"""
import xml.etree.ElementTree as ET
import zipfile
from openpyxl import Workbook, load_workbook

from roadmap.helpers import (
    add_data_validations_to_sheet,
    build_interface,
    get_collaborators,
    load_lc_excel,
    rmtree_with_retry,
    write_xml,
    zip_folder,
)


class TestGetCollaborators:
    """Tests for get_collaborators function."""

    def test_get_collaborators_from_xml(self, tmp_path):
        """TEST-HELP-001: Verify reading collaborator names from XML file."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>CLIGNIEZ Yann</collaborator>
            <collaborator>GANI Karim</collaborator>
            <collaborator>MOUHOUT Marouane</collaborator>
        </collaborators>"""

        xml_file = tmp_path / "collabs.xml"
        xml_file.write_text(xml_content, encoding="utf-8")

        synthese_file = tmp_path / "Synthèse_RM_CE.xlsm"

        collabs = get_collaborators(synthese_file)

        assert collabs == ["CLIGNIEZ Yann", "GANI Karim", "MOUHOUT Marouane"]

    def test_get_collaborators_empty_xml(self, tmp_path):
        """TEST-HELP-002: Verify handling of empty collaborators XML."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
        </collaborators>"""

        xml_file = tmp_path / "collabs.xml"
        xml_file.write_text(xml_content, encoding="utf-8")

        synthese_file = tmp_path / "Synthèse_RM_CE.xlsm"

        collabs = get_collaborators(synthese_file)

        assert collabs == []

    def test_get_collaborators_missing_xml(self, tmp_path):
        """TEST-HELP-003: Verify handling when XML file doesn't exist."""
        synthese_file = tmp_path / "Synthèse_RM_CE.xlsm"

        collabs = get_collaborators(synthese_file)

        assert collabs == []

    def test_get_collaborators_with_whitespace(self, tmp_path):
        """Verify collaborator names are stripped of whitespace."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>  CLIGNIEZ Yann  </collaborator>
            <collaborator>GANI Karim
            </collaborator>
        </collaborators>"""

        xml_file = tmp_path / "collabs.xml"
        xml_file.write_text(xml_content, encoding="utf-8")

        synthese_file = tmp_path / "Synthèse_RM_CE.xlsm"

        collabs = get_collaborators(synthese_file)

        assert collabs == ["CLIGNIEZ Yann", "GANI Karim"]

    def test_get_collaborators_xml_deleted_after_read(self, tmp_path):
        """Verify XML file is deleted after reading."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>NAZIH Imane</collaborator>
        </collaborators>"""

        xml_file = tmp_path / "collabs.xml"
        xml_file.write_text(xml_content, encoding="utf-8")

        synthese_file = tmp_path / "Synthèse_RM_CE.xlsm"

        get_collaborators(synthese_file)

        assert not xml_file.exists()


class TestBuildInterface:
    """Tests for build_interface function."""

    def test_build_interface(self, tmp_path, template_bytes):
        """TEST-HELP-004: Verify building collaborator interface from template."""
        output = tmp_path / "RM_YAHYA Oumaima.xlsx"

        build_interface(template_bytes, str(output), "YAHYA Oumaima")

        assert output.exists()

        generated = load_workbook(output)
        assert generated["POINTAGE"]["B1"].value == "YAHYA Oumaima"
        generated.close()

    def test_build_interface_preserves_lc_sheet(self, tmp_path, template_bytes):
        """Verify LC sheet is preserved in generated interface."""
        output = tmp_path / "RM_NAZIH Imane.xlsx"

        build_interface(template_bytes, str(output), "NAZIH Imane")

        generated = load_workbook(output)
        assert "LC" in generated.sheetnames
        generated.close()

    def test_build_interface_adds_data_validations(self, tmp_path, template_bytes):
        """Verify data validations are added to POINTAGE sheet."""
        output = tmp_path / "RM_CLIGNIEZ Yann.xlsx"

        build_interface(template_bytes, str(output), "CLIGNIEZ Yann")

        generated = load_workbook(output)
        ws = generated["POINTAGE"]

        # Check that data validations exist
        assert len(ws.data_validations.dataValidation) == 4
        generated.close()


class TestWriteXml:
    """Tests for write_xml function."""

    def test_write_xml(self, tmp_path):
        """TEST-HELP-005: Verify XML export with row data."""
        rows = [
            ["CLIGNIEZ Yann", "2024-01", 100],
            ["GANI Karim", "2024-01", 200],
        ]

        xml_output = tmp_path / "test_output.xml"
        write_xml(rows, xml_output)

        assert xml_output.exists()

        tree = ET.parse(xml_output)
        root = tree.getroot()

        assert root.tag == "rows"
        assert len(root.findall("row")) == 2

    def test_write_xml_empty(self, tmp_path):
        """TEST-HELP-006: Verify XML export with empty data."""
        rows = []

        xml_output = tmp_path / "test_empty.xml"
        write_xml(rows, xml_output)

        assert xml_output.exists()

        tree = ET.parse(xml_output)
        root = tree.getroot()

        assert root.tag == "rows"
        assert len(root.findall("row")) == 0

    def test_write_xml_with_none(self, tmp_path):
        """TEST-HELP-007: Verify XML export handles None values."""
        rows = [["MOUHOUT Marouane", None, 100]]

        xml_output = tmp_path / "test_none.xml"
        write_xml(rows, xml_output)

        tree = ET.parse(xml_output)
        root = tree.getroot()
        row = root.find("row")

        col2_text = row.find("col2").text
        assert col2_text is None or col2_text == ""

    def test_write_xml_column_names(self, tmp_path):
        """Verify XML uses correct column naming (col1, col2, etc.)."""
        rows = [["A", "B", "C", "D", "E"]]

        xml_output = tmp_path / "test_cols.xml"
        write_xml(rows, xml_output)

        tree = ET.parse(xml_output)
        root = tree.getroot()
        row = root.find("row")

        for i in range(1, 6):
            assert row.find(f"col{i}") is not None

    def test_write_xml_utf8_encoding(self, tmp_path):
        """Verify XML file uses UTF-8 encoding."""
        rows = [["Élève", "Données"]]

        xml_output = tmp_path / "test_utf8.xml"
        write_xml(rows, xml_output)

        content = xml_output.read_text(encoding="utf-8")
        assert 'encoding="utf-8"' in content.lower() or 'encoding=\'utf-8\'' in content.lower()


class TestZipFolder:
    """Tests for zip_folder function."""

    def test_zip_folder(self, tmp_path):
        """TEST-HELP-008: Verify folder compression."""
        source_folder = tmp_path / "source"
        source_folder.mkdir()
        (source_folder / "file1.txt").write_text("content1")
        (source_folder / "file2.txt").write_text("content2")

        zip_path = tmp_path / "archive.zip"
        zip_folder(source_folder, zip_path)

        assert zip_path.exists()

        with zipfile.ZipFile(zip_path, 'r') as zipf:
            names = zipf.namelist()
            assert len(names) == 2

    def test_zip_folder_preserves_structure(self, tmp_path):
        """Verify folder structure is preserved in zip."""
        source_folder = tmp_path / "source"
        source_folder.mkdir()
        subfolder = source_folder / "subfolder"
        subfolder.mkdir()
        (subfolder / "nested.txt").write_text("nested content")

        zip_path = tmp_path / "archive.zip"
        zip_folder(source_folder, zip_path)

        with zipfile.ZipFile(zip_path, 'r') as zipf:
            names = zipf.namelist()
            assert any("subfolder" in name for name in names)

    def test_zip_folder_empty(self, tmp_path):
        """Verify empty folder creates empty zip."""
        source_folder = tmp_path / "empty_source"
        source_folder.mkdir()

        zip_path = tmp_path / "empty_archive.zip"
        zip_folder(source_folder, zip_path)

        assert zip_path.exists()

        with zipfile.ZipFile(zip_path, 'r') as zipf:
            assert len(zipf.namelist()) == 0


class TestRmtreeWithRetry:
    """Tests for rmtree_with_retry function."""

    def test_rmtree_with_retry(self, tmp_path):
        """TEST-HELP-009: Verify folder removal with retry logic."""
        folder = tmp_path / "to_delete"
        folder.mkdir()
        (folder / "file.txt").write_text("content")

        result = rmtree_with_retry(folder)

        assert result is True
        assert not folder.exists()

    def test_rmtree_nested_folders(self, tmp_path):
        """Verify nested folder removal."""
        folder = tmp_path / "parent"
        folder.mkdir()
        child = folder / "child"
        child.mkdir()
        (child / "file.txt").write_text("content")

        result = rmtree_with_retry(folder)

        assert result is True
        assert not folder.exists()

    def test_rmtree_nonexistent_folder(self, tmp_path):
        """Verify handling of non-existent folder."""
        folder = tmp_path / "nonexistent"

        try:
            result = rmtree_with_retry(folder)
        except FileNotFoundError:
            pass


class TestAddDataValidations:
    """Tests for add_data_validations_to_sheet function."""

    def test_add_data_validations(self, tmp_path):
        """TEST-HELP-010: Verify data validation lists are added correctly."""
        wb = Workbook()
        ws = wb.active
        ws.title = "POINTAGE"
        wb.create_sheet("LC")

        add_data_validations_to_sheet(ws, start_row=3)

        assert len(ws.data_validations.dataValidation) == 4

        output = tmp_path / "test_validations.xlsx"
        wb.save(output)
        wb.close()

    def test_add_data_validations_clears_existing(self, tmp_path):
        """Verify existing validations are cleared before adding new ones."""
        wb = Workbook()
        ws = wb.active
        ws.title = "POINTAGE"
        wb.create_sheet("LC")

        add_data_validations_to_sheet(ws, start_row=3)
        add_data_validations_to_sheet(ws, start_row=3)

        assert len(ws.data_validations.dataValidation) == 4
        wb.close()


class TestLoadLcExcel:
    """Tests for load_lc_excel function."""

    def test_load_lc_excel(self, tmp_path):
        """TEST-HELP-011: Verify LC data loading from Excel file."""
        wb = Workbook()
        ws = wb.active
        ws.title = "LC"

        ws["B2"] = "Key1"
        ws["C2"] = "Label1"
        ws["D2"] = "Func1"
        ws["E2"] = "Extra1"
        ws["B3"] = "Key2"
        ws["C3"] = "Label2"
        ws["D3"] = "Func2"

        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        lc_data = load_lc_excel(tmp_path)

        assert len(lc_data) == 2
        assert lc_data[0][0] == "Key1"
        assert lc_data[0][1] == "Label1"

    def test_load_lc_excel_missing_file(self, tmp_path):
        """Verify handling when LC.xlsx doesn't exist."""
        lc_data = load_lc_excel(tmp_path)

        assert lc_data == []

    def test_load_lc_excel_empty_sheet(self, tmp_path):
        """Verify handling of empty LC sheet."""
        wb = Workbook()
        ws = wb.active
        ws.title = "LC"
        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        lc_data = load_lc_excel(tmp_path)

        assert lc_data == []

    def test_load_lc_excel_deletes_file(self, tmp_path):
        """Verify LC.xlsx is deleted after reading."""
        wb = Workbook()
        ws = wb.active
        ws.title = "LC"
        ws["B2"] = "Key1"
        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        load_lc_excel(tmp_path)

        assert not (tmp_path / "LC.xlsx").exists()

    def test_load_lc_excel_missing_lc_sheet(self, tmp_path):
        """Verify handling when LC sheet is missing."""
        wb = Workbook()
        ws = wb.active
        ws.title = "OtherSheet"
        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        lc_data = load_lc_excel(tmp_path)

        assert lc_data == []
