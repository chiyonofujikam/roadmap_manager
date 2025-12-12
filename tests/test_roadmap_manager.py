"""
RoadmapManager Integration Tests for Roadmap Manager.

Tests for the RoadmapManager class including initialization,
interface creation, deletion, pointage export, and LC updates.
"""
import xml.etree.ElementTree as ET
from openpyxl import Workbook, load_workbook

from roadmap.roadmap import RoadmapManager


class TestRoadmapManagerInit:
    """Tests for RoadmapManager initialization."""

    def test_init_with_valid_files(self, setup_test_environment):
        """TEST-INT-001: Verify RoadmapManager initializes with correct paths."""
        tmp_path = setup_test_environment

        manager = RoadmapManager(tmp_path)

        assert manager.base_path == tmp_path
        assert manager.synthese_file == tmp_path / "Synthèse_RM_CE.xlsm"
        assert manager.template_file == tmp_path / "RM_template.xlsx"
        assert manager.rm_folder == tmp_path / "RM_Collaborateurs"
        assert manager.archived_folder == tmp_path / "Archived"
        assert manager.deleted_folder == tmp_path / "Deleted"
        assert manager.all_ok is True

    def test_init_missing_files(self, tmp_path):
        """TEST-INT-002: Verify RoadmapManager handles missing files."""
        manager = RoadmapManager(tmp_path)

        assert manager.all_ok is False

    def test_init_missing_template(self, tmp_path):
        """Verify handling when only synthese file exists."""
        (tmp_path / "Synthèse_RM_CE.xlsm").touch()

        manager = RoadmapManager(tmp_path)

        assert manager.all_ok is False

    def test_init_missing_synthese(self, tmp_path):
        """Verify handling when only template file exists."""
        (tmp_path / "RM_template.xlsx").touch()

        manager = RoadmapManager(tmp_path)

        assert manager.all_ok is False

    def test_init_creates_folders(self, tmp_path):
        """Verify initialization creates required folders."""
        (tmp_path / "Synthèse_RM_CE.xlsm").touch()
        (tmp_path / "RM_template.xlsx").touch()

        manager = RoadmapManager(tmp_path)

        assert (tmp_path / "RM_Collaborateurs").exists()
        assert (tmp_path / "Archived").exists()
        assert (tmp_path / "Deleted").exists()


class TestCreateInterfaces:
    """Tests for interface creation methods."""

    def test_create_interfaces_normal(self, setup_test_environment):
        """TEST-INT-003: Verify interface creation in normal mode."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        manager.create_interfaces()

        rm_folder = tmp_path / "RM_Collaborateurs"
        created_files = list(rm_folder.glob("RM_*.xlsx"))

        assert len(created_files) == 3

        # Verify file names match collaborators
        file_names = {f.name for f in created_files}
        assert "RM_CLIGNIEZ Yann.xlsx" in file_names
        assert "RM_GANI Karim.xlsx" in file_names
        assert "RM_MOUHOUT Marouane.xlsx" in file_names

    def test_create_interfaces_parallel(self, setup_test_environment):
        """TEST-INT-004: Verify interface creation in parallel mode."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        manager.create_interfaces_fast(max_workers=2)

        rm_folder = tmp_path / "RM_Collaborateurs"
        created_files = list(rm_folder.glob("RM_*.xlsx"))

        assert len(created_files) == 3

    def test_create_interfaces_skip_existing(self, setup_test_environment):
        """TEST-INT-005: Verify existing files are not overwritten."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        rm_folder = tmp_path / "RM_Collaborateurs"
        existing_file = rm_folder / "RM_CLIGNIEZ Yann.xlsx"

        # Create existing file with marker
        wb = Workbook()
        ws = wb.active
        ws.title = "POINTAGE"
        ws["A1"] = "EXISTING_MARKER"
        ws["B1"] = "CLIGNIEZ Yann"
        wb.create_sheet("LC")
        wb.save(existing_file)
        wb.close()

        manager.create_interfaces()

        # Verify marker still exists
        wb = load_workbook(existing_file)
        assert wb["POINTAGE"]["A1"].value == "EXISTING_MARKER"
        wb.close()

    def test_create_interfaces_empty_collaborators(self, tmp_path):
        """Verify handling when no collaborators in list."""
        # Create required files
        (tmp_path / "Synthèse_RM_CE.xlsm").touch()

        template = Workbook()
        template.active.title = "POINTAGE"
        template.create_sheet("LC")
        template.save(tmp_path / "RM_template.xlsx")
        template.close()

        # Create empty collabs.xml
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
        </collaborators>"""
        (tmp_path / "collabs.xml").write_text(xml_content, encoding="utf-8")

        manager = RoadmapManager(tmp_path)
        manager.create_interfaces()

        rm_folder = tmp_path / "RM_Collaborateurs"
        created_files = list(rm_folder.glob("RM_*.xlsx"))

        assert len(created_files) == 0

    def test_create_interfaces_all_exist(self, setup_test_environment_with_interfaces):
        """Verify no action when all files already exist."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        # Get file modification times before
        rm_folder = tmp_path / "RM_Collaborateurs"
        before_times = {f.name: f.stat().st_mtime for f in rm_folder.glob("RM_*.xlsx")}

        # Need to recreate collabs.xml since it was deleted
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>CLIGNIEZ Yann</collaborator>
            <collaborator>GANI Karim</collaborator>
            <collaborator>MOUHOUT Marouane</collaborator>
        </collaborators>"""
        (tmp_path / "collabs.xml").write_text(xml_content, encoding="utf-8")

        manager.create_interfaces()

        # Files should not be modified (times unchanged)
        after_times = {f.name: f.stat().st_mtime for f in rm_folder.glob("RM_*.xlsx")}
        assert before_times == after_times

    def test_created_interface_has_collaborator_name(self, setup_test_environment):
        """Verify created interface has collaborator name in B1."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        manager.create_interfaces()

        rm_folder = tmp_path / "RM_Collaborateurs"
        collab_file = rm_folder / "RM_CLIGNIEZ Yann.xlsx"

        wb = load_workbook(collab_file)
        assert wb["POINTAGE"]["B1"].value == "CLIGNIEZ Yann"
        wb.close()

    def test_created_interface_has_data_validations(self, setup_test_environment):
        """Verify created interface has data validations."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        manager.create_interfaces()

        rm_folder = tmp_path / "RM_Collaborateurs"
        collab_file = rm_folder / "RM_CLIGNIEZ Yann.xlsx"

        wb = load_workbook(collab_file)
        ws = wb["POINTAGE"]

        assert len(ws.data_validations.dataValidation) == 4
        wb.close()


class TestDeleteInterfaces:
    """Tests for interface deletion methods."""

    def test_delete_and_archive(self, setup_test_environment_with_interfaces):
        """TEST-INT-006: Verify deletion with archiving."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        manager.delete_and_archive_interfaces(archive=True)

        rm_folder = tmp_path / "RM_Collaborateurs"
        archived_folder = tmp_path / "Archived"
        deleted_folder = tmp_path / "Deleted"

        # RM folder should be deleted or empty
        assert not rm_folder.exists() or len(list(rm_folder.glob("*.xlsx"))) == 0

        # Archive should have a zip file
        assert len(list(archived_folder.glob("*.zip"))) > 0

        # Deleted should have a zip file
        assert len(list(deleted_folder.glob("*.zip"))) > 0

    def test_delete_without_archive(self, setup_test_environment_with_interfaces):
        """TEST-INT-007: Verify deletion without archiving."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        manager.delete_and_archive_interfaces(archive=False)

        rm_folder = tmp_path / "RM_Collaborateurs"
        archived_folder = tmp_path / "Archived"
        deleted_folder = tmp_path / "Deleted"

        # RM folder should be deleted or empty
        assert not rm_folder.exists() or len(list(rm_folder.glob("*.xlsx"))) == 0

        # Archive should be empty
        assert len(list(archived_folder.glob("*.zip"))) == 0

        # Deleted should have a zip file
        assert len(list(deleted_folder.glob("*.zip"))) > 0

    def test_delete_empty_folder(self, setup_test_environment):
        """Verify handling when RM_Collaborateurs is empty."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        # RM_Collaborateurs exists but is empty
        manager.delete_and_archive_interfaces(archive=True)

        # Should complete without error
        rm_folder = tmp_path / "RM_Collaborateurs"
        assert not rm_folder.exists()

    def test_delete_nonexistent_folder(self, tmp_path):
        """Verify handling when RM_Collaborateurs doesn't exist."""
        (tmp_path / "Synthèse_RM_CE.xlsm").touch()
        (tmp_path / "RM_template.xlsx").touch()

        manager = RoadmapManager(tmp_path)

        # Should complete without error
        manager.delete_and_archive_interfaces(archive=True)


class TestDeleteMissingCollaborators:
    """Tests for cleanup of missing collaborators."""

    def test_delete_missing_collaborators(self, setup_test_environment_with_interfaces):
        """TEST-INT-008: Verify cleanup of orphaned interface files."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        rm_folder = tmp_path / "RM_Collaborateurs"

        # Create orphan file
        orphan_file = rm_folder / "RM_Orphan User.xlsx"
        wb = Workbook()
        wb.active.title = "POINTAGE"
        wb.save(orphan_file)
        wb.close()

        # Create new collabs.xml without Orphan
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>CLIGNIEZ Yann</collaborator>
            <collaborator>GANI Karim</collaborator>
            <collaborator>MOUHOUT Marouane</collaborator>
        </collaborators>"""
        (tmp_path / "collabs.xml").write_text(xml_content, encoding="utf-8")

        manager.delete_missing_collaborators()

        assert not orphan_file.exists()
        # Other files should still exist
        assert (rm_folder / "RM_CLIGNIEZ Yann.xlsx").exists()

    def test_delete_missing_no_orphans(self, setup_test_environment_with_interfaces):
        """Verify no action when all files match collaborators."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        rm_folder = tmp_path / "RM_Collaborateurs"
        before_count = len(list(rm_folder.glob("RM_*.xlsx")))

        # Create matching collabs.xml
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>CLIGNIEZ Yann</collaborator>
            <collaborator>GANI Karim</collaborator>
            <collaborator>MOUHOUT Marouane</collaborator>
        </collaborators>"""
        (tmp_path / "collabs.xml").write_text(xml_content, encoding="utf-8")

        manager.delete_missing_collaborators()

        after_count = len(list(rm_folder.glob("RM_*.xlsx")))
        assert before_count == after_count

    def test_delete_missing_creates_archive(self, setup_test_environment_with_interfaces):
        """Verify orphan files are archived before deletion."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        rm_folder = tmp_path / "RM_Collaborateurs"
        deleted_folder = tmp_path / "Deleted"

        # Create orphan file
        orphan_file = rm_folder / "RM_Orphan User.xlsx"
        wb = Workbook()
        wb.active.title = "POINTAGE"
        wb.save(orphan_file)
        wb.close()

        # Create new collabs.xml without Orphan
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>CLIGNIEZ Yann</collaborator>
            <collaborator>GANI Karim</collaborator>
            <collaborator>MOUHOUT Marouane</collaborator>
        </collaborators>"""
        (tmp_path / "collabs.xml").write_text(xml_content, encoding="utf-8")

        manager.delete_missing_collaborators()

        # Check archive was created
        archives = list(deleted_folder.glob("Deleted_Missing_*.zip"))
        assert len(archives) > 0


class TestPointage:
    """Tests for pointage export functionality."""

    def test_pointage_export(self, setup_test_environment_with_data):
        """TEST-INT-009: Verify pointage data export to XML."""
        tmp_path = setup_test_environment_with_data
        manager = RoadmapManager(tmp_path)

        result = manager.pointage()

        assert result is True

        xml_output = tmp_path / "pointage_output.xml"
        assert xml_output.exists()

        tree = ET.parse(xml_output)
        root = tree.getroot()
        rows = root.findall("row")

        # Should have 6 rows (2 per collaborator × 3 collaborators)
        assert len(rows) == 6

    def test_pointage_empty_interfaces(self, setup_test_environment_with_interfaces):
        """TEST-INT-010: Verify pointage export with empty interfaces."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        result = manager.pointage()

        assert result is False

        xml_output = tmp_path / "pointage_output.xml"
        assert xml_output.exists()

        tree = ET.parse(xml_output)
        root = tree.getroot()
        assert len(root.findall("row")) == 0

    def test_pointage_no_interfaces(self, setup_test_environment):
        """Verify pointage with no interface files."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        result = manager.pointage()

        assert result is False

    def test_pointage_skips_temp_files(self, setup_test_environment_with_data):
        """Verify pointage skips temporary Excel files (~$)."""
        tmp_path = setup_test_environment_with_data
        manager = RoadmapManager(tmp_path)

        rm_folder = tmp_path / "RM_Collaborateurs"

        # Create a temp file
        temp_file = rm_folder / "~$RM_CLIGNIEZ Yann.xlsx"
        temp_file.write_text("temp")

        result = manager.pointage()

        # Should still work
        assert result is True


class TestUpdateLc:
    """Tests for LC update functionality."""

    def test_update_lc(self, setup_test_environment_with_interfaces):
        """TEST-INT-011: Verify LC sheet update in all files."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        # Create LC.xlsx with test data
        wb = Workbook()
        ws = wb.active
        ws.title = "LC"
        ws["B2"] = "NewKey"
        ws["C2"] = "NewLabel"
        ws["D2"] = "NewFunc"
        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        manager.update_lc()

        # Verify template updated
        template_wb = load_workbook(tmp_path / "RM_template.xlsx")
        assert template_wb["LC"]["B2"].value == "NewKey"
        template_wb.close()

        # Verify interface files updated
        rm_folder = tmp_path / "RM_Collaborateurs"
        collab_wb = load_workbook(rm_folder / "RM_CLIGNIEZ Yann.xlsx")
        assert collab_wb["LC"]["B2"].value == "NewKey"
        collab_wb.close()

    def test_update_lc_no_file(self, setup_test_environment_with_interfaces):
        """Verify handling when LC.xlsx doesn't exist."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        # Should complete without error
        manager.update_lc()

    def test_update_lc_no_interfaces(self, setup_test_environment):
        """Verify LC update works with only template."""
        tmp_path = setup_test_environment
        manager = RoadmapManager(tmp_path)

        # Create LC.xlsx with test data
        wb = Workbook()
        ws = wb.active
        ws.title = "LC"
        ws["B2"] = "NewKey"
        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        manager.update_lc()

        # Verify template updated
        template_wb = load_workbook(tmp_path / "RM_template.xlsx")
        assert template_wb["LC"]["B2"].value == "NewKey"
        template_wb.close()

    def test_update_lc_recreates_data_validations(self, setup_test_environment_with_interfaces):
        """Verify data validations are recreated after LC update."""
        tmp_path = setup_test_environment_with_interfaces
        manager = RoadmapManager(tmp_path)

        # Create LC.xlsx with test data
        wb = Workbook()
        ws = wb.active
        ws.title = "LC"
        ws["B2"] = "NewKey"
        wb.save(tmp_path / "LC.xlsx")
        wb.close()

        manager.update_lc()

        # Verify interface files have data validations
        rm_folder = tmp_path / "RM_Collaborateurs"
        collab_wb = load_workbook(rm_folder / "RM_CLIGNIEZ Yann.xlsx")
        ws = collab_wb["POINTAGE"]

        assert len(ws.data_validations.dataValidation) == 4
        collab_wb.close()


class TestEdgeCases:
    """Tests for edge cases and error handling."""

    def test_all_ok_false_prevents_operations(self, tmp_path):
        """Verify operations are skipped when all_ok is False."""
        manager = RoadmapManager(tmp_path)

        assert manager.all_ok is False

        # These should return early without error
        manager.create_interfaces()
        manager.create_interfaces_fast()
        manager.delete_and_archive_interfaces(archive=True)
        manager.delete_missing_collaborators()
        result = manager.pointage()

        assert result is False

    def test_path_as_string(self, setup_test_environment):
        """Verify manager accepts path as string."""
        tmp_path = setup_test_environment

        manager = RoadmapManager(str(tmp_path))

        assert manager.all_ok is True
        assert manager.base_path == tmp_path

    def test_special_characters_in_collaborator_name(self, tmp_path):
        """Verify handling of special characters in collaborator names."""
        # Create required files
        template = Workbook()
        template.active.title = "POINTAGE"
        template.create_sheet("LC")
        template.save(tmp_path / "RM_template.xlsx")
        template.close()

        (tmp_path / "Synthèse_RM_CE.xlsm").touch()

        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <collaborators>
            <collaborator>NAZIH Imane</collaborator>
            <collaborator>YAHYA Oumaima</collaborator>
        </collaborators>"""
        (tmp_path / "collabs.xml").write_text(xml_content, encoding="utf-8")

        manager = RoadmapManager(tmp_path)
        manager.create_interfaces()

        rm_folder = tmp_path / "RM_Collaborateurs"
        assert (rm_folder / "RM_NAZIH Imane.xlsx").exists()
        assert (rm_folder / "RM_YAHYA Oumaima.xlsx").exists()
