import io
from pathlib import Path

from openpyxl import Workbook, load_workbook

from roadmap.helpers import build_interface, get_collaborators


def test_get_collaborators(tmp_path):
    file_p = tmp_path / "synthese.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Gestion_Interfaces"

    ws["B3"] = "Alice"
    ws["B4"] = "Bob"
    ws["B5"] = None  # stop here

    wb.save(file_p)

    collabs = get_collaborators(file_p)
    assert collabs == ["Alice", "Bob"]


def test_build_interface(tmp_path):
    # Create fake template in memory
    wb = Workbook()
    ws_pointage = wb.active
    ws_pointage.title = "POINTAGE"

    ws_lc = wb.create_sheet("LC")

    buffer = io.BytesIO()
    wb.save(buffer)

    output = tmp_path / "RM_Test.xlsx"

    build_interface(buffer.getvalue(), output, "TEST_USER")

    assert output.exists()
    generated = load_workbook(output)
    assert generated["POINTAGE"]["B1"].value == "TEST_USER"
