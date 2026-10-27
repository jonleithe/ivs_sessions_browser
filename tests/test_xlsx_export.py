import io

from openpyxl import load_workbook

from ivs_sessions_browser.xlsx_export import _selected_stations, sessions_to_xlsx_bytes


def test_selected_stations_keeps_station_status_and_spacing() -> None:
    assert _selected_stations("FtNnWzNs") == "Nn      Ns"
    assert _selected_stations("FtNnWz [NsHb]") == "Nn    [Ns]"
    assert _selected_stations("FtWz [NnNs]") == "[Nn]  [Ns]"
    assert _selected_stations("FtWzNs [Hb]") == "        Ns"
    assert _selected_stations("FtWz [Hb]") == ""


def test_xlsx_stations_column_contains_only_nn_and_ns() -> None:
    values = [
        "A", "R1", "R1234", "2026-10-27 12:00", "300", "24:00",
        "FtNnWz [NsHb]", "", "", "", "", "",
    ]
    rows = [(values, None, {})]

    workbook_bytes = sessions_to_xlsx_bytes(rows, ["STATIONS"], {}, {})
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=False)

    sheet = workbook["Sessions"]
    assert sheet["B5"].value == " Stations "
    assert sheet["B6"].value == "Nn    [Ns]"
    assert sheet.column_dimensions["B"].width == 10
    assert sheet["B5"].font.name == "Courier New"
    assert sheet["B6"].font.name == "Courier New"
