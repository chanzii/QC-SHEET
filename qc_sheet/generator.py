"""스펙, 양식, 서명으로 QC시트 바이트를 생성한다."""
import re
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image


class StyleNotFound(Exception):
    pass


class SizeNotFound(Exception):
    pass


class NoMeasurements(Exception):
    pass


def generate_qc_sheet(spec_path, template_path, logo_path, style_number, selected_size,
                      language_choice, use_first_sheet=False):
    wb_spec = load_workbook(spec_path, data_only=True, read_only=True)
    try:
        ws_spec = next((ws for ws in wb_spec.worksheets
                        if ws["A1"].value and style_number.upper() in str(ws["A1"].value).upper()), None)
        if ws_spec is None:
            if not use_first_sheet:
                raise StyleNotFound()
            ws_spec = wb_spec.worksheets[0]

        wb_tpl = load_workbook(template_path)
        try:
            ws_tpl = wb_tpl.active
            ws_tpl["B6"], ws_tpl["G6"] = style_number, selected_size
            ws_tpl.add_image(Image(BytesIO(Path(logo_path).read_bytes())), "F2")
            header = list(ws_spec.iter_rows(min_row=2, max_row=2, values_only=True))[0]
            size_map = {str(v).strip(): idx for idx, v in enumerate(header) if v}
            if selected_size not in size_map:
                raise SizeNotFound()
            idx = size_map[selected_size]
            rows = list(ws_spec.iter_rows(min_row=3, values_only=True))
            data, i = [], 0
            while i < len(rows):
                part = str(rows[i][1]).strip() if rows[i][1] else ""
                val = rows[i][idx]
                if language_choice == "English":
                    if re.search(r"[A-Za-z]", part) and val is not None:
                        data.append((part, val)); i += 1; continue
                else:
                    if re.search(r"[A-Za-z]", part) and val is not None and i + 1 < len(rows):
                        kr = str(rows[i+1][1]).strip() if rows[i+1][1] else ""
                        if re.search(r"[가-힣]", kr):
                            data.append((kr, val)); i += 2; continue
                    if re.search(r"[가-힣]", part) and val is not None:
                        data.append((part, val)); i += 1; continue
                i += 1
            if not data:
                raise NoMeasurements()
            for j, (part, value) in enumerate(data):
                row = 9 + j
                ws_tpl.cell(row, 1, part)
                ws_tpl.cell(row, 2, value)
                ws_tpl.cell(row, 4, f'=IF(C{row}="","",IFERROR(C{row}-B{row},""))')
            out = BytesIO()
            wb_tpl.save(out)
            return out.getvalue()
        finally:
            wb_tpl.close()
    finally:
        wb_spec.close()
