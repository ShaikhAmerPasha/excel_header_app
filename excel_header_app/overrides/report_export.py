import datetime
import io
import json

import frappe
from frappe.utils import cint


def _today_formatted() -> str:
    return datetime.date.today().strftime("%d-%m-%Y")


@frappe.whitelist()
def export_query():
    form_params      = frappe.form_dict
    file_format_type = form_params.get("file_format_type", "Excel")
    report_name      = form_params.get("report_name", "")

    # CSV - fall back to Frappe core, we don't touch it
    if file_format_type != "Excel":
        from frappe.desk.query_report import export_query as _core
        return _core()

    # Report Builder - uses different export path
    report_doc = frappe.get_doc("Report", report_name)
    if report_doc.report_type == "Report Builder":
        from frappe.desk.reportview import export_query as rb_export
        return rb_export()

    # Let Frappe generate the xlsx normally first
    from frappe.desk.query_report import export_query as _core
    _core()

    # Now grab the generated file content and inject date at the top
    filecontent = frappe.response.get("filecontent")
    if not filecontent:
        return

    # Open the xlsx Frappe generated
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    wb = load_workbook(filename=io.BytesIO(filecontent))
    ws = wb.active

    # Insert 2 rows at the very top
    ws.insert_rows(1, amount=2)

    # Get number of columns in the sheet
    num_cols = ws.max_column

    # Row 1 - Date
    ws["A1"] = f"Downloaded On: {_today_formatted()}"
    ws.merge_cells(f"A1:{get_column_letter(num_cols)}1")
    ws["A1"].font      = Font(name="Calibri", italic=True, size=10)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 16

    # Row 2 - Blank spacer
    ws["A2"] = ""
    ws.row_dimensions[2].height = 8

    # Save back
    buffer = io.BytesIO()
    wb.save(buffer)

    frappe.response["filecontent"] = buffer.getvalue()