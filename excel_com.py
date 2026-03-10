from __future__ import annotations

from pathlib import Path
import contextlib

import pythoncom
from win32com.client import Dispatch, DispatchEx, dynamic

XL_TYPE_PDF = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class ExcelConversionError(RuntimeError):
    """Raised when Excel COM automation cannot complete a conversion."""


def _create_excel_application():
    factories = (
        ("dynamic.Dispatch", lambda: dynamic.Dispatch("Excel.Application")),
        ("DispatchEx", lambda: DispatchEx("Excel.Application")),
        ("Dispatch", lambda: Dispatch("Excel.Application")),
    )
    errors: list[str] = []

    for label, factory in factories:
        try:
            return factory()
        except Exception as exc:  # pragma: no cover - depends on local Excel install
            errors.append(f"{label}: {exc}")

    raise ExcelConversionError(
        "Unable to start Excel via COM. Tried: " + " | ".join(errors)
    )



def convert_excel_to_pdf(input_file: str | Path, output_file: str | Path) -> Path:
    input_path = Path(input_file).resolve()
    output_path = Path(output_file).resolve()
    excel = None
    workbook = None

    pythoncom.CoInitialize()
    try:
        excel = _create_excel_application()

        # All property sets are optional — suppress failures so a broken or
        # restricted COM wrapper (e.g. different Excel version on a remote
        # machine) does not prevent the actual conversion.
        for prop, val in (
            ("Visible", False),
            ("DisplayAlerts", False),
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
            ("EnableEvents", False),
            ("ScreenUpdating", False),
        ):
            with contextlib.suppress(Exception):
                setattr(excel, prop, val)

        try:
            workbook = excel.Workbooks.Open(
                str(input_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
        except TypeError:
            workbook = excel.Workbooks.Open(
                str(input_path),
                UpdateLinks=0,
                ReadOnly=True,
            )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.ExportAsFixedFormat(XL_TYPE_PDF, str(output_path))
        return output_path
    except ExcelConversionError:
        raise
    except Exception as exc:
        raise ExcelConversionError(str(exc)) from exc
    finally:
        if workbook is not None:
            with contextlib.suppress(Exception):
                workbook.Close(SaveChanges=False)
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.Quit()
        pythoncom.CoUninitialize()
