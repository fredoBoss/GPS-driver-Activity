"""Reach a workbook that is open in Excel (Windows + pywin32), so imports work without closing it.

While Excel has a workbook open, no other program may save over the file. Instead the
importer asks that Excel to type the values in and save. The open workbook is found through
the Windows Running Object Table, so this never opens a file or starts Excel by itself.
"""
import os
import threading
import traceback

from .excel_log import ActivityLogError

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError:  # not Windows, or pywin32 not installed
    pythoncom = None

RPC_E_CALL_REJECTED = -2147418111  # Excel is busy: a cell is being edited or a dialog is open
NOT_OPEN = object()


def call_with_open_workbook(path, func):
    """Return func(workbook) for the Excel workbook open at `path`, or NOT_OPEN if Excel does not have it.

    The COM work runs in its own thread with its own CoInitialize/CoUninitialize, so it never
    disconnects COM objects the calling thread holds. `func` must return plain Python values,
    and only messages (no tracebacks holding COM objects) cross back to the caller.
    """
    if pythoncom is None:
        return NOT_OPEN
    outcome = {}

    def run():
        pythoncom.CoInitialize()
        workbook = None
        try:
            workbook = _find_open_workbook(path)
            outcome['value'] = NOT_OPEN if workbook is None else func(workbook)
        except ActivityLogError as exc:
            outcome['error'] = ActivityLogError(str(exc))
        except pywintypes.com_error as exc:
            outcome['error'] = ActivityLogError(_describe(exc, path))
        except Exception:
            outcome['error'] = RuntimeError(traceback.format_exc())
        finally:
            workbook = None  # release the COM object before uninitializing
            pythoncom.CoUninitialize()

    worker = threading.Thread(target=run, name='excel-import')
    worker.start()
    worker.join()
    if 'error' in outcome:
        raise outcome['error']
    return outcome['value']


def _find_open_workbook(path):
    target = os.path.normcase(os.path.abspath(path))
    rot = pythoncom.GetRunningObjectTable()
    ctx = pythoncom.CreateBindCtx(0)
    for moniker in rot:
        try:
            name = moniker.GetDisplayName(ctx, None)
        except pythoncom.com_error:
            continue
        if os.path.normcase(name) == target:
            obj = rot.GetObject(moniker)
            return win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return None


def _describe(exc, path):
    if exc.hresult == RPC_E_CALL_REJECTED:
        return (f'Excel is busy, so {path.name} could not be updated. If you are typing in a cell, '
                'press Enter or Esc; close any open dialog box in Excel, then import again.')
    detail = exc.excepinfo[2] if exc.excepinfo and exc.excepinfo[2] else exc.strerror
    return f'Excel could not update {path.name}: {detail}'
