"""PyScript bridge. Workbook bytes remain within the current browser session."""
import base64
from pyscript import window, ffi
from pyodide.ffi import to_js
from converter import convert_bytes

async def convert_uploaded_file(file):
    raw = await file.arrayBuffer()
    data = bytes(window.Uint8Array.new(raw).to_py())
    result, kind, sheets = convert_bytes(data, str(file.name))
    return to_js({
        "base64": base64.b64encode(result).decode("ascii"),
        "format": kind,
        "sheets": sheets,
        "bytes": len(result)
    })

window.convertWithPython = ffi.create_proxy(convert_uploaded_file)
window.dispatchEvent(window.Event.new("python-converter-ready"))
