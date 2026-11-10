"""PyScript bridge. Workbooks are processed locally in the browser."""
import base64
import json
from pyscript import window, ffi
from converter import convert_bytes

async def convert_uploaded_file(file):
    raw = await file.arrayBuffer()
    source = bytes(window.Uint8Array.new(raw).to_py())
    output, file_format, sheet_count = convert_bytes(source, str(file.name))
    return json.dumps({
        "base64": base64.b64encode(output).decode("ascii"),
        "format": file_format,
        "sheets": sheet_count,
        "bytes": len(output),
    })

window.convertWithPython = ffi.create_proxy(convert_uploaded_file)
window.dispatchEvent(window.Event.new("python-converter-ready"))
