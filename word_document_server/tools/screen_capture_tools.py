"""Screen capture tool for Microsoft Word documents via COM + Win32 API."""

import json
import os
import sys


def _capture_window_to_png(hwnd: int) -> bytes:
    """Capture a window using PrintWindow and return PNG bytes.

    Args:
        hwnd: Window handle (HWND).

    Returns:
        PNG image bytes.
    """
    import win32gui
    import win32ui
    from ctypes import windll
    from PIL import Image
    import io

    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]

    if width <= 0 or height <= 0:
        raise RuntimeError(f"Window has invalid dimensions: {width}x{height}")

    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(dcObj, width, height)
    cDC.SelectObject(bmp)

    # PW_RENDERFULLCONTENT = 2 — best quality on modern Windows
    result = windll.user32.PrintWindow(hwnd, cDC.GetSafeHdc(), 2)
    if not result:
        # Fallback to basic PrintWindow
        windll.user32.PrintWindow(hwnd, cDC.GetSafeHdc(), 0)

    bmpinfo = bmp.GetInfo()
    bmpstr = bmp.GetBitmapBits(True)
    img = Image.frombuffer(
        "RGB",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr,
        "raw",
        "BGRX",
        0,
        1,
    )

    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(bmp.GetHandle())

    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


async def word_screen_capture(filename: str = None, output_path: str = None) -> str:
    """Capture a screenshot of a Word document window.

    Args:
        filename: Document name to capture (None = active document).
        output_path: Where to save PNG. If omitted, saves to temp dir.

    Returns:
        JSON with path, dimensions, and document name.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Screen capture is only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        # Activate the document's window so it's visible
        doc.ActiveWindow.Activate()
        hwnd = int(doc.ActiveWindow.Hwnd)

        if not hwnd:
            return json.dumps({"error": "Could not get Word window handle"})

        png_bytes = _capture_window_to_png(hwnd)

        if not output_path:
            temp_dir = os.path.join(
                os.environ.get("TEMP", "/tmp"), "word_mcp_captures"
            )
            os.makedirs(temp_dir, exist_ok=True)
            safe_name = doc.Name.replace(".docx", "").replace(".doc", "")
            output_path = os.path.join(temp_dir, f"word_capture_{safe_name}.png")

        with open(output_path, "wb") as f:
            f.write(png_bytes)

        from PIL import Image
        import io

        img = Image.open(io.BytesIO(png_bytes))

        return json.dumps(
            {
                "success": True,
                "path": output_path,
                "width": img.width,
                "height": img.height,
                "document": doc.Name,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})
