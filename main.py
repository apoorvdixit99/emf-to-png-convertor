# main.py
# Windows-only EMF → PNG conversion utilities
# Uses Win32 GDI via ctypes + pywin32

import ctypes
from ctypes import wintypes
from PIL import Image
import win32ui
import base64
import io
import tempfile
import os

# Globals for Win32 DLLs
gdi32 = None
user32 = None

URI_PREFIX_PNG = "data:image/png;base64,"
URI_PREFIX_EMF = "data:image/x-emf;base64,"
DPI = 300
DEFAULT_OUTPUT_FILENAME = "output.png"

# ----------------------------
# Win32 Structures
# ----------------------------

class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]


class SIZEL(ctypes.Structure):
    _fields_ = [
        ("cx", wintypes.LONG),
        ("cy", wintypes.LONG),
    ]


class ENHMETAHEADER(ctypes.Structure):
    _fields_ = [
        ("iType", wintypes.DWORD),
        ("nSize", wintypes.DWORD),
        ("rclBounds", RECT),        # pixels (recording DC)
        ("rclFrame", RECT),         # .01 mm units
        ("dSignature", wintypes.DWORD),
        ("nVersion", wintypes.DWORD),
        ("nBytes", wintypes.DWORD),
        ("nRecords", wintypes.DWORD),
        ("nHandles", wintypes.WORD),
        ("sReserved", wintypes.WORD),
        ("nDescription", wintypes.DWORD),
        ("offDescription", wintypes.DWORD),
        ("nPalEntries", wintypes.DWORD),
        ("szlDevice", SIZEL),       # pixels of reference device
        ("szlMillimeters", SIZEL),  # mm of reference device
        ("cbPixelFormat", wintypes.DWORD),
        ("offPixelFormat", wintypes.DWORD),
        ("bOpenGL", wintypes.DWORD),
        ("szlMicrometers", SIZEL),  # not always populated
    ]


# ----------------------------
# Converter Class
# ----------------------------

class EMFToPNGConverter:

    def __init__(self):
        self._initialize_win32()

    def _initialize_win32(self):
        global gdi32, user32

        gdi32 = ctypes.windll.gdi32
        user32 = ctypes.windll.user32

        gdi32.GetEnhMetaFileW.argtypes = [wintypes.LPCWSTR]
        gdi32.GetEnhMetaFileW.restype = wintypes.HANDLE

        gdi32.GetEnhMetaFileHeader.argtypes = [
            wintypes.HANDLE,
            wintypes.UINT,
            ctypes.c_void_p,
        ]
        gdi32.GetEnhMetaFileHeader.restype = wintypes.UINT

        gdi32.DeleteEnhMetaFile.argtypes = [wintypes.HANDLE]
        gdi32.DeleteEnhMetaFile.restype = wintypes.BOOL

        gdi32.PlayEnhMetaFile.argtypes = [
            wintypes.HDC,
            wintypes.HANDLE,
            ctypes.POINTER(RECT),
        ]
        gdi32.PlayEnhMetaFile.restype = wintypes.BOOL

    # ----------------------------
    # EMF Dimension Extraction
    # ----------------------------

    def read_emf_bbox(self, emf_path, dpi=DPI):
        hemf = gdi32.GetEnhMetaFileW(emf_path)
        if not hemf:
            raise OSError("Failed to open EMF")

        try:
            hdr = ENHMETAHEADER()
            ok = gdi32.GetEnhMetaFileHeader(
                hemf,
                ctypes.sizeof(hdr),
                ctypes.byref(hdr),
            )
            if ok == 0:
                raise OSError("GetEnhMetaFileHeader failed")

            # Pixel bounds
            bw = hdr.rclBounds.right - hdr.rclBounds.left
            bh = hdr.rclBounds.bottom - hdr.rclBounds.top

            # Frame in 0.01 mm → inches
            fw_01mm = hdr.rclFrame.right - hdr.rclFrame.left
            fh_01mm = hdr.rclFrame.bottom - hdr.rclFrame.top

            fw_mm = fw_01mm / 100.0
            fh_mm = fh_01mm / 100.0
            fw_in = fw_mm / 25.4
            fh_in = fh_mm / 25.4

            fw_px = int(round(fw_in * dpi))
            fh_px = int(round(fh_in * dpi))

            return {
                "bounds_pixels": {"width_px": bw, "height_px": bh},
                "frame_01mm": {"width_01mm": fw_01mm, "height_01mm": fh_01mm},
                "frame_mm": {"width_mm": fw_mm, "height_mm": fh_mm},
                "frame_inches": {"width_in": fw_in, "height_in": fh_in},
                "frame_pixels_at_dpi": {
                    "dpi": dpi,
                    "width_px": fw_px,
                    "height_px": fh_px,
                },
                "device_pixels_hint": {
                    "cx": hdr.szlDevice.cx,
                    "cy": hdr.szlDevice.cy,
                },
                "device_mm_hint": {
                    "cx_mm": hdr.szlMillimeters.cx,
                    "cy_mm": hdr.szlMillimeters.cy,
                },
            }

        finally:
            gdi32.DeleteEnhMetaFile(hemf)

    def get_dimensions(self, emf_path, dpi=DPI):
        info = self.read_emf_bbox(emf_path, dpi=dpi)
        dims = info["frame_pixels_at_dpi"]
        return dims["width_px"], dims["height_px"]

    # ----------------------------
    # EMF → PNG (file)
    # ----------------------------

    def emf_file_to_png_file(self, emf_path, png_path=DEFAULT_OUTPUT_FILENAME, dpi=DPI):
        width, height = self.get_dimensions(emf_path, dpi=dpi)

        hemf = gdi32.GetEnhMetaFileW(emf_path)
        if not hemf:
            raise OSError("Failed to load EMF")

        hdc_screen = user32.GetDC(0)
        dc = win32ui.CreateDCFromHandle(hdc_screen)
        mem_dc = dc.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(dc, width, height)
        mem_dc.SelectObject(bmp)

        # White background
        mem_dc.FillSolidRect((0, 0, width, height), 0xFFFFFF)

        rect = RECT(0, 0, width, height)
        gdi32.PlayEnhMetaFile(mem_dc.GetSafeHdc(), hemf, ctypes.byref(rect))

        tmp_bmp = tempfile.NamedTemporaryFile(delete=False, suffix=".bmp")
        tmp_bmp.close()
        bmp.SaveBitmapFile(mem_dc, tmp_bmp.name)

        Image.open(tmp_bmp.name).save(png_path, "PNG")

        os.remove(tmp_bmp.name)
        gdi32.DeleteEnhMetaFile(hemf)
        mem_dc.DeleteDC()
        dc.DeleteDC()
        user32.ReleaseDC(0, hdc_screen)

    # ----------------------------
    # EMF bytes → PNG bytes
    # ----------------------------

    def emf_bytes_to_png_bytes(self, emf_bytes, dpi=DPI):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".emf") as emf_file:
            emf_file.write(emf_bytes)
            emf_path = emf_file.name

        width, height = self.get_dimensions(emf_path, dpi=dpi)

        hemf = gdi32.GetEnhMetaFileW(emf_path)
        if not hemf:
            os.remove(emf_path)
            raise OSError("Failed to load EMF")

        hdc_screen = user32.GetDC(0)
        dc = win32ui.CreateDCFromHandle(hdc_screen)
        mem_dc = dc.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(dc, width, height)
        mem_dc.SelectObject(bmp)

        mem_dc.FillSolidRect((0, 0, width, height), 0xFFFFFF)

        rect = RECT(0, 0, width, height)
        gdi32.PlayEnhMetaFile(mem_dc.GetSafeHdc(), hemf, ctypes.byref(rect))

        with tempfile.NamedTemporaryFile(delete=False, suffix=".bmp") as bmp_file:
            bmp_path = bmp_file.name

        bmp.SaveBitmapFile(mem_dc, bmp_path)

        with Image.open(bmp_path) as im:
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            png_bytes = buf.getvalue()

        os.remove(emf_path)
        os.remove(bmp_path)
        gdi32.DeleteEnhMetaFile(hemf)
        mem_dc.DeleteDC()
        dc.DeleteDC()
        user32.ReleaseDC(0, hdc_screen)

        return png_bytes

    # ----------------------------
    # Base64 / URI Helpers
    # ----------------------------

    def emf_file_to_emf_base64(self, emf_path):
        with open(emf_path, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")

    def emf_file_to_emf_uri(self, emf_path):
        return URI_PREFIX_EMF + self.emf_file_to_emf_base64(emf_path)

    def emf_uri_to_emf_base64(self, emf_uri):
        prefix = URI_PREFIX_EMF
        return emf_uri[len(prefix):] if emf_uri.startswith(prefix) else None

    def emf_base64_to_emf_uri(self, emf_base64):
        return URI_PREFIX_EMF + emf_base64

    def png_base64_to_png_uri(self, png_base64):
        return URI_PREFIX_PNG + png_base64

    def png_uri_to_png_base64(self, png_uri):
        prefix = URI_PREFIX_PNG
        return png_uri[len(prefix):] if png_uri.startswith(prefix) else None

    def emf_base64_to_png_base64(self, emf_base64, dpi=DPI):
        emf_bytes = base64.b64decode(emf_base64)
        png_bytes = self.emf_bytes_to_png_bytes(emf_bytes, dpi=dpi)
        return base64.b64encode(png_bytes).decode("ascii")

    def emf_uri_to_png_uri(self, emf_uri, dpi=DPI):
        emf_base64 = self.emf_uri_to_emf_base64(emf_uri)
        png_base64 = self.emf_base64_to_png_base64(emf_base64, dpi=dpi)
        return self.png_base64_to_png_uri(png_base64)


# ----------------------------
# CLI Entry
# ----------------------------

if __name__ == "__main__":
    conv = EMFToPNGConverter()
    conv.emf_file_to_png_file("sample/input.emf", "output/output.png")
