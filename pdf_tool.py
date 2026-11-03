# -*- coding: utf-8 -*-
"""
PDF TOOL
========
Version : 1.0.0
Released: 2026-10-27

Converts Word / Excel / PowerPoint files and images to PDF, as separate PDFs
or combined into one PDF. Run it from the folder that holds your files
(use run_pdf_tool.bat, or:  python pdf_tool.py).

HOW TO ADD A NEW TOOL LATER
  1. Write a function:  def tool_something(**opts): ...
  2. Add ONE line to the MENU list near the bottom of this file.
  3. Raise VERSION and add an entry to CHANGELOG.
"""

import sys

if sys.version_info < (3, 8):
    print("Python 3.8 or newer is required. Please install a newer Python.")
    try:
        input("Press Enter to close...")
    except Exception:
        pass
    sys.exit(1)

import datetime
import importlib
import importlib.util
import os
import platform
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# VERSION INFORMATION  (raise these when you change the tool)
# ---------------------------------------------------------------------------
VERSION = "1.0.0"
RELEASE_DATE = "2026-10-27"
CHANGELOG = [
    ("1.0.0", "2026-10-27", [
        "First release.",
        "Word, Excel, PowerPoint and image files to PDF (separate or combined).",
        "Option 5: everything (all file types) combined into ONE PDF.",
        "order.txt is created automatically so the merge order can be edited.",
        "Microsoft Office engine with automatic LibreOffice fallback.",
        "System check that installs missing components (option 9).",
    ]),
]

# ---------------------------------------------------------------------------
# SETTINGS  (safe to edit)
# ---------------------------------------------------------------------------
PREFER_OFFICE = True          # True: Microsoft Office first, LibreOffice as fallback
                              # False: LibreOffice first (same output on every PC)
OUTPUT_DIR_NAME = "PDF_Output"
ORDER_FILE_NAME = "order.txt"
IMAGE_DPI = 200               # page resolution used for images
IMAGE_QUALITY = 90            # JPEG quality inside the PDF (1-100)
IMAGE_MARGIN_INCH = 0.3       # white border around each image
IMAGE_MAX_UPSCALE = 3.0       # small images are enlarged at most this much
A4_INCH = (8.27, 11.69)

# ---------------------------------------------------------------------------
# FILE TYPES
# ---------------------------------------------------------------------------
CATEGORIES = {
    "word": {
        "title": "Word",
        "exts": {".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm",
                 ".rtf", ".odt", ".txt"},
    },
    "excel": {
        "title": "Excel",
        "exts": {".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm",
                 ".csv", ".ods"},
    },
    "powerpoint": {
        "title": "PowerPoint",
        "exts": {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm",
                 ".pot", ".potx", ".potm", ".odp"},
    },
    "images": {
        "title": "Images",
        "exts": {".jpg", ".jpeg", ".jpe", ".jfif", ".png", ".gif", ".bmp",
                 ".dib", ".tif", ".tiff", ".webp", ".ico", ".heic", ".heif",
                 ".ppm", ".pgm", ".pbm", ".tga", ".jp2"},
    },
}
OFFICE_KINDS = ("word", "excel", "powerpoint")
OFFICE_PROGIDS = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "powerpoint": "PowerPoint.Application",
}
OOXML_EXTS = {".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xlsb",
              ".xltx", ".xltm", ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx",
              ".potm"}
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # old-style / encrypted Office file
DUMMY_PASSWORD = "pdf-tool-no-password"           # stops Office from showing a password box

IS_WINDOWS = os.name == "nt"

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")      # never crash on Arabic/odd file names
    except Exception:
        pass


# ---------------------------------------------------------------------------
# SMALL HELPERS
# ---------------------------------------------------------------------------
class EngineError(Exception):
    pass


def _input(prompt=""):
    try:
        return input(prompt)
    except EOFError:
        print()
        raise SystemExit(0)


def pause(text="\n Press Enter to return to the menu..."):
    _input(text)


def confirm(prompt, default=True):
    while True:
        answer = _input(prompt).strip().lower()
        if not answer:
            return default
        if answer in ("y", "yes"):
            return True
        if answer in ("n", "no"):
            return False
        print(" Please type Y or N.")


def ask_choice(prompt, valid):
    while True:
        answer = _input(prompt).strip()
        if answer in valid:
            return answer
        print(" Not a valid choice, try again.")


def natural_key(text):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def has_module(name):
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def category_of(path):
    ext = Path(path).suffix.lower()
    for key, info in CATEGORIES.items():
        if ext in info["exts"]:
            return key
    return None


def timestamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def com_message(exc):
    """Readable text from a COM error."""
    try:
        info = getattr(exc, "excepinfo", None)
        if info and len(info) > 2 and info[2]:
            return str(info[2]).strip()
    except Exception:
        pass
    return str(exc)


# ---------------------------------------------------------------------------
# FINDING FILES + order.txt
# ---------------------------------------------------------------------------
def scan_files(kinds):
    """Supported files in the current folder, in natural name order."""
    found = []
    for p in Path.cwd().iterdir():
        try:
            if not p.is_file():
                continue
        except OSError:
            continue
        name = p.name
        if name.startswith("~$") or name.startswith("."):
            continue
        if name.lower() == ORDER_FILE_NAME:
            continue
        if category_of(p) in kinds:
            found.append(p)
    found.sort(key=lambda p: natural_key(p.name))
    return found


def build_order(files):
    """Create order.txt, let the user edit it, and return the files in that order."""
    order_path = Path.cwd() / ORDER_FILE_NAME
    lines = [
        "# PDF TOOL - merge order. One file name per line, top = first page(s).",
        "# Reorder the lines, or delete a line to skip that file.",
        "# Lines starting with # are ignored. Save the file, then go back to the tool.",
    ] + [p.name for p in files]
    with open(order_path, "w", encoding="utf-8-sig", newline="\r\n") as fh:
        fh.write("\n".join(lines) + "\n")

    print("\n order.txt was created with %d file name(s)." % len(files))
    if hasattr(os, "startfile"):
        try:
            os.startfile(str(order_path))
            print(" It was opened in your text editor.")
        except Exception:
            pass
    print(" Reorder the lines if you want, SAVE the file, then come back here.")
    _input("\n Press Enter to continue...")
    return read_order(order_path)


def read_order(order_path):
    everything = scan_files(set(CATEGORIES))
    by_name = {p.name.lower(): p for p in everything}
    try:
        text = order_path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = order_path.read_text(encoding="cp1252", errors="replace")
    ordered, unknown = [], []
    for raw in text.splitlines():
        line = raw.strip().strip('"')
        if not line or line.startswith("#"):
            continue
        match = by_name.get(line.lower())
        if match:
            ordered.append(match)
        else:
            unknown.append(line)
    for line in unknown:
        print(" Warning: '%s' is not a supported file in this folder - skipped." % line)
    if not ordered:
        print("\n order.txt does not list any valid file. Nothing to do.")
    return ordered


# ---------------------------------------------------------------------------
# SYSTEM CHECK
# ---------------------------------------------------------------------------
def office_installed(kind):
    if not IS_WINDOWS:
        return False
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, OFFICE_PROGIDS[kind] + "\\CLSID")
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def find_soffice():
    """Path to LibreOffice's soffice, or None."""
    candidates = []
    for base in (Path.cwd(), Path(__file__).resolve().parent):
        candidates += [
            base / "LibreOfficePortable" / "App" / "libreoffice" / "program" / "soffice.exe",
            base / "LibreOffice" / "program" / "soffice.exe",
        ]
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        root = os.environ.get(env)
        if root:
            candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(Path(local) / "Programs" / "LibreOffice" / "program" / "soffice.exe")
    candidates.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    for c in candidates:
        try:
            if c.is_file():
                return str(c)
        except OSError:
            pass
    return shutil.which("soffice") or shutil.which("libreoffice")


def get_status():
    return {
        "pillow": has_module("PIL"),
        "pypdf": has_module("pypdf"),
        "pywin32": has_module("win32com"),
        "heif": has_module("pillow_heif"),
        "office": {k: office_installed(k) for k in OFFICE_KINDS},
        "soffice": find_soffice(),
    }


def engines_available(st, kind):
    """Engine names for this file kind, best first."""
    names = []
    if st["office"].get(kind) and st["pywin32"]:
        names.append("Microsoft Office")
    if st["soffice"]:
        names.append("LibreOffice")
    if not PREFER_OFFICE:
        names.reverse()
    return names


def needed_fixes(st):
    """(pip packages that must be installed, whether LibreOffice must be installed)"""
    pip = []
    if not st["pillow"]:
        pip.append("Pillow")
    if not st["pypdf"]:
        pip.append("pypdf")
    if IS_WINDOWS and any(st["office"].values()) and not st["pywin32"]:
        pip.append("pywin32")
    need_lo = False
    if not st["soffice"]:
        for kind in OFFICE_KINDS:
            if not st["office"][kind]:
                need_lo = True
    return pip, need_lo


def _row(label, value):
    print("  " + (label + " ").ljust(24, ".") + " " + value)


def print_status(st):
    print("\n System check")
    _row("Python", platform.python_version())
    _row("Pillow (images)", "OK" if st["pillow"] else "MISSING")
    _row("pypdf (combine)", "OK" if st["pypdf"] else "MISSING")
    if IS_WINDOWS and any(st["office"].values()):
        _row("pywin32 (Office)", "OK" if st["pywin32"] else "MISSING")
    else:
        _row("pywin32 (Office)", "not needed")
    _row("HEIC photos (optional)", "OK" if st["heif"] else "not installed")
    for kind in OFFICE_KINDS:
        names = engines_available(st, kind)
        if names:
            text = names[0] + (" (fallback: %s)" % names[1] if len(names) > 1 else "")
        else:
            text = "NOT AVAILABLE"
        if st["office"][kind] and not st["pywin32"]:
            text += "  [Office found, needs pywin32]"
        _row(CATEGORIES[kind]["title"], text)
    _row("Images", "built-in (no Office needed)")
    _row("LibreOffice", st["soffice"] or "NOT FOUND")


# ---------------------------------------------------------------------------
# INSTALLING MISSING COMPONENTS
# ---------------------------------------------------------------------------
def pip_install(packages):
    print("\n Installing: %s" % ", ".join(packages))
    if not has_module("pip"):
        subprocess.call([sys.executable, "-m", "ensurepip", "--upgrade"])
    base = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check"]
    for extra in ([], ["--user"]):
        try:
            code = subprocess.call(base + extra + list(packages))
        except Exception as exc:
            print(" pip could not be started: %s" % exc)
            return False
        if code == 0:
            importlib.invalidate_caches()
            return True
    print("\n Could not install: %s" % ", ".join(packages))
    print(" Check the internet connection and try again.")
    return False


def _latest_libreoffice_version():
    import urllib.request
    url = "https://download.documentfoundation.org/libreoffice/stable/"
    with urllib.request.urlopen(url, timeout=30) as resp:
        html = resp.read().decode("utf-8", "replace")
    versions = re.findall(r'href="(\d+\.\d+\.\d+)/"', html)
    if not versions:
        raise RuntimeError("no version found")
    return max(versions, key=lambda v: tuple(int(x) for x in v.split(".")))


def _install_libreoffice_msi():
    import urllib.request
    machine = platform.machine().lower()
    if machine in ("amd64", "x86_64"):
        folder, suffix = "x86_64", "x86-64"
    elif machine in ("arm64", "aarch64"):
        folder, suffix = "aarch64", "aarch64"
    else:
        print(" This PC type is not supported for automatic install.")
        return False
    try:
        version = _latest_libreoffice_version()
        url = ("https://download.documentfoundation.org/libreoffice/stable/%s/win/%s/"
               "LibreOffice_%s_Win_%s.msi" % (version, folder, version, suffix))
        msi = os.path.join(tempfile.gettempdir(), "LibreOffice_setup.msi")
        print(" Downloading LibreOffice %s (about 350 MB)..." % version)
        last = [-1]

        def progress(blocks, block_size, total):
            if total > 0:
                pct = min(100, int(blocks * block_size * 100 / total))
                if pct != last[0] and pct % 5 == 0:
                    last[0] = pct
                    print("   %d%%" % pct, end="\r", flush=True)

        urllib.request.urlretrieve(url, msi, progress)
        print("\n Installing (a Windows permission prompt will appear - please accept)...")
        ps = ("Start-Process msiexec.exe -ArgumentList '/i','\"%s\"','/qb','/norestore' "
              "-Verb RunAs -Wait" % msi)
        subprocess.call(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                         "-Command", ps])
        try:
            os.remove(msi)
        except OSError:
            pass
    except Exception as exc:
        print(" Download/install failed: %s" % exc)
        return False
    return find_soffice() is not None


def install_libreoffice():
    manual = "https://www.libreoffice.org/download/download-libreoffice/"
    if not IS_WINDOWS:
        print(" Automatic LibreOffice install is only supported on Windows.")
        print(" Download it from: " + manual)
        return False
    if shutil.which("winget"):
        print("\n Installing LibreOffice with winget (a permission prompt may appear)...")
        try:
            subprocess.call(["winget", "install", "--id", "TheDocumentFoundation.LibreOffice",
                             "-e", "--silent", "--accept-package-agreements",
                             "--accept-source-agreements"])
        except Exception as exc:
            print(" winget failed: %s" % exc)
        if find_soffice():
            print(" LibreOffice installed.")
            return True
        print(" winget did not finish the install. Trying a direct download...")
    if _install_libreoffice_msi():
        print(" LibreOffice installed.")
        return True
    print("\n Automatic install did not work. Install it manually from:\n  " + manual)
    return False


def restart_self():
    print("\n Restarting the tool to load the new components...\n")
    env = dict(os.environ, PDFTOOL_RESTARTED="1")
    code = subprocess.call([sys.executable] + sys.argv, env=env)
    raise SystemExit(code)


def fix_missing(st, include_optional=False):
    pip, need_lo = needed_fixes(st)
    optional = []
    if include_optional and not st["heif"]:
        optional.append("pillow-heif")
    if not pip and not need_lo and not optional:
        print("\n Nothing is missing.")
        return
    print("\n Missing components that can be installed:")
    if pip:
        print("  - Python packages: " + ", ".join(pip))
    if optional:
        print("  - Optional (HEIC photo support): pillow-heif")
    if need_lo:
        print("  - LibreOffice (about 350 MB) - used when Microsoft Office is not available")
    if not confirm("\n Install now? (Y/N) [Y]: "):
        return
    installed_packages = False
    if pip and pip_install(pip):
        installed_packages = True
    if optional and pip_install(optional):
        installed_packages = True
    if need_lo:
        install_libreoffice()
    if installed_packages:
        restart_self()
    print_status(get_status())


# ---------------------------------------------------------------------------
# ENGINES
# ---------------------------------------------------------------------------
class OfficeApp:
    """Drives Microsoft Word / Excel / PowerPoint through COM (Windows only)."""
    name = "Microsoft Office"

    def __init__(self, kind):
        self.kind = kind
        self.app = None
        self.we_started = False
        self.failed = None

    def _start(self):
        import win32com.client
        app = win32com.client.DispatchEx(OFFICE_PROGIDS[self.kind])
        try:
            if self.kind == "word":
                count = app.Documents.Count
            elif self.kind == "excel":
                count = app.Workbooks.Count
            else:
                count = app.Presentations.Count
        except Exception:
            count = 0
        # Only close the app at the end if nothing of the user's was open in it.
        self.we_started = (count == 0)
        try:
            if self.kind == "word":
                app.Visible = False
                app.DisplayAlerts = 0
            elif self.kind == "excel":
                app.Visible = False
                app.DisplayAlerts = False
            else:
                app.DisplayAlerts = 1
        except Exception:
            pass
        try:
            app.AutomationSecurity = 3      # never run macros
        except Exception:
            pass
        self.app = app

    def _alive(self):
        try:
            _ = self.app.Name
            return True
        except Exception:
            return False

    def convert(self, src, dst):
        if self.failed:
            raise EngineError(self.failed)
        if self.app is not None and not self._alive():
            self.app = None
        if self.app is None:
            try:
                self._start()
            except Exception as exc:
                self.failed = "Microsoft %s could not be started (%s)" % (
                    CATEGORIES[self.kind]["title"], com_message(exc))
                raise EngineError(self.failed)
        try:
            getattr(self, "_do_" + self.kind)(src, dst)
        except Exception as exc:
            raise EngineError(com_message(exc))

    def _do_word(self, src, dst):
        doc = None
        try:
            try:
                doc = self.app.Documents.Open(
                    FileName=str(src), ConfirmConversions=False, ReadOnly=True,
                    AddToRecentFiles=False, PasswordDocument=DUMMY_PASSWORD,
                    NoEncodingDialog=True)
            except TypeError:
                doc = self.app.Documents.Open(str(src))
            doc.ExportAsFixedFormat(str(dst), 17)          # 17 = PDF
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass

    def _do_excel(self, src, dst):
        wb = None
        try:
            try:
                wb = self.app.Workbooks.Open(
                    Filename=str(src), UpdateLinks=0, ReadOnly=True,
                    Password=DUMMY_PASSWORD)
            except TypeError:
                wb = self.app.Workbooks.Open(str(src))
            for ws in wb.Worksheets:                       # fit every sheet to page width
                try:
                    setup = ws.PageSetup
                    setup.Zoom = False
                    setup.FitToPagesWide = 1
                    setup.FitToPagesTall = False
                except Exception:
                    pass                                   # e.g. no printer installed
            wb.ExportAsFixedFormat(0, str(dst))            # 0 = PDF
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass

    def _do_powerpoint(self, src, dst):
        pres = None
        try:
            pres = self.app.Presentations.Open(str(src), True, False, False)
            pres.SaveAs(str(dst), 32)                      # 32 = PDF
        finally:
            if pres is not None:
                try:
                    pres.Close()
                except Exception:
                    pass

    def close(self):
        if self.app is not None and self.we_started:
            try:
                self.app.Quit()
            except Exception:
                pass
        self.app = None


class LibreOffice:
    """Converts with LibreOffice in headless mode (works on any PC)."""
    name = "LibreOffice"

    def __init__(self, exe):
        self.exe = exe
        self.profile = Path(tempfile.gettempdir()) / "pdf_tool_lo_profile"

    def convert(self, src, dst):
        tmp = Path(tempfile.mkdtemp(prefix="pdftool_lo_"))
        try:
            # Plain ASCII copy: avoids problems with Arabic names, long paths, lock files.
            inp = tmp / ("input" + Path(src).suffix.lower())
            shutil.copy2(str(src), str(inp))
            out = tmp / "out"
            out.mkdir()
            cmd = [self.exe,
                   "-env:UserInstallation=" + self.profile.as_uri(),
                   "--headless", "--norestore", "--nolockcheck", "--nodefault",
                   "--convert-to", "pdf", "--outdir", str(out), str(inp)]
            flags = subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0
            try:
                proc = subprocess.run(cmd, capture_output=True, timeout=300,
                                      creationflags=flags)
            except subprocess.TimeoutExpired:
                raise EngineError("LibreOffice took too long (over 5 minutes)")
            produced = list(out.glob("*.pdf"))
            if not produced:
                detail = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace").strip()
                raise EngineError("LibreOffice could not convert this file. " + detail[:200])
            shutil.move(str(produced[0]), str(dst))
        finally:
            shutil.rmtree(str(tmp), ignore_errors=True)

    def close(self):
        pass


class EngineManager:
    def __init__(self):
        self.st = get_status()
        self.office = {}
        self.lo = LibreOffice(self.st["soffice"]) if self.st["soffice"] else None

    def engines(self, kind):
        office = None
        if self.st["office"].get(kind) and self.st["pywin32"]:
            office = self.office.setdefault(kind, OfficeApp(kind))
        order = [office, self.lo] if PREFER_OFFICE else [self.lo, office]
        return [e for e in order if e is not None]

    def close(self):
        for app in self.office.values():
            app.close()


# ---------------------------------------------------------------------------
# CONVERTING
# ---------------------------------------------------------------------------
def looks_encrypted(path):
    if Path(path).suffix.lower() not in OOXML_EXTS:
        return False
    try:
        with open(path, "rb") as fh:
            return fh.read(8) == OLE_MAGIC
    except OSError:
        return False


def _to_rgb(img):
    from PIL import Image
    mode = img.mode
    if mode in ("RGBA", "LA", "PA") or (mode == "P" and "transparency" in img.info):
        rgba = img.convert("RGBA")
        background = Image.new("RGB", rgba.size, (255, 255, 255))
        background.paste(rgba, mask=rgba.split()[3])
        return background
    if mode in ("I;16", "I;16L", "I;16B", "I"):
        try:
            return img.point(lambda v: v * (1 / 256)).convert("L").convert("RGB")
        except Exception:
            return img.convert("L").convert("RGB")
    return img.convert("RGB")


def _fit_to_page(img):
    from PIL import Image
    width, height = int(A4_INCH[0] * IMAGE_DPI), int(A4_INCH[1] * IMAGE_DPI)
    if img.width > img.height:
        width, height = height, width                     # landscape page
    margin = int(IMAGE_MARGIN_INCH * IMAGE_DPI)
    scale = min((width - 2 * margin) / img.width, (height - 2 * margin) / img.height)
    scale = min(scale, IMAGE_MAX_UPSCALE)
    size = (max(1, round(img.width * scale)), max(1, round(img.height * scale)))
    resized = img.resize(size, Image.LANCZOS)
    page = Image.new("RGB", (width, height), (255, 255, 255))
    page.paste(resized, ((width - size[0]) // 2, (height - size[1]) // 2))
    return page


def image_to_pdf(src, dst):
    from PIL import Image, ImageOps, ImageSequence
    Image.MAX_IMAGE_PIXELS = None
    if has_module("pillow_heif"):
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except Exception:
            pass
    ext = Path(src).suffix.lower()
    with Image.open(str(src)) as im:
        if ext in (".tif", ".tiff"):                      # multi-page TIFF: every page
            frames = [ImageOps.exif_transpose(f) for f in ImageSequence.Iterator(im)]
        else:                                             # GIF/WebP: first frame only
            im.seek(0)
            frames = [ImageOps.exif_transpose(im)]
        pages = [_fit_to_page(_to_rgb(f)) for f in frames]
    pages[0].save(str(dst), "PDF", save_all=True, append_images=pages[1:],
                  resolution=float(IMAGE_DPI), quality=IMAGE_QUALITY)


def convert_one(src, dst, manager):
    """Convert one file to dst. Returns the name of the engine that did it."""
    kind = category_of(src)
    if kind == "images":
        image_to_pdf(src, dst)
        return "built-in"
    if looks_encrypted(src):
        raise EngineError("the file is password-protected")
    engines = manager.engines(kind)
    if not engines:
        raise EngineError("no Office or LibreOffice engine is available for %s files "
                          "- choose option 9 (System check)" % CATEGORIES[kind]["title"])
    errors = []
    for engine in engines:
        try:
            if Path(dst).exists():
                Path(dst).unlink()
            engine.convert(src, dst)
            if Path(dst).exists() and Path(dst).stat().st_size > 0:
                return engine.name
            errors.append("%s: no PDF was produced" % engine.name)
        except Exception as exc:
            errors.append("%s: %s" % (engine.name, str(exc) or exc.__class__.__name__))
    raise EngineError(" | ".join(errors))


def merge_pdfs(items, dst):
    """items: list of (pdf_path, bookmark_title)"""
    from pypdf import PdfWriter
    writer = PdfWriter()
    for pdf, title in items:
        try:
            writer.append(str(pdf), outline_item=title)
        except TypeError:
            writer.append(str(pdf))
    with open(str(dst), "wb") as fh:
        writer.write(fh)
    try:
        writer.close()
    except Exception:
        pass


def place_pdf(tmp_pdf, out_dir, src, used):
    """Move a finished PDF into the output folder with a unique, safe name."""
    stem = Path(src).stem
    name = stem + ".pdf"
    if name.lower() in used:
        name = "%s_%s.pdf" % (stem, Path(src).suffix.lstrip(".").lower())
    base, n = Path(name).stem, 2
    while name.lower() in used:
        name = "%s_%d.pdf" % (base, n)
        n += 1
    used.add(name.lower())
    target = out_dir / name
    try:
        if target.exists():
            target.unlink()
        shutil.move(str(tmp_pdf), str(target))
    except PermissionError:                              # old PDF is open in a viewer
        target = out_dir / ("%s_%s.pdf" % (Path(name).stem, timestamp()))
        shutil.move(str(tmp_pdf), str(target))
    return target


def write_log(out_dir, label, mode, results, outputs):
    try:
        with open(str(out_dir / "pdf_tool_log.txt"), "a", encoding="utf-8") as fh:
            fh.write("[%s] PDF Tool v%s - %s / %s\n" % (
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), VERSION, label, mode))
            for src, pdf, engine, err in results:
                if pdf:
                    fh.write("   OK     %s  (%s)\n" % (src.name, engine))
                else:
                    fh.write("   FAILED %s  -> %s\n" % (src.name, err))
            for o in outputs:
                fh.write("   OUTPUT %s\n" % o)
            fh.write("\n")
    except OSError:
        pass


def run_conversion(files, mode, label):
    out_dir = Path.cwd() / OUTPUT_DIR_NAME
    out_dir.mkdir(exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="pdftool_"))
    manager = EngineManager()
    results, outputs = [], []
    print()
    try:
        for i, src in enumerate(files, 1):
            print(" [%d/%d] %s ... " % (i, len(files), src.name), end="", flush=True)
            dst = work / ("%04d.pdf" % i)
            try:
                engine = convert_one(src, dst, manager)
                print("OK (%s)" % engine)
                results.append((src, dst, engine, None))
            except Exception as exc:
                message = str(exc) or exc.__class__.__name__
                print("FAILED")
                print("        -> " + message)
                results.append((src, None, None, message))

        ok = [r for r in results if r[1]]
        failed = [r for r in results if not r[1]]
        if not ok:
            print("\n Nothing was converted.")
        else:
            go = True
            if failed and mode == "combine":
                go = confirm("\n %d file(s) failed. Combine the other %d anyway? (Y/N) [Y]: "
                             % (len(failed), len(ok)))
            if go and mode == "combine":
                target = out_dir / ("Combined_%s_%s.pdf" % (label, timestamp()))
                try:
                    merge_pdfs([(r[1], r[0].stem) for r in ok], target)
                    outputs.append(target)
                except Exception as exc:
                    print("\n Could not combine the PDFs: %s" % exc)
            elif go:
                used = set()
                for src, pdf, _engine, _err in ok:
                    outputs.append(place_pdf(pdf, out_dir, src, used))

        print("\n ---------------- Summary ----------------")
        print("  Converted : %d" % len([r for r in results if r[1]]))
        print("  Failed    : %d" % len(failed))
        if outputs:
            if len(outputs) == 1:
                print("  Output    : %s" % outputs[0])
            else:
                print("  Output    : %d PDF files in %s" % (len(outputs), out_dir))
        write_log(out_dir, label, mode, results, outputs)
    finally:
        manager.close()
        shutil.rmtree(str(work), ignore_errors=True)


# ---------------------------------------------------------------------------
# TOOLS (menu entries)
# ---------------------------------------------------------------------------
def require_modules(*modules):
    pretty = {"PIL": "Pillow", "pypdf": "pypdf"}
    missing = [pretty.get(m, m) for m in modules if not has_module(m)]
    if missing:
        print("\n Missing Python components: %s" % ", ".join(missing))
        print(" Choose option 9 (System check) to install them.")
        pause()
        return False
    return True


def tool_convert(kinds, label, combine_only=False):
    if not require_modules("PIL", "pypdf"):
        return
    files = scan_files(set(kinds))
    if not files:
        print("\n No %s files found in this folder." % label)
        pause()
        return
    print("\n Found %d %s file(s)." % (len(files), label))
    if combine_only:
        mode = "combine"
    else:
        print("\n Output type?")
        print("   1 - Combine all into ONE PDF")
        print("   2 - Separate PDF for each file")
        print("   0 - Back")
        choice = ask_choice(" Press a number: ", {"0", "1", "2"})
        if choice == "0":
            return
        mode = "combine" if choice == "1" else "separate"
    if mode == "combine":
        files = build_order(files)
        if not files:
            pause()
            return
    run_conversion(files, mode, label)
    pause()


def tool_about():
    print("\n PDF Tool  v%s  (released %s)" % (VERSION, RELEASE_DATE))
    print(" Working folder: %s" % Path.cwd())
    print("\n Version history")
    for version, date, items in CHANGELOG:
        print("  v%s  (%s)" % (version, date))
        for item in items:
            print("     - " + item)
    pause()


def tool_system_check():
    st = get_status()
    print_status(st)
    fix_missing(st, include_optional=True)
    pause()


# To add a tool: add ONE line here. Keys 8, 9 and 0 are reserved for About,
# System check and Exit. Use 6, 7, 10, 11, ... for new tools.
MENU = [
    ("1", "Word        -> PDF", tool_convert, {"kinds": ["word"], "label": "Word"}),
    ("2", "Images      -> PDF", tool_convert, {"kinds": ["images"], "label": "Images"}),
    ("3", "Excel       -> PDF", tool_convert, {"kinds": ["excel"], "label": "Excel"}),
    ("4", "PowerPoint  -> PDF", tool_convert, {"kinds": ["powerpoint"], "label": "PowerPoint"}),
    ("5", "EVERYTHING  -> ONE PDF (Word + Excel + PowerPoint + images)", tool_convert,
     {"kinds": list(CATEGORIES), "label": "All", "combine_only": True}),
]


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def print_banner():
    line = "=" * 62
    print(line)
    print("  PDF TOOL   v%s   (%s)" % (VERSION, RELEASE_DATE))
    print("  Folder: %s" % Path.cwd())
    print(line)


def startup_check():
    st = get_status()
    print_status(st)
    pip, need_lo = needed_fixes(st)
    if not pip and not need_lo:
        return
    print("\n Some components are missing.")
    if os.environ.get("PDFTOOL_RESTARTED"):
        print(" They are still missing after the install. Choose option 9 for details.")
        return
    fix_missing(st)


def main():
    if "--version" in sys.argv:
        print(VERSION)
        return
    print_banner()
    startup_check()
    handlers = {key: (func, opts) for key, _title, func, opts in MENU}
    while True:
        print("\n What do you want to do?")
        for key, title, _func, opts in MENU:
            extra = ""
            if "kinds" in opts:
                count = len(scan_files(set(opts["kinds"])))
                extra = "   [%d file%s found]" % (count, "" if count == 1 else "s")
            print("   %s - %s%s" % (key, title, extra))
        print("   8 - About / version history")
        print("   9 - System check / install missing components")
        print("   0 - Exit")
        choice = _input("\n Press a number: ").strip()
        if choice == "0":
            print(" Goodbye.")
            return
        try:
            if choice == "8":
                tool_about()
            elif choice == "9":
                tool_system_check()
            elif choice in handlers:
                func, opts = handlers[choice]
                func(**opts)
            else:
                print(" Not a valid choice.")
        except KeyboardInterrupt:
            print("\n Cancelled.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n Cancelled.")
