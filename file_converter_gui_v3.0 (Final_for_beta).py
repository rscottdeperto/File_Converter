
# file_converter_gui_v3.0.py


import os
import re
import csv
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
import sys

# -------------------------------
# Config / constants
# -------------------------------
# Dynamically set ASSET_DIR for script and PyInstaller .exe
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSET_DIR = os.path.join(BASE_DIR, "assets")

SUPPORTED_FORMATS = [
    ('Excel Workbook', '*.xlsx'),
    ('Excel 97-2003 Workbook', '*.xls'),
    ('CSV (Comma delimited)', '*.csv'),
    ('Text (Tab delimited), .tsv', '*.tsv'),
    ('Text (Tab delimited), .tab', '*.tab'),
    ('Text (Plain text)', '*.txt'),
    ('HTML (Web Page)', '*.htm;*.html'),
    ('JSON (JavaScript Object Notation)', '*.json'),
    ('XML (eXtensible Markup Language)', '*.xml')
]

# -------------------------------
# UI: Select sheets dialog
# -------------------------------


def select_excel_sheets_dialog(file_sheet_map):
    dialog = tk.Toplevel()
    dialog.title("Select Sheets to Convert")
    dialog.geometry("700x500")
    try:
        icon_path = os.path.join(ASSET_DIR, "app.ico")
        dialog.iconbitmap(icon_path)
    except Exception:
        pass

    tk.Label(
        dialog,
        text="Choose sheets to convert from each file:",
        font=("Segoe UI", 14, "bold"),
        fg="#314C9D"
    ).pack(pady=(16, 8))

    scroll_canvas = tk.Canvas(dialog, borderwidth=0, bg="#fafaff")
    h_scroll = tk.Scrollbar(dialog, orient="horizontal",
                            command=scroll_canvas.xview)
    scroll_canvas.configure(xscrollcommand=h_scroll.set)
    h_scroll.pack(side="bottom", fill="x")
    scroll_canvas.pack(side="top", fill="both", expand=True)

    files_frame = tk.Frame(scroll_canvas, bg="#fafaff")
    scroll_canvas.create_window((0, 0), window=files_frame, anchor="nw")
    files_frame.bind("<Configure>", lambda e: scroll_canvas.configure(
        scrollregion=scroll_canvas.bbox("all")))

    all_sheet_vars = {}
    col = 0
    for fname, sheet_names in file_sheet_map.items():
        file_frame = tk.Frame(files_frame, bg="#fafaff", bd=2, relief="groove")
        file_frame.grid(row=0, column=col, sticky="n", padx=12, pady=8)

        tk.Label(
            file_frame,
            text=os.path.basename(fname),
            font=("Segoe UI", 12, "bold"),
            fg="#253A7D",
            bg="#fafaff"
        ).pack(anchor="n", padx=8, pady=(8, 2))

        file_btn_frame = tk.Frame(file_frame, bg="#fafaff")
        file_btn_frame.pack(anchor="n", padx=8, pady=(0, 2))
        sheet_vars = {}

        def make_select_all(vars_dict):
            return lambda: [v.set(True) for v in vars_dict.values()]

        def make_deselect_all(vars_dict):
            return lambda: [v.set(False) for v in vars_dict.values()]

        ctk.CTkButton(
            file_btn_frame, text="Select All",
            command=make_select_all(sheet_vars),
            fg_color="#0984e3", text_color="white",
            corner_radius=8, width=100
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            file_btn_frame, text="Deselect All",
            command=make_deselect_all(sheet_vars),
            fg_color="#e17055", text_color="white",
            corner_radius=8, width=100
        ).pack(side="left")

        for sheet in sheet_names:
            var = tk.BooleanVar(value=True)
            tk.Checkbutton(
                file_frame, text=sheet, variable=var,
                font=("Segoe UI", 11), anchor="w", bg="#fafaff"
            ).pack(anchor="w", padx=16, pady=2)
            sheet_vars[sheet] = var

        all_sheet_vars[fname] = sheet_vars
        col += 1

    selected = {}

    def on_ok():
        for fname, vars_dict in all_sheet_vars.items():
            selected[fname] = [sheet for sheet,
                               var in vars_dict.items() if var.get()]
        dialog.result = selected
        dialog.destroy()

    def on_close():
        dialog.result = None
        dialog.destroy()

    ctk.CTkButton(
        dialog, text="OK", command=on_ok,
        corner_radius=12, fg_color="#0984e3",
        hover_color="#314C9D", text_color="white",
        font=("Segoe UI", 12, "bold"), height=32
    ).pack(pady=12)

    dialog.protocol("WM_DELETE_WINDOW", on_close)
    dialog.grab_set()
    dialog.result = None
    dialog.wait_window()
    return dialog.result

# -------------------------------
# Delimiter helpers
# -------------------------------


def guess_csv_delimiter(path, encodings=('utf-8', 'latin1')):
    sample = None
    for enc in encodings:
        try:
            with open(path, 'r', encoding=enc) as f:
                sample = f.read(4096)
            break
        except Exception:
            continue
    if not sample:
        return None
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except Exception:
        if ',' in sample and '\t' in sample:
            return ',' if sample.count(',') > sample.count('\t') else '\t'
        elif '\t' in sample:
            return '\t'
        elif ',' in sample:
            return ','
        else:
            return None


def _looks_tab_delimited(path: str, sample_bytes: int = 4096) -> bool:
    """
    Peek into the file and decide if it's tab-delimited:
    - True if we see tabs and tabs >= commas.
    """
    try:
        with open(path, "rb") as f:
            sample = f.read(sample_bytes)
        for enc in ("utf-8", "utf-8-sig", "latin1"):
            try:
                text = sample.decode(enc, errors="ignore")
                break
            except Exception:
                continue
        else:
            text = sample.decode("latin1", errors="ignore")
        return text.count("\t") > 0 and text.count("\t") >= text.count(",")
    except Exception:
        return False


# -------------------------------
# File type detection
# -------------------------------
OLE_SIGNATURE = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
ZIP_SIGNATURE = b'\x50\x4B\x03\x04'


def _is_ole_binary(path: str) -> bool:
    try:
        with open(path, 'rb') as f:
            sig = f.read(8)
            return sig.startswith(OLE_SIGNATURE)
    except Exception:
        return False


def _is_zip_xlsx(path: str) -> bool:
    try:
        with open(path, 'rb') as f:
            sig = f.read(4)
            return sig.startswith(ZIP_SIGNATURE)
    except Exception:
        return False

# -------------------------------
# Legacy .xls readers
# -------------------------------


def _read_xls_with_xlrd(path: str) -> pd.DataFrame:
    """Read legacy .xls using xlrd==1.2.0; returns empty df if not possible."""
    try:
        import xlrd  # noqa
        ver = getattr(xlrd, '__version__', '')
        if not ver or ver.startswith('2'):
            raise ImportError("xlrd>=2.0 installed; .xls support removed.")
    except ImportError as e:
        raise e  # bubble up for COM fallback

    xls = pd.ExcelFile(path, engine='xlrd')
    dfs = []
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet)
            if not df.empty:
                df['SheetName'] = sheet
                dfs.append(df)
        except Exception:
            continue
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def _read_xls_via_excel_com(path: str) -> pd.DataFrame:
    """
    Use Excel (win32com.client) to save the .xls as .xlsx, then read via openpyxl.
    Works only on Windows with Excel installed.
    """
    try:
        import pythoncom
        from win32com.client import Dispatch
    except Exception as e:
        raise ImportError(
            "Excel COM not available (requires Windows + Excel + pywin32). "
            f"Import error: {e}"
        )

    pythoncom.CoInitialize()
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    wb = None
    tmp_xlsx = None
    try:
        wb = excel.Workbooks.Open(
            os.path.abspath(path),
            UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True
        )
        fd, tmp_xlsx = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        if os.path.exists(tmp_xlsx):
            try:
                os.remove(tmp_xlsx)
            except Exception:
                base, ext = os.path.splitext(tmp_xlsx)
                tmp_xlsx = base + "_1" + ext

        wb.SaveAs(tmp_xlsx, FileFormat=51)  # xlOpenXMLWorkbook
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.DisplayAlerts = True
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    try:
        xls = pd.ExcelFile(tmp_xlsx, engine='openpyxl')
        dfs = []
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            if not df.empty:
                df['SheetName'] = sheet
                dfs.append(df)
        return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    finally:
        try:
            if tmp_xlsx and os.path.exists(tmp_xlsx):
                os.remove(tmp_xlsx)
        except Exception:
            pass

# -------------------------------
# Multi-sheet helpers for .xls
# -------------------------------


def get_xls_sheet_names(path: str) -> list[str]:
    """
    Return sheet names for a legacy .xls workbook.
    Tries xlrd first; if not available/compatible, uses Excel COM.
    """
    # xlrd path
    try:
        import xlrd
        ver = getattr(xlrd, '__version__', '')
        if ver and not ver.startswith('2'):
            book = xlrd.open_workbook(path)
            return [s.name for s in book.sheets()]
    except Exception:
        pass

    # COM fallback
    try:
        import pythoncom
        from win32com.client import Dispatch
        pythoncom.CoInitialize()
        excel = Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(
            path), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
        names = [ws.Name for ws in wb.Worksheets]
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return names
    except Exception as e:
        raise RuntimeError(f"Unable to list .xls sheets via COM: {e}")


def read_xls_selected_sheets(path: str, sheet_names: list[str]) -> list[tuple[str, pd.DataFrame]]:
    """
    Read selected sheets from a legacy .xls workbook and return [(sheet_name, df), ...].
    Prefers xlrd; falls back to Excel COM -> temporary .xlsx -> openpyxl.
    """
    results: list[tuple[str, pd.DataFrame]] = []

    # xlrd path
    try:
        import xlrd
        ver = getattr(xlrd, '__version__', '')
        if ver and not ver.startswith('2'):
            xls = pd.ExcelFile(path, engine='xlrd')
            for sheet in sheet_names:
                try:
                    df = xls.parse(sheet)
                    if not df.empty:
                        results.append((sheet, df))
                except Exception:
                    continue
            return results
    except Exception:
        pass

    # COM fallback: save once to temp .xlsx, then parse selected sheets via openpyxl
    tmp_xlsx = None
    try:
        import pythoncom
        from win32com.client import Dispatch
        pythoncom.CoInitialize()
        excel = Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(
            path), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
        fd, tmp_xlsx = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        if os.path.exists(tmp_xlsx):
            try:
                os.remove(tmp_xlsx)
            except Exception:
                base, ext = os.path.splitext(tmp_xlsx)
                tmp_xlsx = base + "_1" + ext
        wb.SaveAs(tmp_xlsx, FileFormat=51)  # xlOpenXMLWorkbook
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        xls = pd.ExcelFile(tmp_xlsx, engine='openpyxl')
        for sheet in sheet_names:
            try:
                df = xls.parse(sheet)
                if not df.empty:
                    results.append((sheet, df))
            except Exception:
                continue
        return results
    finally:
        try:
            if tmp_xlsx and os.path.exists(tmp_xlsx):
                os.remove(tmp_xlsx)
        except Exception:
            pass

# -------------------------------
# .TAB / .TSV (and tabbed .TXT) strict reader
# -------------------------------


def read_tab_strict(path: str) -> pd.DataFrame:
    """
    Robust reader for .TAB / .TSV royalty statements:
      - Force tab separator
      - Read all fields as string
      - No NA conversion
      - Strip whitespace from headers and values
      - Drop fully-empty columns
      - Normalize "+000000...0.xxx" to "0.xxx" (kept as strings)
    """
    last_err = None
    for enc in ("utf-8", "utf-8-sig", "latin1"):
        try:
            df = pd.read_csv(
                path,
                sep="\t",
                dtype=str,         # preserve leading zeros / signs
                na_filter=False,   # keep empty strings
                engine="python",
                quoting=csv.QUOTE_NONE,
                encoding=enc
            )
            break
        except Exception as e:
            last_err = e
            df = None
    if df is None:
        raise last_err if last_err else RuntimeError(
            "Unable to read .TAB file")

    # Trim headers
    df.columns = [c.strip() for c in df.columns]

    # Trim cell values
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(lambda x: x.strip()
                                  if isinstance(x, str) else x)

    # Drop fully-empty columns
    empty_cols = [c for c in df.columns if (df[c] == "").all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)

    # Normalize padded-plus numbers (kept as strings)
    def _normalize_plus_padded(s: str) -> str:
        if not isinstance(s, str):
            return s
        s2 = s[1:] if s.startswith("+") else s
        if "." in s2:
            left, right = s2.split(".", 1)
            if left.isdigit():
                left_norm = "0" if int(left) == 0 else str(int(left))
                return f"{left_norm}.{right}"
        if s2.isdigit():
            return "0" if int(s2) == 0 else str(int(s2))
        return s

    likely_numeric_cols = [
        c for c in df.columns
        if any(k in c.lower() for k in [
            "units", "amount", "rate", "royalties", "payable",
            "share", "ppd", "retail", "price", "payout", "%", "received"
        ])
    ]
    for col in likely_numeric_cols:
        df[col] = df[col].map(_normalize_plus_padded)

    return df

# -------------------------------
# Unified reader
# -------------------------------


def read_file(path: str) -> pd.DataFrame:
    """
    Robust reader:
    - OLE .xls: try xlrd==1.2.0, else Excel COM with a friendly error if not available.
    - ZIP .xlsx/.xlsm/.xlsb: read via pandas/openpyxl/pyxlsb.
    - HTML/JSON/XML: pandas readers.
    - .tab/.tsv (and tabbed .txt): strict handler.
    - Other delimited text: delimiter guess, read as strings, trim.
    """
    ext = os.path.splitext(path)[1].lower()
    is_ole = _is_ole_binary(path)
    is_zip = _is_zip_xlsx(path)

    # Legacy Excel (.xls / OLE)
    if ext == '.xls' or is_ole:
        try:
            return _read_xls_with_xlrd(path)
        except ImportError:
            try:
                return _read_xls_via_excel_com(path)
            except Exception as com_err:
                root = tk._default_root or tk.Tk()
                try:
                    root.withdraw()
                except Exception:
                    pass
                messagebox.showerror(
                    "Unable to read legacy .xls",
                    "This file is a true Excel 97–2003 binary workbook (.xls).\n\n"
                    "To read it, please install:\n pip install \"xlrd==1.2.0\"\n"
                    "—or— run on Windows with Excel installed and install:\n pip install pywin32\n\n"
                    f"Technical note: COM fallback error was:\n{com_err}"
                )
                return pd.DataFrame()

    # Modern ZIP-based Excel
    if ext in ('.xlsx', '.xlsm', '.xlsb') or is_zip:
        tried_engines = ['openpyxl', 'pyxlsb', None]
        for engine in tried_engines:
            try:
                xls = pd.ExcelFile(
                    path, engine=engine) if engine else pd.ExcelFile(path)
                dfs = []
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    if not df.empty:
                        df['SheetName'] = sheet
                        dfs.append(df)
                if dfs:
                    return pd.concat(dfs, ignore_index=True)
            except Exception:
                continue
        messagebox.showerror("Excel Read Error",
                             "Could not read the modern Excel file.")
        return pd.DataFrame()

    # HTML
    if ext in ('.htm', '.html'):
        try:
            tables = pd.read_html(path, encoding='utf-8')
            return tables[0] if tables else pd.DataFrame()
        except Exception as e:
            messagebox.showerror(
                'HTML Read Error', f'Could not read HTML tables: {e}')
            return pd.DataFrame()

    # JSON
    if ext == '.json':
        try:
            return pd.read_json(path, encoding='utf-8')
        except Exception as e:
            messagebox.showwarning("JSON Parsing Failed", str(e))
            return pd.DataFrame()

    # XML
    if ext == '.xml':
        try:
            return pd.read_xml(path, encoding='utf-8')
        except Exception as e:
            messagebox.showwarning("XML Parsing Failed", str(e))
            return pd.DataFrame()

    # TAB / TSV
    if ext in ('.tab', '.tsv'):
        try:
            return read_tab_strict(path)
        except Exception as e:
            messagebox.showwarning("TAB Parsing Failed",
                                   f"Could not parse as .TAB/.TSV:\n{e}")
            return pd.DataFrame()

    # TXT that looks tabbed
    if ext == '.txt' and _looks_tab_delimited(path):
        try:
            return read_tab_strict(path)
        except Exception as e:
            messagebox.showwarning("TAB Parsing Failed",
                                   f"Could not parse tabbed .TXT:\n{e}")
            return pd.DataFrame()

    # Other delimited text
    try:
        sep = guess_csv_delimiter(path) or ','
        df = pd.read_csv(path, sep=sep, engine='python',
                         dtype=str, na_filter=False)
        df.columns = [c.strip() for c in df.columns]
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].map(lambda x: x.strip()
                                      if isinstance(x, str) else x)
        return df
    except Exception as e:
        messagebox.showwarning(
            "Parsing Failed", f"Could not parse as a delimited text file.\n{e}")
        return pd.DataFrame()

# -------------------------------
# Writer
# -------------------------------


def write_file(df, path, out_format):
    if df is None or df.empty:
        return
    if out_format == 'xlsx':
        df.to_excel(path, index=False, engine='openpyxl')
    elif out_format == 'xls':
        try:
            df.to_excel(path, index=False, engine='xlwt')
        except Exception as e:
            messagebox.showerror(
                "Write Error (.xls)",
                "Writing .xls requires the 'xlwt' package.\n\n"
                f"Error: {e}\n\n"
                "Install with:\n pip install xlwt\n"
                "Or choose 'xlsx' as output."
            )
    elif out_format == 'csv':
        df.to_csv(path, index=False)
    elif out_format in ('tsv', 'tab', 'txt'):
        df.to_csv(path, sep='\t', index=False)
    elif out_format == 'json':
        df.to_json(path, orient='records', lines=False, force_ascii=False)
    elif out_format == 'xml':
        try:
            df.to_xml(path, index=False)
        except Exception as e:
            raise ValueError(f'Error writing XML: {e}')
    else:
        raise ValueError('Unsupported output format')

# -------------------------------
# GUI App
# -------------------------------


class FileConverterApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # Cache for Excel sheet names to avoid duplicate reads
        self._excel_sheet_cache = {}

        # Status bar / progress
        self.statusbar_frame = tk.Frame(self, bg="white")
        self.statusbar_frame.pack(side=tk.BOTTOM, fill="x")
        self.progress_var = tk.DoubleVar(value=0)
        # Set icon as early as possible for taskbar and window

        self.title("File Converter")
        try:
            icon_path = os.path.join(ASSET_DIR, "app.ico")
            self.iconbitmap(icon_path)
            self.wm_iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon set failed: {e}")

        self.progress = ttk.Progressbar(
            self.statusbar_frame, variable=self.progress_var, maximum=100)
        self.progress.pack(side=tk.LEFT, fill="x",
                           expand=True, padx=(10, 4), pady=2)
        self.status_label = tk.Label(
            self.statusbar_frame, text="Ready", bg="white", font=("Segoe UI", 10))
        self.status_label.pack(side=tk.LEFT, padx=(4, 10))

        self.geometry("900x800")
        self.minsize(700, 600)
        self.configure(bg="white")
        self.resizable(True, True)

        # Logo / Title
        try:
            from PIL import Image, ImageTk
            logo_img = None
            for fname in ("logo32.png", "logo64.png", "logo100.png"):
                logo_path = os.path.join(ASSET_DIR, fname)
                try:
                    logo = Image.open(logo_path)
                    logo_img = ImageTk.PhotoImage(logo)
                    break
                except Exception as e:
                    print(f"Logo load failed for {fname}: {e}")
            logo_frame = tk.Frame(self, bg="white")
            logo_frame.pack(fill=tk.X, anchor="nw")
            if logo_img:
                logo_label = tk.Label(logo_frame, image=logo_img, bg="white")
                logo_label.image = logo_img
                logo_label.pack(side=tk.LEFT)
            title_label = tk.Label(logo_frame, text="File Converter",
                                   font=("Segoe UI", 18, "bold"),
                                   bg="white", fg="#2793C9")
            title_label.pack(side=tk.LEFT, padx=10)
        except Exception as e:
            print(f"Logo block failed: {e}")
            label = tk.Label(self, text="File Converter",
                             font=("Segoe UI", 18, "bold"),
                             bg="white", fg="#2793C9")
            label.pack(anchor="nw", padx=10, pady=10)

        # Drag & drop area
        self.file_path = tk.StringVar()
        drop_canvas = tk.Canvas(
            self, height=140, bg="white", highlightthickness=0)
        drop_canvas.pack(fill="x", expand=True, padx=60, pady=(2, 4))

        def draw_rounded_rect(canvas, x1, y1, x2, y2, radius=30, **kwargs):
            points = [
                x1+radius, y1,
                x2-radius, y1,
                x2, y1,
                x2, y1+radius,
                x2, y2-radius,
                x2, y2,
                x2-radius, y2,
                x1+radius, y2,
                x1, y2,
                x1, y2-radius,
                x1, y1+radius,
                x1, y1
            ]
            return canvas.create_polygon(points, smooth=True, **kwargs)

        def resize_drop_area(event):
            drop_canvas.delete("rounded_rect")
            width = event.width
            height = event.height
            draw_rounded_rect(drop_canvas, 5, 5, width-5, height-5, radius=30,
                              fill="#f3f3f3", outline="#2793C9", width=4, tags="rounded_rect")
            drop_canvas.coords(drop_window, width//2, height//2)
            drop_canvas.itemconfig(drop_window, width=max(
                320, width-20), height=max(80, height-20))

        drop_frame = tk.Frame(drop_canvas, bg="#f3f3f3")
        drop_window = drop_canvas.create_window(
            0, 0, window=drop_frame, anchor="center", width=400, height=120)
        drop_canvas.bind("<Configure>", resize_drop_area)

        drop_frame.grid_rowconfigure(0, weight=1)
        drop_frame.grid_columnconfigure(0, weight=1)
        drop_frame.grid_columnconfigure(1, weight=1)

        try:
            from PIL import Image, ImageTk
            drag_icon_path = os.path.join(ASSET_DIR, "drag100.png")
            drag_icon = Image.open(drag_icon_path).resize((80, 80))
            drag_icon_img = ImageTk.PhotoImage(drag_icon)
            drag_label = tk.Label(
                drop_frame, image=drag_icon_img, bg="#f3f3f3", bd=0)
            drag_label.image = drag_icon_img
            drag_label.grid(row=0, column=0, padx=(
                18, 18), pady=10, sticky="e")
        except Exception:
            drag_label = tk.Label(drop_frame, text="⇅", font=(
                "Segoe UI", 38), bg="#f3f3f3", fg="#2793C9")
            drag_label.grid(row=0, column=0, padx=(
                18, 18), pady=10, sticky="e")

        drop_label = tk.Label(
            drop_frame,
            text="Drag & drop files or folders here\nor click to add files or folders",
            font=("Segoe UI", 16), bg="#f3f3f3", fg="#314C9D"
        )
        drop_label.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="w")

        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        drop_canvas.drop_target_register(DND_FILES)
        drop_canvas.dnd_bind('<<Drop>>', self.on_drop)

        drop_frame.bind("<Button-1>", lambda e: self.browse_file())
        drag_label.bind("<Button-1>", self.browse_file)
        drop_label.bind("<Button-1>", lambda e: self.browse_file())

        # Status panel
        self.status_frame = tk.Frame(self, bg="white")
        self.status_frame.pack(fill="both", padx=40, pady=(0, 1), expand=True)

        self.clear_btn = ctk.CTkButton(
            self.status_frame, text="Clear Files", command=self.clear_files,
            corner_radius=12, fg_color="#e17055", hover_color="#d35400",
            text_color="white", font=("Segoe UI", 12, "bold"), height=32, width=120
        )
        self.clear_btn.pack(pady=(6, 2), anchor="e", padx=2)

        self.status_label_border = tk.Canvas(
            self.status_frame, bg="white", highlightthickness=0, height=38)
        self.status_label_border.pack(fill="x", pady=(0, 0))

        def draw_label_border(canvas, x1, y1, x2, y2, radius=12, **kwargs):
            points = [
                x1+radius, y1,
                x2-radius, y1,
                x2, y1,
                x2, y1+radius,
                x2, y2-radius,
                x2, y2,
                x2-radius, y2,
                x1+radius, y2,
                x1, y2,
                x1, y2-radius,
                x1, y1+radius,
                x1, y1
            ]
            return canvas.create_polygon(points, smooth=True, fill="#f3f3f3", outline="#2793C9", width=2, tags="label_border")

        def resize_label_border(event):
            self.status_label_border.delete("label_border")
            width = event.width
            height = event.height
            draw_label_border(self.status_label_border, 2, 2,
                              width-2, height-2, radius=12)

        self.status_label_border.bind("<Configure>", resize_label_border)
        label_frame = tk.Frame(self.status_label_border, bg="#f3f3f3")
        self.status_label = tk.Label(label_frame, text="Files to Convert",
                                     bg="#f3f3f3", font=("Segoe UI", 12, "bold"),
                                     fg="#314C9D", bd=0)
        self.status_label.pack(side=tk.LEFT, padx=(0, 8))
        self.status_label_window = self.status_label_border.create_window(
            10, 19, window=label_frame, anchor="w")

        # Files list
        self.status_box_outer = tk.Frame(self.status_frame, bg="white")
        self.status_box_outer.pack(fill="both", expand=True, pady=(0, 0))

        self.file_list_border = tk.Frame(self.status_box_outer, bg="#f3f3f3",
                                         highlightbackground="#2793C9", highlightthickness=2, bd=0)
        self.file_list_border.pack(fill="both", expand=True, padx=0, pady=0)

        self.status_canvas = tk.Canvas(self.file_list_border, bg="#f3f3f3",
                                       highlightthickness=0, height=210)
        self.status_canvas.pack(fill="both", expand=True, side="left")

        self.file_scrollbar = tk.Scrollbar(self.file_list_border, orient="vertical",
                                           command=self.status_canvas.yview)
        self.file_scrollbar.pack(fill="y", side="right")
        self.status_canvas.configure(yscrollcommand=self.file_scrollbar.set)

        self.file_check_container = tk.Frame(self.status_canvas, bg="#f3f3f3")
        self.file_check_container_id = self.status_canvas.create_window(0, 0,
                                                                        window=self.file_check_container,
                                                                        anchor="nw", width=1)

        def update_canvas_size(event):
            self.status_canvas.configure(
                scrollregion=self.status_canvas.bbox("all"))
            self.status_canvas.itemconfig(
                self.file_check_container_id, width=self.status_canvas.winfo_width())

        self.file_check_container.bind("<Configure>", update_canvas_size)
        self.status_canvas.bind("<Configure>", update_canvas_size)

        self.file_vars = []
        self.select_all_var = tk.BooleanVar(value=True)
        self.select_all_cb = tk.Checkbutton(
            self.file_check_container, text="Select All",
            variable=self.select_all_var, bg="#f3f3f3",
            font=("Segoe UI", 11, "bold"), command=self.toggle_select_all
        )
        self.select_all_cb.pack(anchor="w", padx=4, pady=(0, 2))

        # Output format
        self.format_frame = tk.Frame(self, bg="white")
        self.format_frame.pack(pady=2, fill="x", padx=60)
        self.format_inner = tk.Frame(self.format_frame, bg="white")
        self.format_inner.pack(anchor="center")

        tk.Label(self.format_inner, text="Output Format:",
                 bg="white", font=("Segoe UI", 16, "bold"),
                 fg="#314C9D").pack(side=tk.LEFT, padx=(0, 12))

        # Default to 'xlsx' per your preference
        self.output_format = tk.StringVar(value='xlsx')

        style = ttk.Style()
        style.configure("Custom.TCombobox", font=("Segoe UI", 16), padding=8)

        self.format_menu = ttk.Combobox(
            self.format_inner, textvariable=self.output_format,
            values=['xlsx', 'xls', 'csv', 'tsv', 'tab', 'txt', 'json', 'xml'],
            width=10, state='readonly', style="Custom.TCombobox"
        )
        self.format_menu.pack(side=tk.LEFT)

        # Output folder
        self.output_folder = tk.StringVar()
        folder_frame = tk.Frame(self, bg="white")
        folder_frame.pack(pady=(0, 2), fill="x", padx=40)
        tk.Label(folder_frame, text="Output Folder:",
                 bg="white", font=("Segoe UI", 12),
                 fg="#314C9D").pack(side=tk.LEFT, padx=(0, 8))
        folder_entry = tk.Entry(folder_frame, textvariable=self.output_folder,
                                font=("Segoe UI", 12), width=40, state="readonly")
        folder_entry.pack(side=tk.LEFT, padx=(0, 8), fill="x", expand=True)
        folder_btn = ctk.CTkButton(
            folder_frame, text="Browse...", command=self.browse_output_folder,
            corner_radius=12, fg_color="#0984e3", hover_color="#314C9D",
            text_color="white", font=("Segoe UI", 12, "bold"), height=32, width=120
        )
        folder_btn.pack(side=tk.LEFT, padx=(0, 0), pady=2)

        # Convert button
        self.btn_frame = tk.Frame(self, bg="white")
        self.btn_frame.pack(pady=2)
        self.convert_btn = ctk.CTkButton(
            self.btn_frame, text="Convert", command=self.do_convert,
            corner_radius=16, fg_color="#0984e3", hover_color="#314C9D",
            text_color="white", font=("Segoe UI", 16, "bold"), height=48, width=180
        )
        self.convert_btn.pack(pady=6)

    # ---- UI helpers ----
    def clear_files(self):
        self.file_vars.clear()
        for widget in list(self.file_check_container.winfo_children()):
            if isinstance(widget, tk.Checkbutton) and widget is not self.select_all_cb:
                widget.destroy()
        self.select_all_var.set(False)
        self.file_check_container.update_idletasks()
        self.status_canvas.configure(
            scrollregion=self.status_canvas.bbox("all"))

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)

    def toggle_select_all(self):
        value = self.select_all_var.get()
        for var, _ in self.file_vars:
            var.set(value)

    def browse_file(self):
        all_patterns = [pat for _, pat in SUPPORTED_FORMATS]
        all_types = [("All Supported Files", " ".join(
            all_patterns))] + SUPPORTED_FORMATS
        paths = filedialog.askopenfilenames(
            title="Select files", filetypes=all_types)
        if not paths:
            folder = filedialog.askdirectory(title="Select folder")
            if folder:
                self.file_path.set(folder)
                self.update_status_listbox(folder)
        else:
            for p in paths:
                self.file_path.set(p)
                self.update_status_listbox(p)

    def on_drop(self, event):
        dropped = event.data
        pattern = r'\{([^}]*)\}|([^\s]+)'
        matches = re.findall(pattern, dropped)
        paths = [m[0] if m[0] else m[1] for m in matches]
        if len(paths) == 1:
            self.file_path.set(paths[0])
        else:
            self.file_path.set(';'.join(paths))
        for p in paths:
            self.update_status_listbox(p)

    def update_status_listbox(self, path):
        def add_file(f):
            for var, p in self.file_vars:
                if p == f:
                    return
            self.file_vars.append((tk.BooleanVar(value=True), f))
        if not path:
            return
        excel_files = {}
        files_to_process = []
        if os.path.isfile(path):
            add_file(path)
            files_to_process.append(path)
        elif os.path.isdir(path):
            for root, _, files in os.walk(path):
                for file in files:
                    fpath = os.path.join(root, file)
                    add_file(fpath)
                    files_to_process.append(fpath)

        # Show progress bar for loading sheet names
        total = len(files_to_process)
        progress = 0
        self.status_label.config(text=f"Loading files: 0/{total}")
        self.progress_var.set(0)
        self.update_idletasks()

        for idx, fpath in enumerate(files_to_process, 1):
            ext = os.path.splitext(fpath)[1].lower()
            if ext in [".xlsx", ".xlsm", ".xlsb", ".xls"]:
                if fpath in self._excel_sheet_cache:
                    excel_files[fpath] = self._excel_sheet_cache[fpath]
                else:
                    try:
                        if ext == ".xls":
                            sheet_names = get_xls_sheet_names(fpath)
                        else:
                            xls = pd.ExcelFile(fpath)
                            sheet_names = xls.sheet_names
                        self._excel_sheet_cache[fpath] = sheet_names
                        excel_files[fpath] = sheet_names
                    except Exception:
                        pass
            progress = (idx / total) * 100
            self.status_label.config(text=f"Loading files: {idx}/{total}")
            self.progress_var.set(progress)
            self.update_idletasks()

        self.status_label.config(text="Ready")
        self.progress_var.set(0)

        selected_sheets = select_excel_sheets_dialog(
            excel_files) if len(excel_files) > 1 else None
        for widget in list(self.file_check_container.winfo_children()):
            if isinstance(widget, tk.Checkbutton) and widget is not self.select_all_cb:
                widget.destroy()
        for var, p in self.file_vars:
            name = os.path.basename(p.rstrip("/\\")) or p.rstrip("/\\")
            cb = tk.Checkbutton(self.file_check_container, text=name,
                                variable=var, bg="#f3f3f3", font=("Segoe UI", 11))
            cb.pack(anchor="w", padx=4)
        self.file_check_container.update_idletasks()
        height = self.file_check_container.winfo_height()
        width = self.status_canvas.winfo_width() - 40
        if height < 50:
            height = 50
        if width < 100:
            width = 100
        self.status_canvas.coords(self.file_check_container_id, 20, 0)
        self.status_canvas.itemconfig(
            self.file_check_container_id, width=width)
        self.status_canvas.configure(
            scrollregion=self.status_canvas.bbox("all"))

    def do_convert(self):
        import time
        selected = [p for var, p in self.file_vars if var.get()]
        if not selected:
            messagebox.showerror(
                'Error', 'Please select at least one file or folder to convert!')
            return
        ext = self.output_format.get()
        output_folder = self.output_folder.get()
        if not output_folder:
            messagebox.showerror('Error', 'Please select an output folder!')
            return
        # Gather Excel files and their sheets
        excel_files = {}
        for p in selected:
            file_ext = os.path.splitext(p)[1].lower()
            if file_ext in [".xlsx", ".xlsm", ".xlsb"]:
                try:
                    xls = pd.ExcelFile(p)
                    excel_files[p] = xls.sheet_names
                except Exception:
                    pass
            elif file_ext == ".xls":
                try:
                    excel_files[p] = get_xls_sheet_names(p)
                except Exception:
                    pass
        selected_sheets = None
        if excel_files:
            selected_sheets = select_excel_sheets_dialog(excel_files)
            if selected_sheets is None:
                self.status_label.config(text="Ready")
                return
        total = len(selected)
        start_time = time.time()
        files_written = 0
        for idx, in_path in enumerate(selected, 1):
            filename = os.path.basename(in_path)
            self.status_label.config(
                text=f"Converting {idx} of {total}: {filename} ...")
            self.statusbar_frame.update_idletasks()
            if not in_path or not os.path.exists(in_path):
                messagebox.showerror(
                    'Error', f'File or folder not found: {in_path}')
                continue
            base = os.path.splitext(filename)[0]
            out_path = os.path.join(output_folder, f"{base}.{ext}")
            try:
                file_ext = os.path.splitext(in_path)[1].lower()
                if file_ext in [".xlsx", ".xlsm", ".xlsb"] and selected_sheets and in_path in selected_sheets:
                    xls = pd.ExcelFile(in_path)
                    for sheet in selected_sheets[in_path]:
                        df = xls.parse(sheet)
                        if df is not None and not df.empty:
                            out_path_sheet = os.path.join(
                                output_folder, f"{base}_{sheet}.{ext}")
                            write_file(df, out_path_sheet, ext)
                            if os.path.exists(out_path_sheet) and os.path.getsize(out_path_sheet) > 0:
                                files_written += 1
                elif file_ext == ".xls" and selected_sheets and in_path in selected_sheets:
                    # Use robust .xls sheet reader
                    for sheet, df in read_xls_selected_sheets(in_path, selected_sheets[in_path]):
                        if df is not None and not df.empty:
                            out_path_sheet = os.path.join(
                                output_folder, f"{base}_{sheet}.{ext}")
                            write_file(df, out_path_sheet, ext)
                            if os.path.exists(out_path_sheet) and os.path.getsize(out_path_sheet) > 0:
                                files_written += 1
                else:
                    df = read_file(in_path)
                    if df is not None and not df.empty:
                        write_file(df, out_path, ext)
                        if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                            files_written += 1
            except Exception as e:
                messagebox.showerror('Conversion failed', f'{filename}: {e}')
            percent = (idx / total) * 100
            elapsed = time.time() - start_time
            avg_time = elapsed / idx if idx else 0
            remaining = avg_time * (total - idx)
            self.progress_var.set(percent)
            if idx < total:
                self.status_label.config(
                    text=f"Converting {idx} of {total}: {filename} ...  {percent:.0f}% - Est. {int(remaining)}s left")
        self.status_label.config(text="Done.")
        if files_written > 0:
            messagebox.showinfo(
                'Success', f'{files_written} file(s) converted to {output_folder}')
        else:
            messagebox.showwarning(
                'No Files Converted', 'No files could be converted. Please check the file format or see previous error messages.')
        self.progress_var.set(0)
        self.status_label.config(text="Ready")


def run():
    app = FileConverterApp()
    app.mainloop()


if __name__ == "__main__":
    run()
