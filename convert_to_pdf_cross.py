import sys
import os
import platform
from pathlib import Path
import subprocess
import traceback

# Windows için gerekli modüller
if platform.system() == "Windows":
    try:
        from docx2pdf import convert as docx2pdf_convert
        import win32com.client
        import comtypes.client
        import comtypes
    except ImportError:
        print("Windows ortamında gerekli kütüphaneler (docx2pdf, win32com, comtypes) yüklü değil.")
        print("pip install pywin32 comtypes docx2pdf komutunu çalıştırın.")
        sys.exit(1)

# Uzantılar
EXT_WORD = {'.docx', '.doc'}
EXT_EXCEL = {'.xlsx', '.xls', '.xlsm'}
EXT_PPT = {'.pptx', '.ppt', '.pptm'}

# -------------------
# Windows Dönüştürmeleri
# -------------------
def convert_word_windows(src: Path, dst: Path):
    try:
        docx2pdf_convert(str(src), str(dst.resolve()))
        print(f"[WORD] {src.name} -> {dst.name}")
    except Exception as e:
        print(f"[ERROR] Word dönüştürülürken hata: {src.name} -> {e}")
        traceback.print_exc()

def convert_excel_windows(src: Path, dst: Path):
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(src), ReadOnly=1)
        wb.ExportAsFixedFormat(0, str(dst.resolve()))
        wb.Close(False)
        excel.Quit()
        print(f"[EXCEL] {src.name} -> {dst.name}")
    except Exception as e:
        print(f"[ERROR] Excel dönüştürülürken hata: {src.name} -> {e}")
        traceback.print_exc()

def convert_ppt_windows(src: Path, dst: Path):
    try:
        comtypes.CoInitialize()
        ppt_app = comtypes.client.CreateObject("Powerpoint.Application")
        ppt_app.Visible = 1
        presentation = ppt_app.Presentations.Open(str(src), WithWindow=False)
        presentation.SaveAs(str(dst.resolve()), 32)
        presentation.Close()
        ppt_app.Quit()
        comtypes.CoUninitialize()
        print(f"[PPT] {src.name} -> {dst.name}")
    except Exception as e:
        print(f"[ERROR] PowerPoint dönüştürülürken hata: {src.name} -> {e}")
        traceback.print_exc()

# -------------------
# Linux / macOS Dönüştürmeleri
# -------------------
def convert_office_unix(src: Path, dst: Path):
    try:
        libreoffice_path = "/opt/libreoffice25.2/program/soffice"
        # LibreOffice komut satırı ile PDF'e çevir
        subprocess.run([
            libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(dst.parent.resolve()),
            str(src.resolve())
        ], check=True)
        print(f"[PDF] {src.name} -> {dst.name}")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] LibreOffice dönüştürülürken hata: {src.name} -> {e}")
        traceback.print_exc()
    except FileNotFoundError:
        print("LibreOffice bulunamadı. Lütfen libreoffice kurulu ve PATH'te olduğundan emin olun.")

# -------------------
# Dosya dönüşüm kontrolü
# -------------------
def convert_file(path: Path):
    dst = path.with_suffix('.pdf')
    ext = path.suffix.lower()
    
    system = platform.system()
    
    if system == "Windows":
        if ext in EXT_WORD:
            convert_word_windows(path, dst)
        elif ext in EXT_EXCEL:
            convert_excel_windows(path, dst)
        elif ext in EXT_PPT:
            convert_ppt_windows(path, dst)
        else:
            print(f"[SKIP] Desteklenmeyen dosya: {path.name}")
    else:  # Linux veya macOS
        if ext in EXT_WORD.union(EXT_EXCEL).union(EXT_PPT):
            convert_office_unix(path, dst)
        else:
            print(f"[SKIP] Desteklenmeyen dosya: {path.name}")

# -------------------
# Klasör / Dosya işlemleri
# -------------------
def convert_path(path: Path):
    if path.is_file():
        convert_file(path)
    elif path.is_dir():
        files = list(path.rglob("*"))
        office_files = [f for f in files if f.suffix.lower() in EXT_WORD.union(EXT_EXCEL).union(EXT_PPT)]
        print(f"Toplam {len(office_files)} dosya bulundu. Dönüştürülüyor...")
        for f in office_files:
            convert_file(f)
        print("Tüm dosyalar dönüştürüldü.")
    else:
        print("Geçerli bir dosya veya klasör yolu girin.")

# -------------------
# Main
# -------------------
def main():
    if len(sys.argv) < 2:
        print("Kullanım: python convert_to_pdf.py <dosya_veya_klasor_yolu>")
        return
    p = Path(sys.argv[1]).expanduser().resolve()
    if not p.exists():
        print(f"Hata: Yol bulunamadı: {p}")
        return
    convert_path(p)

if __name__ == "__main__":
    main()
