#!/usr/bin/env python3
"""OCR utility: reads image-based PDFs and images using Windows built-in OCR.
Usage: python run_ocr.py <pdf_or_image_path>
Output: extracted text to stdout"""

import sys, os, subprocess, tempfile
import fitz  # PyMuPDF

def ocr_image_windows(image_path):
    """Use Windows.Media.OCR via PowerShell to extract text from an image."""
    ps_script = '''
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$imageFile = [Windows.Storage.StorageFile]::GetFileFromPathAsync($args[0]).GetAwaiter().GetResult()
$stream = [Windows.Storage.Streams.RandomAccessStreamReference]::CreateFromFile($imageFile)
$bitmap = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream.OpenReadAsync().GetAwaiter().GetResult()).GetAwaiter().GetResult()
$softwareBitmap = [Windows.Graphics.Imaging.BitmapDecoder]::GetSoftwareBitmapAsync($bitmap).GetAwaiter().GetResult()
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$result = $engine.RecognizeAsync($softwareBitmap).GetAwaiter().GetResult()
Write-Output $result.Text
'''
    try:
        with tempfile.NamedTemporaryFile(suffix='.ps1', mode='w', encoding='utf-8', delete=False) as f:
            f.write(ps_script)
            ps_path = f.name

        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_path, image_path],
            capture_output=True, text=True, timeout=60)
        os.unlink(ps_path)
        if result.returncode != 0:
            return f"[OCR Error] {result.stderr[:200]}"
        return result.stdout.strip()
    except Exception as e:
        return f"[OCR Error] {e}"

def ocr_image_tesseract(image_path):
    """Fallback: use pytesseract if available."""
    try:
        import pytesseract
        from PIL import Image
        img = Image.open(image_path)
        return pytesseract.image_to_string(img, lang='chi_sim+eng')
    except ImportError:
        return None

def process_file(filepath):
    """Process a PDF or image file and return OCR text."""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.pdf':
        doc = fitz.open(filepath)
        results = []
        outdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr_temp')
        os.makedirs(outdir, exist_ok=True)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=200)
            img_path = os.path.join(outdir, f'pdf_page_{i+1}.png')
            pix.save(img_path)
            print(f"Processing page {i+1}/{doc.page_count}...", file=sys.stderr)
            text = ocr_image_windows(img_path)
            if not text or text.startswith('[OCR Error]'):
                # Try tesseract fallback
                t = ocr_image_tesseract(img_path)
                if t:
                    text = t
            results.append(f"=== Page {i+1} ===\n{text}")
            os.unlink(img_path)
        doc.close()
        return '\n\n'.join(results)

    elif ext in ('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif'):
        text = ocr_image_windows(filepath)
        if not text or text.startswith('[OCR Error]'):
            t = ocr_image_tesseract(filepath)
            if t:
                text = t
        return text

    else:
        return f"Unsupported file type: {ext}"

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python run_ocr.py <pdf_or_image_path>")
        sys.exit(1)

    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"File not found: {path}")
        sys.exit(1)

    text = process_file(path)
    if text:
        print(text)
    else:
        print("[No text extracted - try installing tesseract: winget install UB-Mannheim.TesseractOCR]")
