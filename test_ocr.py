#!/usr/bin/env python3
"""Simple OCR test using Windows.Media.OCR via PowerShell."""
import subprocess, os, sys

img = sys.argv[1] if len(sys.argv) > 1 else r'f:\0.AI设计库\ai视频识别\gh-pages-deploy\ocr_temp\pdf_page_1.png'

ps_code = '''
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$f = [Windows.Storage.StorageFile]::GetFileFromPathAsync("IMG_PATH").GetAwaiter().GetResult()
$s = [Windows.Storage.Streams.RandomAccessStreamReference]::CreateFromFile($f)
$d = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($s.OpenReadAsync().GetAwaiter().GetResult()).GetAwaiter().GetResult()
$sw = $d.GetSoftwareBitmapAsync().GetAwaiter().GetResult()
$e = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$r = $e.RecognizeAsync($sw).GetAwaiter().GetResult()
Write-Output $r.Text
'''.replace('IMG_PATH', img.replace('\\', '\\\\'))

ps_file = os.path.join(os.environ['TEMP'], 'ocr_simple.ps1')
with open(ps_file, 'w', encoding='utf-8') as f:
    f.write(ps_code)

result = subprocess.run(
    ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
    capture_output=True, timeout=60)

outpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr_result.txt')
with open(outpath, 'w', encoding='utf-8') as f:
    f.write('=== STDOUT ===\n')
    f.write(result.stdout.decode('utf-8', errors='replace'))
    f.write('\n=== STDERR ===\n')
    f.write(result.stderr.decode('utf-8', errors='replace'))
    f.write(f'\n=== RC: {result.returncode} ===\n')

print(f'Output written to: {outpath}')
print(f'Length: stdout={len(result.stdout)}, stderr={len(result.stderr)}')
