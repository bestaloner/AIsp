$imagePath = $args[0]
if (-not $imagePath) {
    Write-Error "Usage: powershell -File ocr_windows.ps1 <image_path>"
    exit 1
}

# Load WinRT types via C# code
$csharp = @'
using System;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

public static class OcrHelper
{
    public static async Task<string> RecognizeAsync(string path)
    {
        try
        {
            var file = await StorageFile.GetFileFromPathAsync(path);
            using (var stream = await file.OpenReadAsync())
            {
                var decoder = await BitmapDecoder.CreateAsync(stream);
                var bitmap = await decoder.GetSoftwareBitmapAsync();
                var engine = OcrEngine.TryCreateFromUserProfileLanguages();
                if (engine == null) return "[ERROR: No OCR engine for current language]";
                var result = await engine.RecognizeAsync(bitmap);
                return result.Text;
            }
        }
        catch (Exception ex)
        {
            return "[ERROR: " + ex.Message + "]";
        }
    }
}
'@

# Add the C# type with WinRT references
$winmdPath = "$env:SystemRoot\System32\WinMetadata"
Add-Type -TypeDefinition $csharp `
    -ReferencedAssemblies "$winmdPath\Windows.Foundation.winmd", `
                         "$winmdPath\Windows.Graphics.Imaging.winmd", `
                         "$winmdPath\Windows.Media.Ocr.winmd", `
                         "$winmdPath\Windows.Storage.winmd", `
                         "$winmdPath\Windows.Storage.Streams.winmd"

# Run OCR
$task = [OcrHelper]::RecognizeAsync($imagePath)
$task.Wait()
Write-Output $task.Result
