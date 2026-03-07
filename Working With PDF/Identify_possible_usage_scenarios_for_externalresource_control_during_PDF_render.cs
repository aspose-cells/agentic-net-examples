using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExternalResourceControlDemo
{
    static void Main()
    {
        // Create a temporary image file to use in the demo (1×1 pixel PNG)
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK8cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a workbook and add a picture
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Pictures.Add(1, 1, imagePath);

        // Scenario 1: Default loading – let Aspose.Cells load the external image normally
        wb.Settings.ResourceProvider = new DefaultStreamProvider();
        SavePdf(wb, "DefaultLoading.pdf");

        // Scenario 2: Skip loading – ignore external resources to speed up rendering
        wb.Settings.ResourceProvider = new SkipStreamProvider();
        SavePdf(wb, "SkipLoading.pdf");

        // Scenario 3: User‑provided stream – supply a custom image stream (e.g., from memory)
        wb.Settings.ResourceProvider = new MemoryStreamProvider();
        SavePdf(wb, "UserProvided.pdf");

        // Clean up temporary image file
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    // Helper method to save a workbook as PDF with basic options
    static void SavePdf(Workbook wb, string fileName)
    {
        PdfSaveOptions options = new PdfSaveOptions
        {
            // Continue rendering even if a resource cannot be loaded
            IgnoreError = true
        };
        wb.Save(fileName, options);
        Console.WriteLine($"Saved {fileName}");
    }
}

// --------------------------------------------------------------------
// Default provider – uses the standard file‑system loading behavior
// --------------------------------------------------------------------
class DefaultStreamProvider : IStreamProvider
{
    public void InitStream(StreamProviderOptions options)
    {
        // Load the resource from its default path on disk
        options.Stream = new FileStream(options.DefaultPath, FileMode.Open, FileAccess.Read);
    }

    public void CloseStream(StreamProviderOptions options)
    {
        options.Stream?.Close();
    }
}

// --------------------------------------------------------------------
// Skip provider – tells the renderer to ignore the external resource
// --------------------------------------------------------------------
class SkipStreamProvider : IStreamProvider
{
    public void InitStream(StreamProviderOptions options)
    {
        // Provide an empty stream to indicate the resource should be skipped
        options.Stream = Stream.Null;
    }

    public void CloseStream(StreamProviderOptions options)
    {
        // No cleanup required for Stream.Null
    }
}

// --------------------------------------------------------------------
// User‑provided provider – supplies a custom stream (e.g., an in‑memory image)
// --------------------------------------------------------------------
class MemoryStreamProvider : IStreamProvider
{
    public void InitStream(StreamProviderOptions options)
    {
        if (options.ResourceLoadingType == ResourceLoadingType.UserProvided)
        {
            // Example: a 1×1 pixel PNG created in memory
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK8cAAAAASUVORK5CYII=");
            options.Stream = new MemoryStream(pngBytes);
        }
        else if (options.ResourceLoadingType == ResourceLoadingType.Skip)
        {
            options.Stream = Stream.Null;
        }
        else
        {
            // Fallback to default file loading for other cases
            options.Stream = new FileStream(options.DefaultPath, FileMode.Open, FileAccess.Read);
        }
    }

    public void CloseStream(StreamProviderOptions options)
    {
        options.Stream?.Close();
    }
}