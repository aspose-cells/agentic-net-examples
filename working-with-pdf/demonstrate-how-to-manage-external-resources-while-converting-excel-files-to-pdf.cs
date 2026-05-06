using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Drawing;

public class ExternalResourcesPdfConversionDemo
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Paths for source Excel, destination PDF, temporary image, and cache folder
        string sourceFile = "sample.xlsx";
        string destFile = "output.pdf";
        string tempImageFile = "logo.png";
        string cachedFolder = Path.Combine(Path.GetTempPath(), "AsposeCache");

        try
        {
            // Ensure the cache folder exists (used by PdfSaveOptions for temporary files)
            Directory.CreateDirectory(cachedFolder);

            // Create a simple PNG image (1x1 pixel) to be used as an external resource
            // Base64 for a 1x1 transparent PNG
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(tempImageFile, pngBytes);

            // Create a workbook and insert the external image
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Demo with external image");
            int pictureIndex = ws.Pictures.Add(2, 0, tempImageFile);
            ws.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;

            // Save the workbook to a temporary Excel file (source for conversion)
            workbook.Save(sourceFile);

            // Load options for the source Excel file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Configure PDF save options to manage external resources
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Folder for temporary cache files generated during conversion
                CachedFileFolder = cachedFolder,
                // Embed external attachments (e.g., OLE objects) into the PDF
                EmbedAttachments = true,
                // Hide rendering errors (e.g., missing images) to avoid conversion failure
                IgnoreError = true
            };

            // Convert the Excel file to PDF using the ConversionUtility method
            ConversionUtility.Convert(sourceFile, loadOptions, destFile, pdfOptions);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
        finally
        {
            // Clean up temporary files and folders
            if (File.Exists(sourceFile)) File.Delete(sourceFile);
            if (File.Exists(destFile)) File.Delete(destFile);
            if (File.Exists(tempImageFile)) File.Delete(tempImageFile);
            if (Directory.Exists(cachedFolder))
            {
                try { Directory.Delete(cachedFolder, true); } catch { /* ignore cleanup errors */ }
            }
        }
    }
}