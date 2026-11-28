// Title: Save an Excel worksheet as HTML5 with images embedded as Base64 strings using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to HTML5 with embedded Base64 images | how to embed picture in HTML output using Aspose.Cells HtmlSaveOptions | C# generate HTML from Excel and include images as Base64 data URIs | set ExportImagesAsBase64 true in Aspose.Cells example code | save Excel sheet as HTML with gridlines and Base64 images in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to HTML5 with Base64 images | embed picture as Base64 data URI Aspose.Cells | free-floating picture placement Aspose.Cells | ExportActiveWorksheetOnly option

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for PlacementType enum

// Creates a workbook with sample data and a PNG picture, then saves the active worksheet to an HTML5 file where all images are embedded directly as Base64 strings by enabling ExportImagesAsBase64 in HtmlSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some sample data.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Widget");
            sheet.Cells["B2"].PutValue(19.99);
            sheet.Cells["A3"].PutValue("Gadget");
            sheet.Cells["B3"].PutValue(29.99);

            // -----------------------------------------------------------------
            // 2. Insert a simple PNG image (embedded as a byte array) into the sheet.
            // -----------------------------------------------------------------
            // A tiny red dot PNG (1x1 pixel) encoded in Base64.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAFklEQVR42mNk+M9Qz0AEYBxVSFIAABcAAf+V6nUAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);

            // Write the image to a temporary file (required by the Add method that uses a file path).
            string tempImagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");
            File.WriteAllBytes(tempImagePath, pngBytes);

            // Insert the image at row 5, column 0 (cell A6) as a free‑floating picture.
            int pictureIndex = sheet.Pictures.Add(5, 0, tempImagePath);
            sheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;

            // Optionally delete the temporary file after insertion.
            if (File.Exists(tempImagePath))
            {
                try { File.Delete(tempImagePath); } catch { /* ignore cleanup errors */ }
            }

            // -----------------------------------------------------------------
            // 3. Configure HTML save options to embed images as Base64 strings.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,          // Embed images directly as Base64.
                ExportActiveWorksheetOnly = true,     // Export only the prepared sheet.
                ExportGridLines = true,               // Keep grid lines visible (optional).
                HtmlVersion = HtmlVersion.Html5       // Generate HTML5 output (optional).
            };

            string outputHtmlFile = "WorkbookWithEmbeddedImages.html";

            // -----------------------------------------------------------------
            // 4. Save the workbook as HTML using the configured options.
            // -----------------------------------------------------------------
            workbook.Save(outputHtmlFile, htmlOptions);

            Console.WriteLine($"HTML file '{outputHtmlFile}' has been created.");
            Console.WriteLine("All images are embedded directly as Base64 strings.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
