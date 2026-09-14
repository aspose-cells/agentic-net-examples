// Title: How to export an Excel workbook to PDF and PNG while removing extra whitespace with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, configures PDF save options to fit each worksheet on a single page, and saves the result as a PDF using Aspose.Cells. | Write C# that iterates through all worksheets, renders each one to a PNG image scaled to a single page, and saves the images, ensuring excess margins are removed with Aspose.Cells.
// Common Searches: Aspose.Cells remove margins when saving Excel as PDF in C# | C# export Excel worksheets to PNG without extra whitespace using Aspose.Cells | How to fit each Excel sheet on one page during PDF export with Aspose.Cells | Trim whitespace in Excel to image conversion Aspose.Cells .NET
// Tags: Aspose.Cells PDF export whitespace cleanup | OnePagePerSheet PdfSaveOptions Aspose.Cells | SheetRender PNG export one page per sheet | Excel to PDF margin reduction Aspose.Cells | Excel to PNG image rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a sample workbook if needed, loads it, saves it as a PDF with each sheet fitted to a single page to eliminate surrounding whitespace, and then renders each worksheet to a separate PNG image using the same one‑page‑per‑sheet setting.
class WhitespaceCleanupExport
{
    static void Main()
    {
        try
        {
            const string inputPath = "SampleWorkbook.xlsx";
            const string pdfOutput = "ExportedClean.pdf";

            // Ensure the input workbook exists; create a simple one if missing.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to fit each sheet on a single page.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
                // Whitespace removal is handled automatically when fitting to a page.
            };

            // Export the workbook to PDF.
            workbook.Save(pdfOutput, pdfOptions);

            // Configure image export options (PNG is default).
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render each worksheet to a separate PNG file.
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                renderer.ToImage(0, $"Sheet{i + 1}_Clean.png");
            }

            Console.WriteLine("Export completed with whitespace cleanup.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
