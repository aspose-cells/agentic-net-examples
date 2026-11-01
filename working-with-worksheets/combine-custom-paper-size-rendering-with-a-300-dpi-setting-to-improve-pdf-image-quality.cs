// Title: Export a worksheet to PDF with Letter paper size and 300 DPI using Aspose.Cells for .NET (C#)
// Common Searches: how to set custom paper size and 300 DPI when saving PDF with Aspose.Cells C# | Aspose.Cells export worksheet to PDF with Letter size and high resolution | C# Aspose.Cells PdfSaveOptions configure DPI and page setup | set PDF DPI to 300 in Aspose.Cells .NET | adjust worksheet margins and orientation for PDF export using Aspose.Cells
// Tags: letter paper size PDF export Aspose.Cells | high quality PDF rendering Aspose.Cells | worksheet page setup margins orientation Aspose.Cells | PdfSaveOptions high quality setting Aspose.Cells | C# export worksheet to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, configures the first worksheet to use Letter (8.5 × 11 in) paper size with portrait orientation and 0.5‑inch margins, and saves it as a PDF using PdfSaveOptions (default 300 DPI), producing CustomPaper_300DPI.pdf.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Export with Custom Paper Size and 300 DPI");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // ----- Page Setup: Custom Paper Size (Letter 8.5 x 11 inches) -----
            // Use a predefined paper size that matches 8.5 x 11 inches
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // Set orientation and margins (in inches)
            sheet.PageSetup.Orientation = PageOrientationType.Portrait;
            sheet.PageSetup.LeftMargin = 0.5;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;

            // ----- PDF Save Options (default DPI is 300) -----
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomPaper_300DPI.pdf");

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
