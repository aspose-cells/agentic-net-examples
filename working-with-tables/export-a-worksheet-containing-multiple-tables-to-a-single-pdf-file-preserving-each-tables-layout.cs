// Title: Export a worksheet with multiple Excel tables to a single PDF while preserving table layouts using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet containing several tables to one PDF | preserve Excel table formatting when converting to PDF with Aspose.Cells | fit worksheet width to a single PDF page using Aspose.Cells C# | disable OnePagePerSheet option in Aspose.Cells PDF export | create PDF/A-1b from Excel workbook with multiple tables using Aspose.Cells
// Tags: worksheet to PDF export Aspose.Cells | maintain table layout during PDF conversion | page setup fit width single PDF page | PdfSaveOptions OnePagePerSheet false setting | PDF/A-1b generation with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, selects a worksheet that contains several tables, configures the page setup to fit the worksheet width to one PDF page, sets PdfSaveOptions with OnePagePerSheet disabled and PDF/A‑1b compliance, and saves the result as a single PDF file while keeping each table's layout intact. Includes basic file‑existence checking and exception handling.
class ExportTablesToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the worksheet with multiple tables.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify the required one).
            Worksheet worksheet = workbook.Worksheets[0];

            // Adjust page setup: fit width to one page, height unrestricted.
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.FitToPagesWide = 1;   // Fit width to one page.
            pageSetup.FitToPagesTall = 0;   // No restriction on height.

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,                 // Do not force everything onto a single page.
                Compliance = PdfCompliance.PdfA1b        // Optional: set PDF/A compliance.
                // Additional options (e.g., grid lines) are not available in this version of the API.
            };

            // Save the workbook (or the selected worksheet) as a PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
