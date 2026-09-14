// Title: Convert specific worksheet ranges to PDF while ignoring hidden rows and columns using Aspose.Cells for .NET
// AI Prompts: Write C# code that defines multiple non‑contiguous print areas on a worksheet and saves only those areas to a PDF with Aspose.Cells, automatically omitting any hidden rows or columns. | Show how to configure Aspose.Cells PdfSaveOptions to export selected ranges to PDF and ensure hidden rows and columns are not rendered in the output.
// Common Searches: Aspose.Cells C# export only visible cells from selected ranges to PDF | how to set multiple print areas for PDF conversion with Aspose.Cells .NET | exclude hidden rows and columns when saving Excel to PDF using Aspose.Cells | C# convert non‑contiguous worksheet ranges to PDF with Aspose.Cells
// Tags: Aspose.Cells set print area PDF | export selected Excel ranges to PDF C# | skip hidden rows columns Aspose.Cells | multiple non‑contiguous print areas .NET | PdfSaveOptions hide hidden rows

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, assigns multiple non‑contiguous ranges (e.g., A1:C10 and E1:G5) as the print area on the first worksheet, configures PdfSaveOptions, and saves only the visible cells of those ranges to a PDF, automatically ignoring any hidden rows or columns.
class ConvertSelectedRangesToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the ranges you want to convert.
            // Multiple ranges can be combined using commas.
            // Example: Convert A1:C10 and E1:G5
            sheet.PageSetup.PrintArea = "A1:C10,E1:G5";

            // Set up PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: set other PDF options as required
                // OnePagePerSheet = true,
            };

            // Save the selected ranges as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
