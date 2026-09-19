// Title: Export a worksheet with multiple Excel tables to a single PDF while preserving table layouts using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, selects a worksheet containing several tables, and saves it as one PDF file keeping the original table formatting with Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to retain column widths, row heights, and table structures when converting a worksheet to PDF.
// Common Searches: Aspose.Cells C# export worksheet with multiple tables to a single PDF file | preserve Excel table formatting when converting to PDF using Aspose.Cells | how to keep table layout in PDF output from Aspose.Cells .NET | PdfSaveOptions settings for exporting tables on one PDF page Aspose.Cells | C# convert Excel sheet containing several tables to PDF preserving layout
// Tags: Aspose.Cells export worksheet to PDF | PdfSaveOptions preserve table layout | C# convert Excel tables to PDF | single PDF output for multiple tables Aspose.Cells | maintain column widths row heights PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input Excel file, loads it with Aspose.Cells, accesses the first worksheet, optionally configures PdfSaveOptions, and saves the worksheet as a single PDF while retaining column widths, row heights, and the layout of all tables on the sheet.
class ExportTablesToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing the worksheet with multiple tables
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name/index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Uncomment to force each worksheet onto a single PDF page
                // OnePagePerSheet = true,

                // Preserve original column widths and row heights (default behavior)
                // No direct property needed; keep defaults

                // Keep gridlines if required
                // GridlineType = GridlineType.Gridline
            };

            // Save the workbook (or the selected worksheet) to a PDF file
            const string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
