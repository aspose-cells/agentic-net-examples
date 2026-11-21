// Title: Export an Excel worksheet to PDF using Aspose.Cells for .NET (C#) while preserving vector graphics
// Common Searches: how to export an .xlsx worksheet to PDF with vector graphics using Aspose.Cells in C# | Aspose.Cells C# example for saving workbook as PDF preserving SVG images | convert Excel file to PDF without rasterizing charts Aspose.Cells .NET | C# code to generate PDF from Excel while keeping graphics as vectors Aspose.Cells | save Excel workbook as PDF with high‑quality vector output using Aspose.Cells
// Tags: Aspose.Cells export worksheet to PDF C# | preserve vector graphics in PDF export Aspose.Cells | Excel to PDF conversion with Aspose.Cells .NET | save workbook as PDF preserving SVG Aspose.Cells | high‑quality PDF generation from XLSX Aspose.Cells | Aspose.Cells PDF output with vector images C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file with Aspose.Cells, optionally activates the first worksheet, and saves the workbook as a PDF document, ensuring that any SVG or vector graphics in the worksheet are retained as vector elements in the resulting PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Set the first worksheet as active (optional, default is 0)
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
