// Title: Convert selected Excel ranges to PDF while skipping hidden rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export multiple non‑contiguous ranges to PDF in C# | How to ignore hidden rows and columns when saving Excel as PDF with Aspose.Cells | Set print area for several ranges before PDF conversion using Aspose.Cells .NET | C# code to save specific worksheet ranges to PDF with Aspose.Cells | Aspose.Cells PDF conversion exclude hidden rows columns example
// Tags: selected ranges PDF export Aspose.Cells | exclude hidden rows columns PDF Aspose.Cells | print area multiple ranges Aspose.Cells | PdfSaveOptions OnePagePerSheet false | C# Aspose.Cells Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, defines comma‑separated cell ranges (e.g., "A1:C10,E5:G15"), assigns them to the worksheet's PrintArea, configures PdfSaveOptions to keep the original layout, ensures the output folder exists, and saves only the specified ranges as a PDF. Hidden rows and columns are automatically omitted during the conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the ranges to export (comma‑separated)
            string selectedRanges = "A1:C10,E5:G15";

            // Apply the print area to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintArea = selectedRanges;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false // keep original layout
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the selected ranges to a PDF file using the options overload
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
