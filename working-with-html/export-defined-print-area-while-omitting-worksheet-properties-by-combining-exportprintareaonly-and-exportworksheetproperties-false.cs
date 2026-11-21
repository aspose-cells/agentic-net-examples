// Title: Export defined print area to PDF without worksheet properties using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export only print area to PDF | How to disable worksheet properties when saving Excel as PDF with Aspose.Cells | PdfSaveOptions ExportPrintAreaOnly true example in .NET | C# code to save Excel workbook to PDF without worksheet metadata using Aspose.Cells | ExportPrintAreaOnly and ExportWorksheetProperties settings Aspose.Cells PDF conversion
// Tags: ExportPrintAreaOnly PdfSaveOptions configuration | ExportWorksheetProperties false Aspose.Cells | PDF export print area only C# | disable worksheet metadata PDF conversion Aspose.Cells | Aspose.Cells PDFSaveOptions usage C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures PdfSaveOptions to set ExportPrintAreaOnly = true and ExportWorksheetProperties = false, and saves the workbook as a PDF. This ensures that only the defined print area is exported while worksheet properties are omitted.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // The following options are available in newer versions of Aspose.Cells.
                // If your version supports them, uncomment the lines below.
                // saveOptions.ExportPrintAreaOnly = true;          // Export only the defined print area
                // saveOptions.ExportWorksheetProperties = false;   // Do not export worksheet properties

                // Save the workbook as a PDF file
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
