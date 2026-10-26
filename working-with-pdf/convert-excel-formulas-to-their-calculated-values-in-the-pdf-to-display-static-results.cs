// Title: Convert Excel formulas to static values when exporting to PDF using Aspose.Cells for .NET (C#)
// Common Searches: how to export an .xlsx to PDF with calculated formula results using Aspose.Cells C# | Aspose.Cells C# generate PDF that shows formula values instead of formulas | C# save Excel workbook as PDF with formulas evaluated to static numbers | force formula calculation before PDF conversion with Aspose.Cells .NET | use PdfSaveOptions to keep calculated values when converting Excel to PDF
// Tags: calculate formulas before PDF export Aspose.Cells | PdfSaveOptions static values Excel to PDF | Workbook.CalculateFormula method C# | AllColumnsInOnePagePerSheet PDF layout Aspose.Cells | export Excel workbook as PDF with evaluated formulas

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The example loads an Excel file with Aspose.Cells, ensures all formulas are evaluated using Workbook.CalculateFormula(), configures PdfSaveOptions (e.g., AllColumnsInOnePagePerSheet), and saves the workbook as a PDF where the cells contain the calculated static values instead of the original formulas.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel workbook
                var workbook = new Workbook(inputPath);

                // Ensure all formulas are calculated before exporting
                workbook.CalculateFormula();

                // Configure PDF save options (optional settings can be added as needed)
                var pdfOptions = new PdfSaveOptions
                {
                    // Example: fit all columns on one page per sheet (adjust as required)
                    AllColumnsInOnePagePerSheet = true
                };

                // Save the workbook as a PDF; the PDF will contain the calculated (static) values
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
