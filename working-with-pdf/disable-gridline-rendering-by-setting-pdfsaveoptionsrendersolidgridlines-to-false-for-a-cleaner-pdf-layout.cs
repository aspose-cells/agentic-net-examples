// Title: Export an Excel workbook to a clean PDF without gridlines using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, suppresses worksheet gridlines, and saves the workbook as a PDF. | Modify a C# Aspose.Cells PDF export routine so that the generated PDF contains no gridlines by turning off gridline visibility before saving.
// Common Searches: Aspose.Cells C# hide gridlines when converting Excel to PDF | export Excel workbook to PDF without gridlines using Aspose.Cells .NET | remove worksheet gridlines in PDF output Aspose.Cells C# example
// Tags: Aspose.Cells suppress worksheet gridlines | C# generate PDF from Excel without gridlines | PdfSaveOptions RenderSolidGridlines false | clean PDF layout using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // Loads an Excel file, sets IsGridlinesVisible = false for each worksheet, and saves the workbook as a PDF with default PdfSaveOptions, producing a PDF without visible gridlines.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Ensure the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Hide gridlines for all worksheets (cleaner PDF layout)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.IsGridlinesVisible = false;
                }

                // Configure PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
