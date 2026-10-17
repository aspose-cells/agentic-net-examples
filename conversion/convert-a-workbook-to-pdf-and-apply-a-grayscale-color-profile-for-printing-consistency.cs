// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET (grayscale rendering unavailable)
// Common Searches: how to save an .xlsx file as PDF with Aspose.Cells in C# | Aspose.Cells PDFSaveOptions grayscale rendering limitation | C# convert Excel to PDF with printing color profile using Aspose.Cells | example of workbook.Save to PDF with error handling in Aspose.Cells | verify input Excel file exists before converting to PDF with Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions configuration | Aspose.Cells workbook to PDF conversion | grayscale PDF rendering limitation Aspose.Cells | Excel to PDF printing color profile Aspose.Cells | Aspose.Cells conversion exception handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The sample verifies that the source XLSX file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (noting that grayscale rendering is not supported in the current version), and saves the workbook as a PDF while handling any runtime exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (grayscale rendering not supported in this version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
