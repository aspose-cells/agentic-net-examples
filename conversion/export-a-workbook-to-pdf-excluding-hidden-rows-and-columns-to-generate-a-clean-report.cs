// Title: Export an Excel workbook to PDF while excluding hidden rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF ignoring hidden rows and columns | How to generate a clean PDF report from an XLSX file with Aspose.Cells | C# PdfSaveOptions to omit hidden rows during Excel to PDF conversion | Save workbook as PDF without hidden columns using Aspose.Cells .NET | Exclude hidden rows and columns when converting Excel to PDF in C#
// Tags: Aspose.Cells PdfSaveOptions hide rows and columns | C# Excel workbook PDF export without hidden elements | Aspose.Cells clean PDF generation from XLSX | exclude hidden rows PDF conversion Aspose.Cells | hidden columns omitted in PDF save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, configures PdfSaveOptions (noting that recent versions automatically exclude hidden rows and columns), and saves the workbook as a PDF, producing a clean report while handling file‑not‑found and other runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // In newer Aspose.Cells versions hidden rows/columns are excluded by default.
            // If the property becomes available, it can be set as:
            // pdfOptions.HiddenRowsAndColumns = true;

            // Export the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
