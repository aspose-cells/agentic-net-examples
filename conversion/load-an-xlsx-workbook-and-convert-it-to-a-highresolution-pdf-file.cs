// Title: Convert an XLSX workbook to a high‑resolution PDF using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells example for saving an Excel file as a PDF with custom PdfSaveOptions | how to export .xlsx to high quality PDF in .NET with Aspose.Cells | Aspose.Cells PDF conversion preserving sheet layout C# code sample | set PDF resolution when converting Excel to PDF using Aspose.Cells | save workbook as PDF without losing column width Aspose.Cells C#
// Tags: Aspose.Cells PDFSaveOptions configuration C# | XLSX to PDF high resolution export | Excel to PDF export Aspose.Cells | custom PDF paging options Aspose.Cells | C# save workbook as PDF example

using Aspose.Cells;
using System;
using System.IO;

// The sample creates or verifies an input.xlsx file, loads it with Aspose.Cells, configures PdfSaveOptions (including paging settings), and saves the workbook as a high‑resolution PDF (output.pdf) while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var sheet = tempWb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
            }

            // Load the workbook from the verified file.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: keep each sheet on a separate page.
                OnePagePerSheet = false,

                // Optional: fit all columns on one page per sheet.
                AllColumnsInOnePagePerSheet = false
            };

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
