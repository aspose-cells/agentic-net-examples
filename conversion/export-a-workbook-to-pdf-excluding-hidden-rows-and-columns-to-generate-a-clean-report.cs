// Title: Export an Excel workbook to PDF while omitting hidden rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to PDF without hidden rows C# | How to ignore hidden columns when saving a workbook as PDF in .NET | C# generate PDF report from .xlsx excluding hidden cells using Aspose | PdfSaveOptions hide rows and columns Aspose.Cells example | Save Excel workbook as clean PDF report with hidden data removed in C#
// Tags: Aspose.Cells PDF export hide rows | PdfSaveOptions exclude hidden columns | C# Excel to PDF clean report | Workbook.Save PDF without hidden data | Aspose.Cells generate PDF report from workbook

using System;
using System.IO;
using Aspose.Cells;

// The code checks for the existence of 'input.xlsx', loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which by default skips hidden rows and columns), and saves the workbook as 'report.pdf', handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "report.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default behavior excludes hidden rows/columns in recent versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF report generated successfully at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
