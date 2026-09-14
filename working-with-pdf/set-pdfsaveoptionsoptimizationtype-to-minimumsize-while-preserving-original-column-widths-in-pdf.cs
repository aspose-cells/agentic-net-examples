// Title: Export an Excel workbook to PDF with MinimumSize optimization while keeping original column widths using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets PdfSaveOptions.OptimizationType to MinimumSize, and saves the workbook as a PDF without altering the column widths. | Show a complete example that checks for the input file, creates the output folder, configures PdfSaveOptions for smallest PDF size, and preserves the worksheet layout when exporting with Aspose.Cells.
// Common Searches: Aspose.Cells C# set PDF OptimizationType MinimumSize keep column widths | How to reduce PDF size when converting Excel to PDF with Aspose.Cells without changing layout | C# export Excel to PDF with minimum file size and original column widths | PdfSaveOptions MinimumSize effect on column width preservation in Aspose.Cells | Save workbook as PDF with smallest size using Aspose.Cells .NET
// Tags: PdfSaveOptions MinimumSize | Aspose.Cells preserve column widths PDF | C# Excel to PDF size reduction | Aspose.Cells PDF export layout fidelity | configure PDF save options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the source Excel file exists, loads it into a Workbook, creates PdfSaveOptions with OptimizationType set to MinimumSize, ensures the target directory is present, and saves the workbook as a PDF while preserving the original column widths. Errors are caught and reported.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient for most cases)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF using the options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
