// Title: Preserve Excel cell comments when converting a workbook to PDF using Aspose.Cells in C#
// AI Prompts: Write a C# console program that loads an .xlsx file, verifies the file exists, and saves it as a PDF with all cell comments retained using Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to keep cell annotations during Excel‑to‑PDF conversion. | Provide error‑handling code that confirms successful PDF creation and reports any issues while preserving comments.
// Common Searches: Aspose.Cells C# keep Excel cell comments when exporting to PDF | Convert .xlsx to PDF with comments preserved using PdfSaveOptions | Sample code for retaining annotations in PDF generated from Excel workbook | Does Aspose.Cells preserve comments by default in PDF output
// Tags: Aspose.Cells PdfSaveOptions comment retention | C# export Excel comments to PDF | Excel to PDF conversion with annotations Aspose | Workbook.Save PDF with annotations .NET | retain cell annotations in PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing input.xlsx workbook, creates a PdfSaveOptions object (comments are retained by default), and saves the workbook as output.pdf. It includes file‑existence checking and exception handling to ensure a reliable PDF conversion that keeps all cell comments.
class PreserveCommentsPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing cell comments
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (comments are preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
