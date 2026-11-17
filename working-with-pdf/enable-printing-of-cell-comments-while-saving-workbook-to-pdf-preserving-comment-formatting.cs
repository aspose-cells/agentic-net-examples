// Title: Print cell comments while converting an Excel workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel comments to PDF | include printable cell comments in PDF output with Aspose.Cells | preserve Excel comment formatting when saving as PDF in .NET | how to enable comment rendering in PdfSaveOptions Aspose.Cells | convert workbook to PDF with cell comments displayed using Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions include comments | export Excel cell comments to PDF .NET | preserve comment formatting in PDF conversion | printable comments Aspose.Cells PDF export | save workbook as PDF with comments

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, configures PdfSaveOptions (which by default includes printable comments), ensures the output directory exists, and saves the workbook as a PDF while preserving the original formatting of any cell comments.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options (default options will include comments if they are set to be printed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
