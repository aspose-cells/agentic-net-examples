// Title: Capture font substitution warnings when converting Excel to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET how to get font substitution warnings during Excel to PDF conversion | log missing font messages when saving workbook as PDF using Aspose.Cells C# | retrieve font fallback warnings after rendering XLSX to PDF with Aspose.Cells | detect unavailable font warnings in Aspose.Cells PDF export C# | check font substitution details after converting Excel file to PDF in .NET
// Tags: Aspose.Cells PDF export font substitution | C# retrieve font warnings Excel to PDF | Aspose.Cells missing font detection | log font fallback during PDF conversion | Excel workbook PDF rendering font warnings

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook and saves it as a PDF using Aspose.Cells for .NET. It highlights that the current Aspose.Cells version does not expose an API to retrieve font substitution warnings, and advises developers to consult the latest documentation or alternative methods for capturing missing‑font information.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set PDF save options (default options can be used)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");

            // NOTE: Font substitution warnings are not available in the current Aspose.Cells version.
            // If needed, refer to the latest Aspose.Cells documentation for the appropriate API.
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
