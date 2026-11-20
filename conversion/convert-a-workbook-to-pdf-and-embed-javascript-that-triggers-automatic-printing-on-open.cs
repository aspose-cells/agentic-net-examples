// Title: Convert an Excel workbook to PDF with Aspose.Cells in C# (JavaScript auto‑print not supported)
// Common Searches: c# convert xlsx to pdf using aspose.cells | how to embed javascript print command in pdf created by aspose.cells | asp.net save workbook as pdf with pdfsaveoptions | generate pdf from excel and trigger print on open in c# | aspose.cells pdfsaveoptions does not support javascript
// Tags: Aspose.Cells workbook to PDF conversion C# | PdfSaveOptions usage Aspose.Cells | JavaScript embedding limitation Aspose.Cells PDF | auto‑print script in generated PDF .NET | Excel to PDF generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, configures PdfSaveOptions, and saves the workbook as a PDF. Aspose.Cells for .NET does not provide a way to embed JavaScript (such as an auto‑print script) into the generated PDF, so the PDF is created without any print‑on‑open functionality.
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

            // Configure PDF save options (Aspose.Cells does not expose a JavaScript property)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
