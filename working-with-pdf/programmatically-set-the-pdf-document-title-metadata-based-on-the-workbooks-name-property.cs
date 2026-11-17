// Title: Setting PDF document title metadata from an Excel workbook name using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PDF title to Excel workbook filename | convert XLSX to PDF with custom title metadata using Aspose.Cells | C# example for assigning BuiltInDocumentProperties.Title before PDF export | how to preserve workbook name as PDF document title in Aspose.Cells | Aspose.Cells PDFSaveOptions custom document properties tutorial
// Tags: Aspose.Cells PDF title metadata assignment | Workbook BuiltInDocumentProperties Title usage | Excel to PDF conversion with custom document properties | PdfSaveOptions configuration for PDF metadata | C# programmatic PDF metadata from workbook name

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, sets the BuiltInDocumentProperties.Title to the source file name (without extension), configures PdfSaveOptions, and saves the workbook as a PDF, ensuring the resulting PDF carries the workbook's name as its document title.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set PDF title metadata via workbook's built‑in document properties
            workbook.BuiltInDocumentProperties.Title = Path.GetFileNameWithoutExtension(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
