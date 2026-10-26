// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET – default chart rendering (no vector option)
// Common Searches: how to export an Excel workbook to PDF with Aspose.Cells while keeping chart quality | Aspose.Cells C# PDFSaveOptions chart rendering settings | convert Excel file to PDF using Aspose.Cells .NET and preserve charts | Aspose.Cells PDF export does not support vector chart rendering | save workbook as PDF with charts using Aspose.Cells for .NET
// Tags: Aspose.Cells PDF conversion chart rendering | C# Aspose.Cells PdfSaveOptions usage | Excel to PDF export preserving chart fidelity | default chart rendering in Aspose.Cells PDF export | Aspose.Cells .NET PDF export without vector charts

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, creates a PdfSaveOptions object (noting that the RenderChartsAsVector property is unavailable in this version), and saves the workbook as a PDF. The process demonstrates default chart rendering during PDF conversion and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input Excel file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default rendering; RenderChartsAsVector not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
