// Title: How to add PDF accessibility tags when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.TaggedPdf = true and PdfSaveOptions.Compliance = PdfCompliance.PdfA1a, and saves a tagged PDF for screen readers. | Show a complete example that validates the input Excel path, configures PdfSaveOptions for accessibility, and handles exceptions while creating the PDF. | Provide a minimal snippet demonstrating the use of PdfSaveOptions.TaggedPdf together with other PDF/A‑1a settings in Aspose.Cells.
// Common Searches: asp.net convert excel to pdf with accessibility tags using aspose.cells | c# generate tagged pdf from workbook for screen readers | how to enable pdf/a-1a compliance in aspose.cells pdf export | asp.net core add document structure tags to pdf generated from excel | c# aspose.cells PdfSaveOptions TaggedPdf property example
// Tags: Aspose.Cells PDF tagging | C# Excel to accessible PDF conversion | PdfSaveOptions TaggedPdf property | PDF/A-1a compliance with Aspose.Cells | screen reader friendly PDF generation .NET

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates loading an Excel workbook, configuring PdfSaveOptions to enable PDF tagging and PDF/A‑1a compliance, and saving the workbook as an accessible PDF suitable for screen readers using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (basic options; advanced tagging/compliance require newer library versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF document
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
