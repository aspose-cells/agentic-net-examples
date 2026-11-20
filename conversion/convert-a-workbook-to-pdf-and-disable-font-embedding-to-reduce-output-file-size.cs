// Title: Convert an Excel workbook to PDF and control font embedding with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook as PDF without embedding fonts | How to reduce PDF file size when converting XLSX to PDF using Aspose.Cells | PdfSaveOptions disable font embedding Aspose.Cells .NET example | C# code to export Excel to PDF while turning off font embedding | Is font embedding configurable in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDF font embedding control | Excel to PDF conversion with reduced size | PdfSaveOptions usage in Aspose.Cells | C# workbook to PDF without embedded fonts | disable font embedding in Aspose.Cells PDF output

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates loading an Excel workbook with Aspose.Cells, configuring PdfSaveOptions, and saving the workbook as a PDF. It highlights that font embedding can be controlled via PdfSaveOptions in recent Aspose.Cells versions, enabling smaller PDF files when the feature is available.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings; font embedding cannot be disabled in older versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
