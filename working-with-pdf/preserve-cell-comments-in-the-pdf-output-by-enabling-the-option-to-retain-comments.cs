// Title: How to preserve Excel cell comments when converting a workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel comments to PDF | keep worksheet cell notes in PDF using Aspose.Cells .NET | PdfSaveOptions retain comments during Excel to PDF conversion | preserve Excel annotations when saving as PDF with Aspose.Cells | export Excel cell comments as PDF annotations using C#
// Tags: Aspose.Cells PDF export comments | C# PdfSaveOptions retain cell notes | Excel to PDF with annotations Aspose | preserve worksheet comments .NET | export cell comments as PDF annotations

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions object (which exports cell comments by default), and saves the workbook as a PDF, ensuring that any worksheet comments are retained in the resulting PDF file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (comments are exported by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
