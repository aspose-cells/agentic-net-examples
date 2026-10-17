// Title: How to enable PdfSaveOptions.ScaleCrop and keep the setting when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PdfSaveOptions.ScaleCrop true before saving workbook as PDF | preserve ScaleCrop flag after exporting Excel to PDF using Aspose.Cells .NET | PdfSaveOptions.ScaleCrop not retained in output PDF Aspose.Cells example | C# Aspose.Cells export Excel to PDF with scaling options enabled
// Tags: PdfSaveOptions.ScaleCrop true | Aspose.Cells Excel to PDF scaling | C# preserve PDF scaling options | Aspose.Cells PDF export settings | ScaleCrop flag persistence

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, sets PdfSaveOptions.ScaleCrop to true (along with any other desired PDF options), and saves the workbook as a PDF, ensuring the ScaleCrop setting is retained in the generated file.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (e.g., fit each sheet to a single page)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
