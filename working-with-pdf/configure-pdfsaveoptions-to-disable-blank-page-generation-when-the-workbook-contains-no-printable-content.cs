// Title: C# – Using Aspose.Cells PdfSaveOptions to prevent blank pages when converting an empty Excel workbook to PDF
// Common Searches: C# Aspose.Cells how to stop blank page in PDF when sheet has no data | PdfSaveOptions setting to avoid empty page generation in Aspose.Cells conversion | Convert Excel workbook to PDF without extra blank page using Aspose.Cells .NET | Aspose.Cells PDF export blank page issue when workbook contains no printable content | Skip generating a blank PDF page when saving an empty workbook with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions suppress empty PDF page | C# Excel to PDF conversion without blank pages | PDF export settings for empty worksheets Aspose.Cells | disable blank page generation Aspose.Cells PDF | Aspose.Cells PDF output skip empty content

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, creates a PdfSaveOptions object, and saves the workbook as a PDF. It highlights that the current Aspose.Cells version lacks a direct property to disable blank page creation, guiding developers on how to configure PDF export when the workbook has no printable content.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: The DisablePageBreaks property is not available in the current Aspose.Cells version.
            // If needed, other options (e.g., OnePagePerSheet) can be set here.

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
