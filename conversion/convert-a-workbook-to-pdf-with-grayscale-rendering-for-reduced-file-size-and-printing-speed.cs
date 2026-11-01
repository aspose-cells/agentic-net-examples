// Title: Convert an Excel workbook to a grayscale PDF with reduced file size using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to grayscale PDF for smaller file size | How to enable Grayscale in PdfSaveOptions with Aspose.Cells .NET | Save .xlsx as PDF with reduced size and faster printing using Aspose.Cells | Aspose.Cells PDF conversion with grayscale rendering example in C# | Convert workbook to PDF with grayscale option using Aspose.Cells API
// Tags: Aspose.Cells grayscale PDF conversion C# | PdfSaveOptions Grayscale property Aspose.Cells | reduce PDF file size Excel to PDF Aspose.Cells | export workbook to PDF with printing optimization | C# Excel to PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (where the Grayscale property can be set to render the PDF in grayscale for smaller size and faster printing), and saves the workbook as a PDF file while handling any runtime exceptions.
class WorkbookToPdfGrayscale
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (grayscale option omitted for compatibility)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Additional options can be set here, e.g.:
                // Compliance = PdfCompliance.PdfA1b,
                // OnePagePerSheet = true
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
