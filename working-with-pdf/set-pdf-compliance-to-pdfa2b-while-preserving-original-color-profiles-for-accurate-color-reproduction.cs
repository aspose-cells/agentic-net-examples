// Title: Convert an Excel workbook to PDF/A‑2b compliant PDF while preserving color profiles using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert .xlsx to PDF/A-2b preserving original color profiles | How to set PDF compliance to PDF/A-2b with Aspose.Cells PdfSaveOptions in .NET | Save Excel workbook as PDF/A-2b with accurate colors using Aspose.Cells example | Aspose.Cells PDF/A-2b export keeping embedded color profiles from Excel file | PdfSaveOptions Compliance property PDF/A-2b example for Excel to PDF conversion in C#
// Tags: Aspose.Cells PdfSaveOptions PDF/A-2b compliance | Excel workbook to PDF/A-2b export C# | retain embedded color profiles during PDF conversion | set PdfCompliance.PdfA2b in Aspose.Cells | C# generate PDF/A-2b compliant document from Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, configures PdfSaveOptions with the Compliance property set to PdfCompliance.PdfA2b, and saves the file as a PDF/A‑2b document. Aspose.Cells automatically retains the workbook's original color profiles, ensuring accurate color reproduction in the resulting PDF.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set PDF/A‑2b compliance
                Compliance = PdfCompliance.PdfA2b
            };

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
