// Title: Configure PdfSaveOptions for PDF/A‑1b compliance and page layout when converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, creates a PdfSaveOptions object with OnePagePerSheet = true, Compliance = PdfCompliance.PdfA1b, AllColumnsInOnePagePerSheet = false, and saves the workbook as a PDF using Aspose.Cells. | Show how to apply PdfSaveOptions to enforce PDF/A‑1b compliance and control pagination (single page per sheet vs. column overflow) during Excel‑to‑PDF conversion with Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to PDF with PDF/A-1b compliance | How to set OnePagePerSheet in PdfSaveOptions for Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet false example in C# | Configure PDF export settings for workbook conversion using Aspose.Cells .NET | Aspose.Cells pagination options when saving workbook as PDF
// Tags: Aspose.Cells PdfSaveOptions PDF/A-1b compliance | Aspose.Cells one page per sheet PDF export | Aspose.Cells column pagination control PDF | Aspose.Cells workbook to PDF conversion C# | Aspose.Cells PDF export settings .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel file, configures PdfSaveOptions (OnePagePerSheet, PDF/A-1b compliance, column pagination), and saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Instantiate PdfSaveOptions for PDF export
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // Export each worksheet on a single PDF page
                Compliance = PdfCompliance.PdfA1b,    // Set PDF/A-1b compliance for archiving
                AllColumnsInOnePagePerSheet = false   // Allow multiple pages per sheet if needed
                // ImageResolution property is not available in this version; omitted.
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
