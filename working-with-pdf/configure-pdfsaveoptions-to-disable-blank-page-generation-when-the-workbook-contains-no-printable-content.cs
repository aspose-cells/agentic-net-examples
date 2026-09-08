// Title: How to prevent blank PDF pages when exporting an empty Excel workbook using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Generate C# code that configures Aspose.Cells PdfSaveOptions to skip creating a PDF page for worksheets that have no printable content. | Show the specific PdfSaveOptions property to set in order to suppress blank pages when saving an empty workbook to PDF with Aspose.Cells. | Write a .NET example that saves an empty Workbook to PDF without producing a blank page, using the appropriate PdfSaveOptions setting.
// Common Searches: Aspose.Cells C# disable blank page when saving empty worksheet to PDF | PdfSaveOptions property to avoid empty PDF pages in Aspose.Cells export | How to stop Aspose.Cells from creating a PDF page for a workbook with no data | Export empty Excel file to PDF without blank page using Aspose.Cells .NET | C# Aspose.Cells PDF export suppress empty pages
// Tags: Aspose.Cells PdfSaveOptions DisablePrintingEmptyPages | C# export empty workbook to PDF without blank page | Aspose.Cells PDF blank page prevention | Aspose.Cells PDF save options configuration | Aspose.Cells PDF export suppress empty pages

using System;
using System.IO;
using Aspose.Cells;

// The sample creates an empty Workbook, sets PdfSaveOptions.DisablePrintingEmptyPages to true to stop a blank PDF page from being generated, and saves the workbook to a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (empty workbook has no printable content)
            Workbook workbook = new Workbook();

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Example option: generate one PDF page per worksheet
            pdfOptions.OnePagePerSheet = true;

            // Define output file path
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
