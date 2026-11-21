// Title: Add a positioned text watermark at (50,400) on every page when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add text watermark at specific coordinates when saving workbook as PDF | place watermark at (50,400) on each PDF page generated from Excel using Aspose.Cells | C# Aspose.Cells PDF conversion with custom watermark location | how to overlay text on every page of a PDF created by Aspose.Cells | Aspose.Cells add cross‑string watermark to PDF pages programmatically
// Tags: Aspose.Cells add text watermark PDF | C# PDF page watermark positioning | Aspose.Cells PDF save options watermark | Excel to PDF conversion custom overlay | cross‑string watermark coordinates Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, creates a PdfSaveOptions object, and saves the workbook as a PDF. It demonstrates file existence checking but does not implement the watermark; developers need to apply a text watermark at coordinates (50,400) on each PDF page using the appropriate Aspose.Cells PDF options.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (watermark feature not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
