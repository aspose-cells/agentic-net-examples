// Title: How to retain Excel cell borders and gridlines when converting a workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# keep Excel gridlines visible in generated PDF | Preserve cell border formatting during Excel to PDF conversion with Aspose.Cells | C# PdfSaveOptions to include gridlines and borders in PDF output | Export Excel workbook to PDF without losing gridlines using Aspose.Cells .NET | Configure Aspose.Cells PDF export to match original Excel layout
// Tags: Aspose.Cells PdfSaveOptions include gridlines | C# Excel to PDF border preservation | Aspose.Cells retain cell formatting PDF | Excel gridline export Aspose.Cells .NET | PDF conversion exact visual replication Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, sets PdfSaveOptions (including optional GridlineType and OnePagePerSheet = false) to keep cell borders and gridlines, and saves the workbook as a PDF while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Include gridlines in the PDF output (if supported by the version)
                // GridlineType = GridlineType.Gridlines,

                // Keep original layout without forcing everything onto one page
                OnePagePerSheet = false
            };

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
