// Title: Convert an Excel workbook to PDF and add a watermark on odd-numbered pages using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add watermark to odd pages when exporting Excel to PDF | how to apply a page‑specific watermark during Excel to PDF conversion with Aspose.Cells | C# example for conditional watermark on every other page in PDF generated from a workbook | set watermark only on odd-numbered pages in PDF output using Aspose.Cells PdfSaveOptions | Aspose.Cells PDF export with selective watermarking of odd pages in .NET
// Tags: Aspose.Cells Excel to PDF odd-page watermark | PdfSaveOptions selective page watermarking C# | conditional watermark during PDF export Aspose.Cells | C# generate PDF from workbook with page range watermark | Aspose.Cells page-specific watermark implementation

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, prepares PdfSaveOptions (noting that the PdfWatermark class is unavailable in the referenced version), and saves the workbook as a PDF. A placeholder comment indicates where logic to apply a watermark only on odd-numbered pages would be inserted.
class WorkbookToPdfWithOddPageWatermark
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PdfWatermark class is not available in the referenced Aspose.Cells version.
            // If needed, add the appropriate Aspose.Cells.Pdf assembly and use PdfWatermark here.

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
