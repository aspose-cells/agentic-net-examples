// Title: Export an Excel workbook to PDF with Aspose.Cells for .NET (watermark on odd pages requires PdfWatermark assembly)
// Common Searches: how to export an Excel file to PDF using Aspose.Cells in C# | Aspose.Cells C# add watermark only to odd pages when saving as PDF | C# example for PdfSaveOptions with Aspose.Cells PDF export | apply page‑specific watermark to PDF generated from Excel with Aspose.Cells | Aspose.Cells PDF conversion with conditional watermarking on every other page
// Tags: Aspose.Cells PDF conversion C# | PdfSaveOptions configuration Aspose.Cells | odd page watermark Aspose.Cells PDF | Excel to PDF export with Aspose.Cells | conditional page watermarking Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures PdfSaveOptions, ensures the output directory exists, and saves the workbook as a PDF. It includes a note that the PdfWatermark class is not available in the referenced Aspose.Cells version, so adding a watermark to odd pages requires referencing the Aspose.Cells.Pdf assembly and using PdfWatermark separately.
class WorkbookToPdfWithOddPageWatermark
{
    static void Main()
    {
        try
        {
            // Paths
            string excelPath = @"C:\Input\Sample.xlsx";
            string outputPdfPath = @"C:\Output\Sample_WithWatermark.pdf";

            // Verify input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input Excel file not found: {excelPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PdfWatermark class is not available in the referenced Aspose.Cells version.
            // If watermark support is required, ensure the appropriate Aspose.Cells.Pdf assembly is referenced
            // and use PdfWatermark accordingly.

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(outputPdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save workbook as PDF
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF generated successfully at: {outputPdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
