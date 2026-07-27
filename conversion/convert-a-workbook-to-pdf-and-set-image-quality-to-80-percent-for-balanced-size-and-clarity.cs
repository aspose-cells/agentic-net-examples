// Title: C# – Convert Aspose.Cells Workbook to PDF with 80% JPEG Image Quality
// Description: Demonstrates how to save an Aspose.Cells workbook as a PDF while resampling images to 150 PPI and applying 80 % JPEG compression using PdfSaveOptions.SetImageResample.
// Keywords: Aspose.Cells PDF conversion C# | PdfSaveOptions SetImageResample | 80% JPEG quality PDF | image resampling 150 PPI | reduce PDF size Aspose | C# workbook to PDF example | Aspose.Cells image compression
// Common Searches: Aspose.Cells set JPEG quality when saving PDF | PdfSaveOptions image resample C# | convert Excel to PDF with reduced image size | Aspose.Cells PDF image compression 80 percent | how to control PDF file size Aspose.Cells
// Developer Intent: Export a workbook to PDF while controlling image resolution and JPEG compression for a balanced file size and visual clarity.
// Use Cases: Generate email‑friendly PDF reports from dynamically created spreadsheets. | Export large spreadsheets to PDF with predictable file size limits. | Batch‑process workbooks with consistent image quality across all PDFs.
// AI Prompts: Show how to change the image resolution to 200 PPI while keeping JPEG quality at 80 %. | Provide code that writes the PDF to a MemoryStream instead of a file using the same PdfSaveOptions. | Explain how to calculate the JPEG quality value needed to target a specific PDF size.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to save an Aspose.Cells workbook as a PDF while resampling images to 150 PPI and applying 80 % JPEG compression using PdfSaveOptions.SetImageResample.
public class WorkbookToPdfConverter
{
    public static void Run()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set image resampling: 150 PPI for reasonable resolution
            // and JPEG quality to 80% for balanced size and clarity
            pdfOptions.SetImageResample(150, 80);

            // Save the workbook as a PDF using the configured options
            workbook.Save("ConvertedWorkbook.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during PDF conversion: {ex.Message}");
        }
    }

    // Entry point required for console application
    public static void Main(string[] args)
    {
        Run();
    }
}
