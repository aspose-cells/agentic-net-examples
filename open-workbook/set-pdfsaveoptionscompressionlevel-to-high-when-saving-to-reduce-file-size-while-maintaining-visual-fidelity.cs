// Title: Aspose.Cells C# – Save Workbook as PDF with High (Flate) Compression
// Description: Creates a new workbook, adds sample text, configures PdfSaveOptions to use Flate compression and MinimumSize optimization, and saves the workbook as a PDF. The result is a smaller file that retains visual fidelity.
// Keywords: Aspose.Cells PDF compression C# | PdfSaveOptions Flate compression | MinimumSize PDF optimization Aspose.Cells | Reduce PDF file size .NET | High compression PDF Aspose.Cells | Save workbook as PDF Aspose.Cells | Aspose.Cells PDF export performance
// Common Searches: How to export an Aspose.Cells workbook to a small PDF in C# | Aspose.Cells set PdfSaveOptions compression level to high | Flate compression for PDF output using Aspose.Cells .NET | MinimumSize PDF optimization Aspose.Cells example | Reduce PDF size when saving workbook with Aspose.Cells | GitHub Aspose.Cells HighCompressionPdfDemo
// Developer Intent: Export a workbook to PDF with maximum compression to keep the file size minimal while preserving visual quality.
// Use Cases: Generate compact PDF invoices from Excel data for email attachment. | Archive large numbers of financial spreadsheets as low‑size PDFs for compliance. | Provide a web service that returns high‑compression PDFs generated on‑the‑fly. | Batch convert Excel reports to PDFs for mobile device distribution.
// AI Prompts: Show C# code using Aspose.Cells to save a workbook as a PDF with Flate compression and MinimumSize optimization. | Explain the impact of PdfCompression and PdfOptimizationType on PDF size and quality in Aspose.Cells. | Write a GitHub‑style README snippet for the HighCompressionPdfDemo example. | Suggest how to adjust PdfSaveOptions for different compression levels (Low, Medium, High) in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds sample text, configures PdfSaveOptions to use Flate compression and MinimumSize optimization, and saves the workbook as a PDF. The result is a smaller file that retains visual fidelity.
    public class HighCompressionPdfDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("High Compression PDF Demo");
                sheet.Cells["A2"].PutValue("This PDF is saved with high compression settings.");

                // Configure PDF save options for high compression
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    PdfCompression = PdfCompressionCore.Flate,          // Use Flate compression
                    OptimizationType = PdfOptimizationType.MinimumSize // Prioritize smaller file size
                };

                // Save the workbook as a PDF with the specified options
                string outputPath = "HighCompressionOutput.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during PDF generation: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            HighCompressionPdfDemo.Run();
        }
    }
}
