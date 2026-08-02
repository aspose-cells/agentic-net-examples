// Title: C# – Export Aspose.Cells Workbook to PDF with 300 DPI Images for Print‑Ready Quality
// Description: Demonstrates how to set the global DPI, configure PdfSaveOptions with SetImageResample(300, 100), and save a workbook as a PDF where all raster images are rendered at 300 DPI, delivering print‑ready output.
// Keywords: Aspose.Cells PDF export C# | 300 DPI image resample | PdfSaveOptions SetImageResample | CellsHelper DPI | high resolution PDF from Excel | .NET workbook to PDF | print quality Excel conversion
// Common Searches: Aspose.Cells export PDF 300 DPI | Set image DPI when saving Excel to PDF .NET | PdfSaveOptions high resolution images | How to increase PDF image quality with Aspose.Cells | global DPI setting Aspose.Cells C#
// Developer Intent: Create a PDF from an Excel workbook where every embedded image is rendered at 300 DPI for high‑quality printing.
// Use Cases: Generating marketing brochures with crisp graphics from Excel data. | Producing financial reports that require sharp charts for professional printing. | Batch converting multiple spreadsheets to print‑ready PDFs for archival compliance.
// AI Prompts: Show C# code that loads an existing .xlsx, sets CellsHelper.DPI to 300, applies PdfSaveOptions.SetImageResample(300, 100), and saves it as a PDF. | Give an example of naming the output PDF based on a cell value while keeping the 300 DPI image setting. | Explain how to adjust Aspose.Cells global DPI without affecting other rendering operations in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

namespace AsposeCellsPdfDpiDemo
{
    // Demonstrates how to set the global DPI, configure PdfSaveOptions with SetImageResample(300, 100), and save a workbook as a PDF where all raster images are rendered at 300 DPI, delivering print‑ready output.
    class Program
    {
        static void Main()
        {
            // Set the global DPI for the machine to 300.
            // This influences rendering of images when the workbook is saved.
            CellsHelper.DPI = 300;

            // Create a new workbook (or load an existing one).
            // Here we create a simple workbook with sample data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("High‑Resolution PDF Export");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(12345);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the desired image resample DPI to 300 and JPEG quality to 100.
            // This ensures that any raster images embedded in the PDF are rendered at 300 dpi.
            pdfOptions.SetImageResample(300, 100);

            // Optional: choose a high‑quality optimization type.
            pdfOptions.OptimizationType = PdfOptimizationType.Standard;

            // Save the workbook as a PDF with the specified options.
            string outputPath = "HighResolutionOutput.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved to PDF with 300 dpi images at: {outputPath}");
        }
    }
}
