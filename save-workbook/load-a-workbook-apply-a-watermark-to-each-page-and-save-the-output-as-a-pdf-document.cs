// Title: Apply a diagonal semi‑transparent CONFIDENTIAL text watermark to every page while converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, creates a RenderingWatermark with custom font, rotation, opacity, and centering, then saves the workbook as a PDF using PdfSaveOptions. | Show how to set PdfSaveOptions.Watermark to embed a background text watermark on all pages of the PDF generated from a workbook in Aspose.Cells for .NET.
// Common Searches: how to add a diagonal text watermark to PDF when converting Excel with Aspose.Cells C# | Aspose.Cells C# set watermark opacity and rotation in PdfSaveOptions | add a semi‑transparent diagonal watermark to each PDF page during Excel to PDF conversion with Aspose.Cells | C# code sample for rendering watermark on PDF output from Excel file using Aspose.Cells
// Tags: Aspose.Cells RenderingWatermark PDF conversion | C# add text watermark to PDF via PdfSaveOptions | Excel to PDF watermark Aspose.Cells | semi-transparent diagonal watermark Aspose.Cells | watermark behind page content C# Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // The example loads 'input.xlsx' into an Aspose.Cells Workbook, defines a blue, bold, italic Calibri font, creates a centered 45-degree semi‑transparent 'CONFIDENTIAL' RenderingWatermark, assigns it to PdfSaveOptions.Watermark, and saves the workbook as 'output_watermark.pdf' with the watermark applied to every PDF page.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Create a font for the watermark text
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create a text watermark with the desired text and font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for visual effect
                Rotation = 45,
                // Set opacity (0 = fully transparent, 1 = fully opaque)
                Opacity = 0.3f,
                // Scale watermark relative to page size (percentage)
                ScaleToPagePercent = 75,
                // Place the watermark behind the page content
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            string outputPath = "output_watermark.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved as PDF with watermark to: {outputPath}");
        }
    }
}
