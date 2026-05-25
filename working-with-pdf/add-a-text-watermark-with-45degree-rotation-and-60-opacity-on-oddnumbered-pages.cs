using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Populate some data to generate multiple pages when exported to PDF
                Worksheet sheet = workbook.Worksheets[0];
                for (int row = 0; row < 200; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                }

                // Create a rendering font for the watermark text
                RenderingFont font = new RenderingFont("Arial", 72)
                {
                    Bold = true,
                    Color = Color.Gray
                };

                // Create a text watermark with the specified font
                RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
                {
                    Rotation = 45f,          // 45‑degree rotation
                    Opacity = 0.6f,          // 60% opacity
                    IsBackground = true,    // placed behind page contents
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center
                };

                // Configure PDF save options and assign the watermark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                    // Note: Aspose.Cells applies the watermark to all pages.
                    // Per‑page control is not supported via PdfSaveOptions.
                };

                // Save the workbook as PDF with the configured watermark (lifecycle save)
                string outputPath = "Workbook_With_OddPage_Watermark.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}