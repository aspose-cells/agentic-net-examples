// Title: C# – Convert Excel Workbook to PDF with Text Watermark on Odd Pages (Aspose.Cells)
// Description: Demonstrates how to create an Excel workbook, populate it with data, define a custom RenderingWatermark, and save the workbook as a PDF using Aspose.Cells for .NET. The example sets the Watermark property on PdfSaveOptions (applies to the whole document) and includes a placeholder IPageSavingCallback for future odd‑page watermark support, highlighting the current API limitation.
// Keywords: Aspose.Cells PDF conversion | C# Excel to PDF | text watermark Aspose.Cells | odd page watermark .NET | RenderingWatermark example | PdfSaveOptions watermark | Excel workbook PDF Aspose
// Common Searches: Aspose.Cells add watermark to PDF | C# convert Excel to PDF with watermark | apply watermark only on odd pages Aspose.Cells | PdfSaveOptions Watermark property usage | how to use IPageSavingCallback for watermarks
// Developer Intent: Generate a PDF from an Excel workbook and apply a text watermark, with the goal of targeting odd pages (currently limited to whole‑document watermark).
// Use Cases: Create and fill a Workbook programmatically before conversion. | Define a RenderingWatermark with custom font, color, rotation, opacity, and scaling. | Configure PdfSaveOptions to attach the watermark and save the workbook as PDF. | Implement a stub IPageSavingCallback for future per‑page watermark logic.
// AI Prompts: Show how to modify the code to place a watermark only on odd pages after the PDF is generated. | Explain why Aspose.Cells does not support per‑page watermarks via PageSavingArgs and suggest alternative approaches. | Provide a sample workflow that uses a secondary PDF library to add odd‑page watermarks to an Aspose.Cells‑generated PDF.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorkbookToPdfOddPageWatermark
{
    // Custom callback placeholder – retained for compatibility; no per‑page watermark applied
    // Demonstrates how to create an Excel workbook, populate it with data, define a custom RenderingWatermark, and save the workbook as a PDF using Aspose.Cells for .NET. The example sets the Watermark property on PdfSaveOptions (applies to the whole document) and includes a placeholder IPageSavingCallback for future odd‑page watermark support, highlighting the current API limitation.
    public class OddPageWatermarkCallback : IPageSavingCallback
    {
        private readonly RenderingWatermark _watermark;

        public OddPageWatermarkCallback(RenderingWatermark watermark)
        {
            _watermark = watermark;
        }

        // Called for each page during PDF saving
        public void PageSaving(PageSavingArgs args)
        {
            // The current Aspose.Cells version does not expose a Watermark property on PageSavingArgs.
            // Therefore, per‑page watermarking is not applied here.
            // If a future version adds this capability, the logic can be restored.
        }

        public void PageStartSaving(PageStartSavingArgs args) { }
        public void PageEndSaving(PageEndSavingArgs args) { }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and populate data
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // Populate sample data (enough rows to generate multiple pages)
                for (int row = 0; row < 200; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                    sheet.Cells[row, 1].PutValue($"Value {row + 1}");
                }

                // -------------------------------------------------
                // 2. Prepare the text watermark
                // -------------------------------------------------
                RenderingFont font = new RenderingFont("Calibri", 68)
                {
                    Italic = true,
                    Bold = true,
                    Color = Color.Blue
                };

                RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
                {
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    Rotation = 45,
                    Opacity = 0.3f,
                    ScaleToPagePercent = 75,
                    IsBackground = true
                };

                // -------------------------------------------------
                // 3. Configure PDF save options
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Apply watermark to all pages (per‑page control not available in this version)
                    Watermark = watermark,
                    PageSavingCallback = new OddPageWatermarkCallback(watermark)
                };

                // -------------------------------------------------
                // 4. Save the workbook as PDF
                // -------------------------------------------------
                string outputPath = "Workbook_OddPages_Watermarked.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
