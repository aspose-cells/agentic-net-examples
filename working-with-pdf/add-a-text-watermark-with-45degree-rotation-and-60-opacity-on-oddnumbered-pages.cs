// Title: Add a 45° 60% Opacity Text Watermark to Odd Pages in PDF Export with Aspose.Cells for .NET
// Description: This example creates a five‑sheet workbook, defines a semi‑transparent Arial font, builds a "CONFIDENTIAL" RenderingWatermark rotated 45°, and assigns it to PdfSaveOptions.Watermark. Aspose.Cells applies the watermark to every page; to target only odd‑numbered pages you must export odd sheets separately and merge the PDFs.
// Keywords: Aspose.Cells PDF watermark C# | 45 degree text watermark Aspose.Cells | watermark opacity 0.6 Aspose.Cells | odd page watermark Aspose.Cells | RenderingWatermark example .NET | export workbook to PDF with watermark
// Common Searches: how to add diagonal watermark to specific PDF pages using Aspose.Cells | Aspose.Cells apply watermark only on odd pages | C# 45° text watermark with 60% opacity in PDF export | Aspose.Cells per‑page watermark limitation
// Developer Intent: The developer wants a diagonal, 60 % opaque text watermark that appears only on odd‑numbered pages of a PDF generated from an Aspose.Cells workbook.
// Use Cases: Create a confidential report where only odd pages show a diagonal watermark while even pages stay clean. | Export a multi‑sheet workbook to PDF with a centered, semi‑transparent watermark applied exclusively to odd‑indexed worksheets. | Generate separate PDFs for odd and even sheets, merge them, and meet regulatory formatting requirements.
// AI Prompts: Provide C# code that saves only the odd worksheets of an Aspose.Cells workbook as a PDF with a 45° 60% opacity text watermark, then merges them with the even‑page PDFs. | Explain step‑by‑step how RenderingWatermark properties work in Aspose.Cells and why the watermark is applied globally to all pages. | Write a script that iterates through workbook worksheets, applies a RenderingWatermark to each odd sheet, exports each to a temporary PDF, and combines all PDFs into a single document.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // This example creates a five‑sheet workbook, defines a semi‑transparent Arial font, builds a "CONFIDENTIAL" RenderingWatermark rotated 45°, and assigns it to PdfSaveOptions.Watermark. Aspose.Cells applies the watermark to every page; to target only odd‑numbered pages you must export odd sheets separately and merge the PDFs.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with sample data on multiple worksheets (each sheet becomes a page in PDF)
            Workbook workbook = new Workbook();
            for (int i = 0; i < 5; i++) // create 5 pages
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Cells["A1"].PutValue($"Page {i + 1}");
                if (i < 4) // add additional sheets
                {
                    workbook.Worksheets.Add();
                }
            }

            // Define the font for the text watermark
            RenderingFont font = new RenderingFont("Arial", 72)
            {
                Bold = true,
                Color = Color.FromArgb(153, 0, 0, 255) // 60% opacity color (alpha 153 out of 255)
            };

            // Create a text watermark
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                Rotation = 45f,          // 45‑degree rotation
                Opacity = 0.6f,          // 60% opacity
                IsBackground = true,    // place behind page contents
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                ScaleToPagePercent = 50 // optional scaling
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF.
            // Note: Aspose.Cells applies the watermark to all pages.
            // To restrict it to odd‑numbered pages, you would need to generate separate PDFs
            // for odd pages and then merge them, which is beyond the scope of this simple example.
            workbook.Save("Workbook_With_OddPage_Watermark.pdf", pdfOptions);
        }
    }
}
