// Title: Add a Semi‑Transparent Diagonal Watermark While Converting Excel to PDF with Aspose.Cells for .NET (C#)
// Description: This C# example loads an .xlsx file, creates a RenderingWatermark with Calibri 68 pt bold italic blue text, centers it, rotates 45°, sets 30 % opacity, scales to 75 % of the page, places it behind the content, assigns it to PdfSaveOptions, and saves the workbook as a PDF that shows the watermark on every page.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | PDF watermark | RenderingWatermark | PdfSaveOptions | semi-transparent watermark | diagonal watermark | Aspose.Cells .NET example
// Common Searches: Add diagonal semi‑transparent watermark to PDF generated from Excel using Aspose.Cells | Aspose.Cells C# set watermark opacity when saving workbook as PDF | RenderingWatermark with PdfSaveOptions example | How to place watermark behind content in PDF conversion with Aspose.Cells | Convert Excel workbook to PDF with text watermark in .NET
// Developer Intent: Generate a PDF from an Excel workbook and overlay a semi‑transparent diagonal text watermark on every page using Aspose.Cells for .NET.
// Use Cases: Confidential reports that require a 'CONFIDENTIAL' watermark on the PDF version of an Excel file | Brand‑consistent marketing PDFs with a company logo or slogan watermarked across all pages | Automated batch processing of workbooks to PDF with regulatory‑required watermarks
// AI Prompts: Write C# code to add an image watermark with 50% opacity to a PDF saved from an Aspose.Cells workbook | Show how to change watermark rotation, scaling, and opacity for different page sizes in Aspose.Cells | Explain how to apply distinct watermarks to individual worksheets before merging them into a single PDF | Provide a PowerShell script that calls a .NET assembly to batch‑convert Excel files to PDF with a predefined watermark

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // This C# example loads an .xlsx file, creates a RenderingWatermark with Calibri 68 pt bold italic blue text, centers it, rotates 45°, sets 30 % opacity, scales to 75 % of the page, places it behind the content, assigns it to PdfSaveOptions, and saves the workbook as a PDF that shows the watermark on every page.
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create a font for the watermark text
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create a text watermark with the desired appearance
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for a diagonal effect
                Rotation = 45f,
                // Set semi‑transparent opacity (0 = fully transparent, 1 = fully opaque)
                Opacity = 0.3f,
                // Scale relative to the page size
                ScaleToPagePercent = 75,
                // Place the watermark behind the page content
                IsBackground = true
            };

            // Configure PDF save options to include the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            workbook.Save("output_watermarked.pdf", pdfOptions);

            Console.WriteLine("Workbook has been saved to PDF with a semi‑transparent watermark.");
        }
    }
}
