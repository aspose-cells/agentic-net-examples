// Title: Add a PNG image watermark to a PDF with Aspose.Cells (C#) – full‑page scaling, 40% opacity
// Description: Demonstrates how to load a PNG file into a byte array, create a RenderingWatermark, set it to cover the entire page, apply 40% opacity, and embed it as a background image when saving an Aspose.Cells workbook to PDF.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark PNG byte array | scale watermark to page Aspose.Cells | watermark opacity 0.4 Aspose.Cells | background image PDF export .NET | add image watermark workbook PDF
// Common Searches: Aspose.Cells add PNG watermark to PDF | C# scale watermark to full page Aspose.Cells | set watermark opacity when exporting PDF with Aspose.Cells | use RenderingWatermark with byte array in .NET | apply background image to every PDF page Aspose.Cells
// Developer Intent: Insert a semi‑transparent PNG as a full‑page background watermark during PDF conversion of an Aspose.Cells workbook.
// Use Cases: Brand a financial report PDF with the company logo centered and faded across each page. | Mark confidential spreadsheets with a “CONFIDENTIAL” stamp image that covers the whole page at 40% opacity. | Retrieve a PNG stored in a database, convert it to a byte array, and automatically apply it as a watermark when generating PDFs from workbooks.
// AI Prompts: Generate C# code that reads a PNG from a byte array and applies it as a centered, page‑scaled watermark with 40% opacity using Aspose.Cells PdfSaveOptions. | Explain how to modify RenderingWatermark properties to change alignment, scaling percentage, and opacity for PDF output in Aspose.Cells. | Provide a robust pattern for handling a missing watermark file while still exporting the workbook to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to load a PNG file into a byte array, create a RenderingWatermark, set it to cover the entire page, apply 40% opacity, and embed it as a background image when saving an Aspose.Cells workbook to PDF.
class AddImageWatermark
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for watermark demo");

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Load watermark image if the file exists
            string watermarkPath = "watermark.png";
            if (File.Exists(watermarkPath))
            {
                byte[] pngBytes = File.ReadAllBytes(watermarkPath);
                RenderingWatermark watermark = new RenderingWatermark(pngBytes)
                {
                    ScaleToPagePercent = 100, // Scale to full page size
                    Opacity = 0.4f,            // 40% opacity
                    IsBackground = true,       // Place behind page contents
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center
                };
                pdfOptions.Watermark = watermark;
            }
            else
            {
                Console.WriteLine($"Warning: Watermark image '{watermarkPath}' not found. PDF will be saved without watermark.");
            }

            // Save the workbook as a PDF with (or without) the watermark
            string outputPath = "WatermarkedOutput.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
