// Title: Add a logo at (100,200) in a PDF using Aspose.Cells for .NET (C#)
// Description: Creates an in‑memory workbook, loads a PNG logo, configures a RenderingWatermark with OffsetX = 100 and OffsetY = 200 points, sets it as a foreground watermark without scaling, assigns it to PdfSaveOptions, and saves the result as a PDF.
// Keywords: Aspose.Cells | C# PDF watermark | RenderingWatermark OffsetX | RenderingWatermark OffsetY | PDFSaveOptions | place image in PDF | logo positioning | Aspose.Cells .NET example | add image to PDF | watermark coordinates
// Common Searches: Aspose.Cells place logo at specific coordinates | C# set RenderingWatermark offset X Y | How to add a foreground image watermark with Aspose.Cells | PDFSaveOptions watermark position example | Aspose.Cells add PNG logo to PDF
// Developer Intent: Insert a PNG logo at X = 100, Y = 200 points in the PDF generated from a workbook.
// Use Cases: Brand invoices with the company logo fixed at a precise spot on every page. | Generate corporate reports that embed a logo at an exact location in the PDF output. | Add a stamp or seal image to a specific coordinate when exporting spreadsheets to PDF.
// AI Prompts: Show how to scale the logo to 50 % while keeping the (100,200) position. | Provide code to add multiple RenderingWatermark objects at different coordinates in the same PDF. | Explain how to convert OffsetX/OffsetY from points to millimeters for the watermark.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLogoPdf
{
    // Creates an in‑memory workbook, loads a PNG logo, configures a RenderingWatermark with OffsetX = 100 and OffsetY = 200 points, sets it as a foreground watermark without scaling, assigns it to PdfSaveOptions, and saves the result as a PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Load logo image if the file exists
                string logoPath = "logo.png";
                if (File.Exists(logoPath))
                {
                    byte[] logoBytes = File.ReadAllBytes(logoPath);
                    RenderingWatermark watermark = new RenderingWatermark(logoBytes)
                    {
                        // Position the watermark at (100, 200) points
                        OffsetX = 100,
                        OffsetY = 200,
                        // Place the watermark in the foreground
                        IsBackground = false,
                        // Keep the original size of the image
                        ScaleToPagePercent = 100
                    };
                    pdfOptions.Watermark = watermark;
                }
                else
                {
                    Console.WriteLine($"Warning: Logo file '{logoPath}' not found. PDF will be generated without a watermark.");
                }

                // Save the workbook as PDF with the configured options
                workbook.Save("LogoPositioned.pdf", pdfOptions);
                Console.WriteLine("PDF generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
