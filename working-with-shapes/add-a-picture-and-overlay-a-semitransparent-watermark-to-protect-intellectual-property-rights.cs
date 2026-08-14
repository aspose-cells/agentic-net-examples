// Title: Add a Semi‑Transparent Picture and Diagonal Text Watermark to a PDF with Aspose.Cells (C#)
// Description: Demonstrates how to create a new workbook, insert an image into cell B2, set the image's transparency to 50%, define a red bold RenderingWatermark with 45° rotation and 20% opacity, and save the sheet as a PDF that contains both the semi‑transparent picture and the diagonal watermark.
// Keywords: Aspose.Cells | C# | add picture to worksheet | picture transparency | RenderingWatermark | PDF watermark | Excel to PDF | semi transparent watermark | background image Aspose.Cells | save workbook as PDF
// Common Searches: Aspose.Cells insert picture with transparency | C# create diagonal text watermark in PDF using Aspose.Cells | how to add semi‑transparent image and watermark when exporting Excel to PDF | RenderingWatermark example C# | save Excel workbook as PDF with background image Aspose.Cells
// Developer Intent: Insert a picture, make it semi‑transparent, and overlay a diagonal semi‑transparent text watermark while exporting the workbook to PDF using Aspose.Cells for .NET.
// Use Cases: Secure confidential reports by embedding a faint logo and a "CONFIDENTIAL" watermark in the PDF. | Brand marketing worksheets with a semi‑transparent product image and company watermark without obscuring data. | Produce printable training materials that include a background illustration and a light watermark to deter unauthorized copying.
// AI Prompts: Generate C# code with Aspose.Cells that places an image at cell C3, sets its transparency to 30%, and adds a "TOP SECRET" watermark rotated 30° with 15% opacity when saving as PDF. | Explain how to adjust RenderingWatermark properties (font, color, rotation, ScaleToPagePercent) for different page sizes in Aspose.Cells PDF export.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Demonstrates how to create a new workbook, insert an image into cell B2, set the image's transparency to 50%, define a red bold RenderingWatermark with 45° rotation and 20% opacity, and save the sheet as a PDF that contains both the semi‑transparent picture and the diagonal watermark.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the picture that will be inserted into the worksheet
                string picturePath = "sample_picture.jpg";

                // Add the picture if the file exists
                if (File.Exists(picturePath))
                {
                    // Add the picture to the worksheet at cell B2 (row index 1, column index 1)
                    int pictureIndex = sheet.Pictures.Add(1, 1, picturePath);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Set picture transparency to make it semi‑transparent (0.0 = opaque, 1.0 = fully transparent)
                    picture.FormatPicture.Transparency = 0.5; // 50% transparent
                }
                else
                {
                    Console.WriteLine($"Warning: Picture file '{picturePath}' not found. Skipping picture insertion.");
                }

                // -------------------------------------------------
                // Create a semi‑transparent text watermark
                // -------------------------------------------------
                // Define the font for the watermark text
                RenderingFont watermarkFont = new RenderingFont("Arial", 48)
                {
                    Bold = true,
                    Color = Color.Red
                };

                // Create the watermark with desired text and font
                RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
                {
                    // Place the watermark behind the worksheet content
                    IsBackground = true,
                    // Rotate the watermark for a diagonal appearance
                    Rotation = 45,
                    // Set opacity (0.0 – 1.0). 0.2 makes it faint but visible.
                    Opacity = 0.2f,
                    // Center the watermark on each page
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    // Scale watermark relative to page size (optional)
                    ScaleToPagePercent = 80
                };

                // Configure PDF save options to include the watermark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save the workbook as a PDF with the picture and watermark applied
                workbook.Save("Workbook_With_Picture_And_Watermark.pdf", pdfOptions);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
