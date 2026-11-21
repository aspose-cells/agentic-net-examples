// Title: Apply a diagonal “Confidential” text watermark to an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add diagonal confidential watermark to printed Excel pages | How to rotate and set transparency for a text watermark shape in Aspose.Cells .NET | Center a text effect shape on an Excel worksheet using Aspose.Cells page dimensions | Create semi‑transparent watermark on an Excel workbook with Aspose.Cells API | Add custom text watermark to Excel file programmatically in C#
// Tags: Aspose.Cells text effect shape watermark | C# rotate shape Aspose.Cells | Aspose.Cells shape fill transparency | position watermark with PageSetup dimensions | diagonal confidential watermark Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WatermarkExample
{
    // The example creates a new Workbook, adds a TextEffect shape containing the word "Confidential", rotates it -45° for a diagonal layout, sets 50% transparency, centers it using the worksheet's PageSetup dimensions, and saves the workbook as Watermarked.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a text effect shape that will serve as the watermark
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // left, top, width, height, shape type, shape id
                Shape watermark = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "Confidential",
                    "Arial",
                    72,
                    false,
                    false,
                    0,
                    0,
                    500,
                    200,
                    0,   // shape type (default)
                    0);  // shape id (auto)

                // Rotate the shape to display diagonally
                watermark.RotationAngle = -45;

                // Make the watermark semi‑transparent (0 = opaque, 1 = fully transparent)
                watermark.Fill.Transparency = 0.5;

                // Position the watermark roughly at the center of the page
                PageSetup pageSetup = sheet.PageSetup;
                double pageWidth = pageSetup.PaperSize == PaperSizeType.PaperA4 ? 8.27 * 72 : 8.5 * 72;   // width in points
                double pageHeight = pageSetup.PaperSize == PaperSizeType.PaperA4 ? 11.69 * 72 : 11 * 72; // height in points

                watermark.Left = (int)((pageWidth - watermark.Width) / 2);
                watermark.Top = (int)((pageHeight - watermark.Height) / 2);

                // Save the workbook with the watermark applied
                string outputPath = "Watermarked.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
