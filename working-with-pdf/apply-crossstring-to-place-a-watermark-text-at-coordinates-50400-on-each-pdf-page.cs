// Title: Add a diagonal cross‑string watermark at coordinates (50,400) to each page when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add diagonal text watermark to PDF pages at specific coordinates | How to place a cross‑string watermark at (50,400) when saving Excel as PDF using Aspose.Cells | C# code to add a semi‑transparent rotated watermark to every worksheet before PDF export | Set watermark position and rotation angle in Aspose.Cells during Excel to PDF conversion | Apply WordArt watermark across all pages in a PDF generated from Excel with Aspose.Cells
// Tags: Aspose.Cells add diagonal watermark PDF | C# set watermark coordinates Aspose.Cells | WordArt shape watermark Excel to PDF | rotate shape watermark Aspose.Cells | semi transparent watermark Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, creates a semi‑transparent diagonal WordArt shape with the text "CONFIDENTIAL" positioned at (50,400) on each worksheet, rotates it to form a cross‑string, and saves the workbook as a PDF so the watermark appears on every PDF page.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Watermark settings
            string watermarkText = "CONFIDENTIAL";
            int posX = 50;   // horizontal offset in points
            int posY = 400;  // vertical offset in points

            // Add the watermark to every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a WordArt shape (row/column bounds and size are placeholders)
                Shape watermarkShape = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    watermarkText,                  // text
                    "Arial",                        // font name
                    36,                             // font size
                    false,                          // bold
                    false,                          // italic
                    0, 0, 0, 0,                     // upper-left and lower-right cell positions
                    0, 0);                          // width and height (auto‑size)

                // Position the shape
                watermarkShape.Left = posX;
                watermarkShape.Top = posY;

                // Make the shape semi‑transparent
                watermarkShape.Fill.Transparency = 0.8;   // 80% transparent fill
                watermarkShape.Line.Transparency = 0.8;   // 80% transparent border

                // Rotate the watermark to appear as a cross‑string
                watermarkShape.RotationAngle = -45; // diagonal orientation
            }

            // Save the workbook as PDF; the watermark will appear on each PDF page
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
