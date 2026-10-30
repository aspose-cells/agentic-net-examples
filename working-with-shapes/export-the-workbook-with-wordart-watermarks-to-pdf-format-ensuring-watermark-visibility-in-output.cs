// Title: Create a diagonal semi‑transparent WordArt watermark in Excel and export the workbook to PDF using Aspose.Cells for .NET
// Common Searches: how to add a diagonal WordArt watermark to an Excel worksheet and save as PDF with Aspose.Cells C# | Aspose.Cells C# export PDF with transparent WordArt shape watermark | save Excel workbook to PDF while preserving rotated shape watermark using Aspose.Cells | C# code to create semi‑transparent WordArt watermark in Excel and generate PDF | Aspose.Cells add text effect watermark and convert workbook to PDF
// Tags: addtexteffect shape watermark Aspose.Cells | rotate diagonal watermark Excel | transparent shape export PDF Aspose.Cells | excel workbook pdf conversion watermark | wordart watermark generation C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a WordArt shape with the text "CONFIDENTIAL" as a semi‑transparent, diagonal watermark across the first worksheet, and then saves the workbook as a PDF, ensuring the watermark appears in the exported document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add WordArt watermark to the worksheet
            // Parameters: preset effect, text, font name, font size, isBold, isItalic,
            // upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, height, width
            Shape watermark = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset effect
                "CONFIDENTIAL",                  // watermark text
                "Arial",                         // font name
                48,                              // font size
                false,                           // isBold
                false,                           // isItalic
                0, 0,                            // upper‑left cell
                30, 10,                          // lower‑right cell (covers the sheet)
                200, 400);                       // height and width of the shape

            // Make the watermark semi‑transparent
            watermark.Fill.Transparency = 0.8;   // 80% transparent
            watermark.Line.Transparency = 1.0;   // hide border

            // Rotate the watermark for diagonal appearance
            watermark.RotationAngle = -45;

            // Save the workbook as PDF ensuring the watermark is rendered
            workbook.Save("WorkbookWithWatermark.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
