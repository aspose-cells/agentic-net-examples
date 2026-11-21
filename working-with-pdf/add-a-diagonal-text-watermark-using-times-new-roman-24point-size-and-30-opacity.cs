// Title: Create a diagonal Times New Roman 24‑pt text watermark with 30% opacity in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add diagonal CONFIDENTIAL watermark to Excel workbook | How to set 30% opacity for a text watermark using Aspose.Cells .NET | Rotate a text effect shape -45 degrees in Aspose.Cells C# example | Center a watermark shape on an Excel worksheet with Aspose.Cells | Apply Times New Roman 24pt text watermark in Aspose.Cells workbook
// Tags: diagonal text watermark Aspose.Cells | text effect shape transparency C# | rotate watermark shape Aspose.Cells | center watermark on worksheet Aspose.Cells | Times New Roman 24pt watermark Excel

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, adds a TextEffect shape containing the word "CONFIDENTIAL" in Times New Roman 24‑pt, sets the fill transparency to 0.7 (30% opacity), hides the outline, rotates the shape -45 degrees, centers it using the worksheet's page dimensions, and saves the file as Watermarked.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a text effect shape that will serve as the watermark
            // Parameters: preset effect, text, font name, font size, bold, italic,
            // left, top, width, height, text effect alignment, text effect direction
            Shape watermarkShape = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "CONFIDENTIAL",
                "Times New Roman",
                24,
                false,
                false,
                0,
                0,
                500,
                200,
                0,   // text effect alignment (default)
                0);  // text effect direction (default)

            // Set shape properties: make background transparent, hide outline, rotate diagonally
            watermarkShape.Fill.Transparency = 0.7; // 70% transparent (approx. 30% opacity)

            // Hide the outline by making it fully transparent (Line.IsVisible is not available)
            watermarkShape.Line.Transparency = 1.0; // 100% transparent line

            watermarkShape.RotationAngle = -45; // degrees

            // Position the watermark at the center of the page using page dimensions
            double pageWidth = sheet.PageSetup.PaperWidth;
            double pageHeight = sheet.PageSetup.PaperHeight;
            watermarkShape.Left = (int)((pageWidth - watermarkShape.Width) / 2);
            watermarkShape.Top = (int)((pageHeight - watermarkShape.Height) / 2);

            // Save the workbook with the watermark applied
            workbook.Save("Watermarked.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
