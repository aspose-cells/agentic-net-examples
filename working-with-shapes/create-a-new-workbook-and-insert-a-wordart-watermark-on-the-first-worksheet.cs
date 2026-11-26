// Title: Create a new Excel workbook and add a diagonal semi‑transparent WordArt watermark to the first worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add WordArt shape as watermark to Excel worksheet | how to set rotation and transparency for WordArt in Aspose.Cells | place WordArt behind cells in an Excel file using Aspose.Cells .NET | generate CONFIDENTIAL watermark in .xlsx using Aspose.Cells API
// Tags: add WordArt watermark Aspose.Cells .NET | set shape rotation transparency Aspose.Cells | freefloating placement shape behind cells | create diagonal watermark Excel using C# | presetwordartstyle usage Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, inserts a WordArt shape with the text "CONFIDENTIAL" as a watermark, rotates it -45°, sets 80% transparency, applies a light‑gray bold font, positions the shape behind the cells using FreeFloating placement, and saves the file as WordArtWatermark.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape that will act as a watermark
            // Parameters: style, text, topRow, top (pixel), leftColumn, left (pixel), height (pixel), width (pixel)
            Shape watermark = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,
                "CONFIDENTIAL",
                0,          // top row
                0,          // top offset in pixels
                0,          // left column
                0,          // left offset in pixels
                100,        // height in pixels
                600);       // width in pixels

            // Configure the appearance of the watermark
            watermark.RotationAngle = -45;                     // Diagonal orientation
            watermark.Fill.Transparency = 0.8;                // Make it semi‑transparent
            watermark.HasLine = false;                        // No border line
            watermark.Font.Color = Color.LightGray;
            watermark.Font.Size = 48;
            watermark.Font.IsBold = true;                     // Bold text

            // Send the shape behind the cells so it behaves like a true watermark
            // PlacementType.FreeFloating places the shape on the drawing layer (behind cells)
            watermark.Placement = PlacementType.FreeFloating;

            // Save the workbook
            string outputPath = "WordArtWatermark.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
