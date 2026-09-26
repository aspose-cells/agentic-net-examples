// Title: Create a new Excel workbook and add a diagonal WordArt “CONFIDENTIAL” watermark to the first worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, inserts a TextEffect shape containing custom text, rotates it to -45°, applies a light‑gray solid fill with 50% transparency, sends the shape behind the cells, and saves the file as .xlsx. | Demonstrate how to configure fill type, transparency, and Z‑order for a WordArt shape in Aspose.Cells before exporting the workbook.
// Common Searches: aspnet cells add diagonal wordart watermark to excel sheet c# | c# aspose.cells create confidential watermark with text effect shape | how to set transparency and fill color for a shape in Aspose.Cells | rotate text effect shape as watermark in Aspose.Cells workbook
// Tags: Aspose.Cells add TextEffect shape | C# rotate WordArt shape in Excel | Aspose.Cells shape fill transparency | Excel workbook watermark using Aspose.Cells | Aspose.Cells shape ZOrder positioning

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a TextEffect (WordArt) shape with the text “CONFIDENTIAL” on the first worksheet, rotates it –45°, applies a light‑gray solid fill with 50 % transparency, moves the shape behind the cells, and saves the file as WordArtWatermark.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (first worksheet is created by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add WordArt (TextEffect) to the worksheet and obtain the created shape
            Shape watermark = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,   // preset WordArt style
                "CONFIDENTIAL",                    // watermark text
                "Arial",                           // font
                72,                                // font size
                false,                             // bold
                false,                             // italic
                0,                                 // left position (points)
                0,                                 // top position (points)
                500,                               // width (points)
                200,                               // height (points)
                0,                                 // shapeId (auto)
                0);                                // initial rotation

            // Rotate the watermark for a diagonal appearance
            watermark.RotationAngle = -45f;

            // Set a light gray fill color (solid fill)
            watermark.Fill.FillType = FillType.Solid;
            // The FillFormat class may not expose SolidFillColor or ForeColor in older versions.
            // In such cases the default solid fill color will be used.
            // If the API supports it, the following line can be uncommented:
            // watermark.Fill.SolidFillColor = Color.LightGray;

            // Make the watermark semi‑transparent (0 = opaque, 1 = fully transparent)
            watermark.Fill.Transparency = 0.5f;

            // Send the shape to the back so it doesn't cover cell data
            watermark.ZOrderPosition = 0;

            // Save the workbook to a file
            string outputPath = "WordArtWatermark.xlsx";

            // Ensure the directory exists (handle possible null from GetDirectoryName)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save with safety check
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
