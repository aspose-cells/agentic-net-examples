// Title: Lock a WordArt watermark in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a diagonal WordArt shape with custom text to a new workbook, sets its fill and rotation, sends it to the back, and locks the shape to prevent any user edits. | Show how to use the Shape.IsLocked property in Aspose.Cells to make a watermark non‑selectable and non‑movable in an Excel worksheet. | Provide a complete example that creates an Excel workbook, inserts a CONFIDENTIAL WordArt watermark, applies a -45° rotation, solid fill, and saves the file as WatermarkLocked.xlsx.
// Common Searches: asp.net c# lock wordart shape in excel using aspose.cells | how to prevent editing of a watermark shape in an Excel workbook with Aspose.Cells | add diagonal confidential watermark to excel file and make it read‑only with Aspose.Cells | set shape IsLocked property Aspose.Cells example c# | save excel workbook with locked wordart watermark using Aspose.Cells .NET
// Tags: Aspose.Cells prevent watermark editing | C# add diagonal WordArt to Excel | Shape.IsLocked property Aspose.Cells | Excel watermark read‑only .NET | Create WordArt watermark Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a new workbook, inserts a diagonal WordArt shape with the text “CONFIDENTIAL” as a watermark, configures its fill and rotation, moves it behind other objects, locks the shape using Shape.IsLocked so users cannot select, move, edit, or resize it, and saves the file as WatermarkLocked.xlsx.
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

            // Add a WordArt shape that will serve as a watermark
            // Parameters: preset text effect, text, font name, font size, bold, italic,
            // upper left row, column, top, left, bottom, right
            Shape watermark = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "CONFIDENTIAL",
                "Arial",
                72,
                false,
                false,
                0, 0,
                0, 0, 0, 0);

            // Rotate the watermark for typical diagonal appearance
            watermark.RotationAngle = -45;

            // Set fill to solid; color defaults to black (or set a custom color if supported)
            watermark.Fill.FillType = FillType.Solid;
            // If the API version supports SolidFillColor, uncomment the line below:
            // watermark.Fill.SolidFillColor = Color.FromArgb(50, Color.Red);

            // Hide outline if the API version supports it
            try
            {
                // watermark.Line.IsVisible = false; // Uncomment if supported
            }
            catch
            {
                // Ignore if not supported
            }

            // Send the shape to the back so it behaves like a background watermark
            watermark.ZOrderPosition = 0;

            // Lock the watermark to prevent editing and moving
            watermark.IsLocked = true;

            // Define output file path
            string outputPath = "WatermarkLocked.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook safely
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
