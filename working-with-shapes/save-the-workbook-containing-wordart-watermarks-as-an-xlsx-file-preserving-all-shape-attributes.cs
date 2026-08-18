// Title: Add a WordArt watermark to an Excel workbook and save as XLSX while preserving shape attributes with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, insert a WordArt shape as a diagonal CONFIDENTIAL watermark, set rotation, transparency, hide borders, use FreeFloating placement, and save the file as XLSX, ensuring all shape properties are retained.
// Keywords: Aspose.Cells | C# Excel watermark | WordArt shape | preserve shape properties | save workbook XLSX | rotation transparency | FreeFloating placement | .NET | Excel shape attributes | programmatic watermark
// Common Searches: Aspose.Cells add WordArt watermark C# | save Excel workbook with rotated shape Aspose.Cells | preserve shape transparency when exporting to XLSX | set FreeFloating placement for Excel shape Aspose | C# code to create diagonal watermark in Excel
// Developer Intent: Create an XLSX file that contains a WordArt watermark with specific formatting (rotation, transparency, no border) and keep those formatting details intact after saving.
// Use Cases: Automated confidential reports with a diagonal CONFIDENTIAL overlay. | Branding templates that overlay semi‑transparent logo text on every sheet. | Invoice generation that adds a PAID watermark while maintaining shape rotation and transparency. | Regulatory documents requiring a hidden watermark that persists across Excel viewers.
// AI Prompts: Generate C# code using Aspose.Cells to add a WordArt watermark, rotate -45°, set 80% transparency, hide border, and save as XLSX. | Explain how the Placement property affects WordArt visibility in saved Excel files with Aspose.Cells. | Show how to apply the same WordArt watermark to all worksheets in a workbook while preserving shape attributes. | Provide steps to verify that shape properties are retained after opening the saved XLSX in Excel.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtWatermark
{
    // Demonstrates how to create a new Workbook, insert a WordArt shape as a diagonal CONFIDENTIAL watermark, set rotation, transparency, hide borders, use FreeFloating placement, and save the file as XLSX, ensuring all shape properties are retained.
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

                // Add sample data (optional, just to have some content)
                sheet.Cells["A1"].PutValue("Sample Data");
                sheet.Cells["B2"].PutValue(123);

                // Add a WordArt shape that will act as a watermark
                // Parameters: preset effect, text, font name, font size, bold, italic,
                // left, top, width, height, shape type, text effect
                Shape wordArt = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1,
                    "CONFIDENTIAL",
                    "Arial",
                    72,
                    false,
                    false,
                    0,
                    0,
                    500,
                    100,
                    0,   // shape type (default)
                    0);  // text effect (default)

                // Set shape properties to make it look like a watermark
                wordArt.RotationAngle = -45;                     // Rotate for diagonal appearance
                wordArt.Fill.Transparency = 0.8;                 // Make it semi‑transparent
                // Hide the shape border (line). Setting line weight to 0 effectively hides it.
                wordArt.Line.Weight = 0;
                wordArt.Placement = PlacementType.FreeFloating; // Ensure it stays on top

                // Save the workbook to an XLSX file
                workbook.Save("WordArtWatermark.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
