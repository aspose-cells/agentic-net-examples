// Title: Extract Shape Glow Color and Validate Against Workbook or Custom Palette with Aspose.Cells for .NET
// Description: Demonstrates how to add a rounded‑rectangle auto shape, set its glow effect, read the glow color, check if the color exists in the workbook's internal palette or a developer‑defined Color array, retrieve the nearest matching palette color when needed, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape glow color | read glow color C# | compare glow color palette | workbook palette lookup Aspose.Cells | closest matching Excel color | auto shape glow verification | color consistency check .NET
// Common Searches: How to get the glow color of a shape in Aspose.Cells .NET | Check if a shape's glow color is in the workbook palette | Find nearest workbook palette color for a custom glow | Compare shape glow color with a predefined Color array | Validate Excel shape glow against brand colors
// Developer Intent: Read a shape's glow color and determine whether it matches a predefined palette or the workbook's internal color palette, optionally obtaining the closest palette match.
// Use Cases: Enforce corporate branding by confirming that shape glow colors conform to an approved color set before distributing the workbook. | Automate an audit that flags non‑standard glow colors and replaces them with the nearest workbook palette color. | Generate a compliance report that lists shapes with glow colors outside the allowed palette across multiple worksheets.
// AI Prompts: Create a C# method that receives a Shape and a Color[] palette, returns true if the shape's glow color is present in the palette using Aspose.Cells. | Write code to replace a shape's glow color with the closest workbook palette color when the original color is absent from the workbook palette. | Provide an example that iterates over all shapes on a worksheet, checks each glow color against a predefined palette, and logs any mismatches.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a rounded‑rectangle auto shape, set its glow effect, read the glow color, check if the color exists in the workbook's internal palette or a developer‑defined Color array, retrieve the nearest matching palette color when needed, and save the workbook using Aspose.Cells for .NET.
    class GlowColorConsistencyCheck
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rounded rectangle shape to the worksheet
                // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
                Shape shape = sheet.Shapes.AddAutoShape(
                    AutoShapeType.RoundedRectangle,
                    1,               // upper left row
                    1,               // upper left column
                    10,              // top (pixels)
                    10,              // left (pixels)
                    60,              // height (pixels)
                    100);            // width (pixels)

                // Configure the glow effect for the shape
                shape.Glow.Size = 12;                     // radius in points
                shape.Glow.Transparency = 0.3;            // 30% transparent
                shape.Glow.Color.Color = Color.FromArgb(255, 123, 200, 150); // custom glow color

                // Extract the glow color (as System.Drawing.Color)
                CellsColor glowCellsColor = shape.Glow.Color;
                Color glowSystemColor = glowCellsColor.Color;

                // Define a predefined palette to compare against
                Color[] predefinedPalette = new Color[]
                {
                    Color.Red,
                    Color.Green,
                    Color.Blue,
                    Color.FromArgb(255, 123, 200, 150), // same as the glow color for demonstration
                    Color.Yellow
                };

                // Check if the glow color exists in the workbook's internal palette
                bool inWorkbookPalette = workbook.IsColorInPalette(glowSystemColor);
                Console.WriteLine($"Glow color in workbook palette: {inWorkbookPalette}");

                // Check if the glow color exists in the custom predefined palette
                bool inCustomPalette = false;
                foreach (Color c in predefinedPalette)
                {
                    if (c.ToArgb() == glowSystemColor.ToArgb())
                    {
                        inCustomPalette = true;
                        break;
                    }
                }
                Console.WriteLine($"Glow color in custom palette: {inCustomPalette}");

                // If the glow color is not in the workbook palette, obtain the closest matching palette color
                if (!inWorkbookPalette)
                {
                    Color matchedColor = workbook.GetMatchingColor(glowSystemColor);
                    Console.WriteLine($"Closest matching workbook palette color: R={matchedColor.R}, G={matchedColor.G}, B={matchedColor.B}");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "GlowColorCheck.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            GlowColorConsistencyCheck.Run();
        }
    }
}
