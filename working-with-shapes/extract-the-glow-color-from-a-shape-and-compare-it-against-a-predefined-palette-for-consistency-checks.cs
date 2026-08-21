// Title: Extract and Validate a Shape’s Glow Color with Aspose.Cells for .NET
// Description: Shows how to read a shape’s GlowEffect color, verify its existence in the workbook palette via Workbook.IsColorInPalette, compare it to a custom Color[] list, and obtain the closest workbook palette match using Workbook.GetMatchingColor.
// Keywords: Aspose.Cells shape glow color | Workbook.IsColorInPalette | Workbook.GetMatchingColor | C# extract glow effect | Excel palette validation | custom color palette Aspose.Cells | compare shape color .NET | color consistency check | Aspose.Cells GetMatchingColor example
// Common Searches: read glow color of a shape Aspose.Cells C# | check if shape glow color is in workbook palette | compare shape glow color to custom palette | find nearest Excel palette color for shape glow | Aspose.Cells IsColorInPalette usage
// Developer Intent: Retrieve a worksheet shape’s glow color and determine whether it matches a predefined set or the workbook’s internal palette.
// Use Cases: Enforce brand guidelines by confirming that shape glows use approved colors before publishing a workbook. | Detect and flag out‑of‑spec glow colors in automatically generated reports by comparing them to a whitelist. | Automatically replace non‑standard glow colors with the nearest workbook palette color to maintain compatibility with older Excel versions.
// AI Prompts: Create a C# method that takes a Shape and a Color[] and returns true if the shape's Glow.Color exists in the array, using Aspose.Cells. | Generate code that replaces a shape's glow color with the closest workbook palette color when it is not found in a custom palette, leveraging Workbook.GetMatchingColor. | Explain the difference between Workbook.IsColorInPalette and Workbook.GetMatchingColor when working with System.Drawing.Color objects in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGlowPaletteCheck
{
    // Shows how to read a shape’s GlowEffect color, verify its existence in the workbook palette via Workbook.IsColorInPalette, compare it to a custom Color[] list, and obtain the closest workbook palette match using Workbook.GetMatchingColor.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // Configure the glow effect for the shape
            GlowEffect glow = shape.Glow;
            glow.Size = 12; // radius in points
            glow.Transparency = 0.3; // 30% transparent
            glow.Color = workbook.CreateCellsColor(); // create a CellsColor instance
            glow.Color.Color = Color.FromArgb(255, 123, 200, 150); // custom glow color

            // Extract the glow color (System.Drawing.Color) from the shape
            Color extractedGlowColor = shape.Glow.Color.Color;

            // Define a predefined palette (example: standard Excel palette colors)
            Color[] predefinedPalette = new Color[]
            {
                Color.Black,
                Color.White,
                Color.Red,
                Color.Green,
                Color.Blue,
                Color.Yellow,
                Color.Magenta,
                Color.Cyan,
                Color.FromArgb(255, 123, 200, 150) // include the custom color for demonstration
            };

            // Check if the extracted glow color exists in the workbook's palette
            bool isInWorkbookPalette = workbook.IsColorInPalette(extractedGlowColor);
            Console.WriteLine($"Glow color {extractedGlowColor} in workbook palette: {isInWorkbookPalette}");

            // Perform consistency check against the predefined palette
            bool isInPredefinedPalette = false;
            foreach (Color paletteColor in predefinedPalette)
            {
                if (paletteColor.ToArgb() == extractedGlowColor.ToArgb())
                {
                    isInPredefinedPalette = true;
                    break;
                }
            }

            Console.WriteLine($"Glow color {extractedGlowColor} in predefined palette: {isInPredefinedPalette}");

            // If not in predefined palette, get the closest matching color from the workbook palette
            if (!isInPredefinedPalette)
            {
                Color matchedColor = workbook.GetMatchingColor(extractedGlowColor);
                Console.WriteLine($"Closest matching color in workbook palette: {matchedColor}");
            }

            // Save the workbook
            workbook.Save("GlowPaletteCheck.xlsx");
        }
    }
}
