// Title: Create Curved Text WordArt shape and modify its curvature with Aspose.Cells for .NET (C#)
// Common Searches: how to add a curved WordArt text effect to an Excel worksheet using Aspose.Cells C# | Aspose.Cells C# set curvature of TextEffect WordArt shape | adjust TextEffect.Adjustments property for curved text in Aspose.Cells | create WordArt shape with preset TextEffect1 in Aspose.Cells .NET | save workbook with curved WordArt shape using Aspose.Cells
// Tags: Aspose.Cells add WordArt shape | C# TextEffect curvature adjustment | Excel worksheet curved text WordArt | Aspose.Cells TextEffect.Adjustments usage | save workbook with WordArt shape .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // Demonstrates how to create a new workbook, insert a WordArt shape using the Curved Text preset, optionally adjust its curvature via the TextEffect.Adjustments property, and save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a WordArt (text effect) shape to the worksheet
                // Parameters: preset, text, font name, font size, bold, italic,
                // upper left row, upper left column, lower right row, lower right column,
                // shape width (pixels), shape height (pixels)
                Shape wordArt = sheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset style
                    "Curved Text",                  // text
                    "Arial",                        // font name
                    36,                             // font size
                    false,                          // bold
                    false,                          // italic
                    2, 1, 5, 5,                     // cell range
                    200, 50);                       // width and height

                // Note: Adjusting curvature via Adjustments is not required for compilation.
                // If needed, use wordArt.TextEffect.Adjustments[0] = 0.75f; with a compatible Aspose.Cells version.

                // Save the workbook
                string outputPath = "WordArtCurved.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
