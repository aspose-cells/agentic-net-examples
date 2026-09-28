// Title: Insert a Curved WordArt Text Effect into an Excel Worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells to add a WordArt shape with the Curved Text preset, specifying font, size, and cell‑based coordinates. | Show how to switch between different MsoPresetTextEffect values in Aspose.Cells to achieve alternative WordArt curvature levels.
// Common Searches: how to add curved WordArt text effect to Excel using Aspose.Cells C# | Aspose.Cells C# set WordArt shape position with cell indices | list of MsoPresetTextEffect values for WordArt in Aspose.Cells | change WordArt curvature in Aspose.Cells when Adjustments property is unavailable
// Tags: Aspose.Cells WordArt shape creation | C# apply Curved Text preset Aspose.Cells | Excel shape placement via cell indices Aspose.Cells | MsoPresetTextEffect enumeration Aspose.Cells | WordArt curvature control Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a WordArt shape using the built‑in Curved Text preset, defines its font, size, and cell‑based position and dimensions, and saves the file as WordArtCurved.xlsx. Direct curvature adjustment via an Adjustments property is not supported; the selected preset determines the curve.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the position of the WordArt shape (in cells)
            int upperLeftRow = 2;      // zero‑based row index
            int upperLeftColumn = 2;   // zero‑based column index
            int lowerRightRow = 6;
            int lowerRightColumn = 10;

            // Width and height of the shape (in points)
            int shapeWidth = 300;
            int shapeHeight = 100;

            // Add a WordArt shape with initial text
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1, // preset
                "Curved Text",                   // displayed text
                "Arial",                         // font name
                36,                              // font size
                false,                           // bold
                false,                           // italic
                upperLeftRow,
                upperLeftColumn,
                lowerRightRow,
                lowerRightColumn,
                shapeWidth,
                shapeHeight);

            // Note: Adjustments property is not available in Aspose.Cells Shape.
            // Curvature can be set via preset effects; the chosen preset already provides a curved style.

            // Save the workbook
            string outputPath = "WordArtCurved.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
