// Title: Add Curved WordArt (Arch Up Curve) to an Excel Sheet with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, insert a WordArt text‑effect shape, apply the ArchUpCurve preset, and save the result as CurvedWordArt.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | WordArt | Curved text | ArchUpCurve | TextEffectFormat | Excel shape | preset shape | programmatic WordArt | Excel automation
// Common Searches: Aspose.Cells add WordArt C# | Curved WordArt ArchUpCurve Aspose.Cells | Set TextEffectFormat preset shape in Excel via C# | Create text effect shape with Aspose.Cells | Adjust curvature of WordArt programmatically
// Developer Intent: Create a WordArt shape with a built‑in curved preset and store it in an Excel workbook.
// Use Cases: Design a decorative header for financial reports using ArchUpCurve WordArt. | Add a curved banner to a dashboard worksheet to highlight key performance indicators. | Automate the insertion of stylized titles into template workbooks for mass document generation.
// AI Prompts: Provide C# code that changes the curvature of an ArchUpCurve WordArt shape by modifying its TextEffectFormat.Adjustments collection in Aspose.Cells. | Show an example that loops through all MsoPresetTextEffectShape values and applies a selected shape to a WordArt object. | Explain how to read and update the font name, size, bold, and italic properties of a WordArt shape after setting its preset shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, insert a WordArt text‑effect shape, apply the ArchUpCurve preset, and save the result as CurvedWordArt.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt (text effect) shape
            Shape wordArt = shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,   // preset text effect
                "Curved Text",                     // text
                "Arial",                           // font name
                36,                                // font size
                false,                             // bold
                false,                             // italic
                2, 0,                              // top row, top offset
                2, 0,                              // left column, left offset
                200,                               // height (pixels)
                400);                              // width (pixels)

            // Apply a built‑in curved WordArt shape
            TextEffectFormat textEffect = wordArt.TextEffect;
            textEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;

            // Save the workbook
            workbook.Save("CurvedWordArt.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
