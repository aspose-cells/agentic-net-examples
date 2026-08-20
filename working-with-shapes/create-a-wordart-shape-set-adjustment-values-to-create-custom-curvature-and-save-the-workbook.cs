// Title: Add Curved WordArt to an Excel Sheet with Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a WordArt shape, changes its preset to ArchUpCurve, adjusts the curvature via the shape's adjustment guide, and saves the file as an .xlsx document using Aspose.Cells for C#.
// Keywords: Aspose.Cells WordArt | C# curved WordArt | ArchUpCurve shape | shape adjustment guide | custom WordArt curvature | add WordArt Excel .NET | save workbook with WordArt
// Common Searches: how to add curved WordArt with Aspose.Cells | adjust WordArt curvature C# Aspose | set ArchUpCurve preset shape Aspose.Cells | modify shape adjustment values in Excel using .NET | save Excel file after inserting WordArt
// Developer Intent: Programmatically insert a WordArt shape, apply a curved preset, tweak its curvature, and write the workbook to disk.
// Use Cases: Design eye‑catching titles for financial reports with custom‑curved WordArt. | Enhance Excel dashboards by adding stylized, curved headings for better visual hierarchy. | Automate branding by embedding curved WordArt labels into generated spreadsheets.
// AI Prompts: Generate code to change multiple adjustment guides of a WordArt shape in Aspose.Cells. | Show how to set WordArt curvature based on a percentage variable in C#. | List all PresetWordArtStyle options and demonstrate selecting one at runtime.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a WordArt shape, changes its preset to ArchUpCurve, adjusts the curvature via the shape's adjustment guide, and saves the file as an .xlsx document using Aspose.Cells for C#.
class WordArtCurvatureExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape with a preset style
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Custom Curved WordArt",
            2,      // topRow
            0,      // top offset (pixels)
            2,      // leftColumn
            0,      // left offset (pixels)
            100,    // height (pixels)
            400);   // width (pixels)

        // Set the preset shape to a curved type (e.g., ArchUpCurve)
        wordArt.TextEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;

        // Adjust the curvature by modifying the shape's adjustment guide values
        // The first adjustment guide typically controls the curvature amount
        if (wordArt.Geometry.ShapeAdjustValues.Count > 0)
        {
            // Set a custom curvature value (range 0.0 to 1.0)
            wordArt.Geometry.ShapeAdjustValues[0].Value = 0.6;
        }
        else
        {
            // If no guides exist, add a new one named "Adj1"
            wordArt.Geometry.ShapeAdjustValues.Add("Adj1", 0.6);
        }

        // Save the workbook
        workbook.Save("WordArtCurvatureExample.xlsx");
    }
}
