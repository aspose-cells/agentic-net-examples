// Title: Add WordArt with Custom Curvature in Aspose.Cells for .NET and Save Workbook
// Description: Creates a new Workbook, inserts a WordArt shape using a preset style, changes its TextEffect to ArchUpCurve, fine‑tunes the curve via Geometry.ShapeAdjustValues, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells WordArt | custom curvature | ShapeAdjustValues | Geometry adjustment | C# Excel shape | save workbook with WordArt | ArchUpCurve preset
// Common Searches: Aspose.Cells add WordArt with curve | how to adjust WordArt curvature in .NET | set ShapeAdjustValues for WordArt Aspose | save Excel file containing custom WordArt
// Developer Intent: Insert a WordArt shape, modify its curvature, and persist the workbook.
// Use Cases: Design a report header where the title appears as a curved WordArt banner. | Build a marketing template that automatically adds custom‑curved WordArt to each worksheet. | Programmatically apply different curvature levels to WordArt across multiple sheets for dynamic branding.
// AI Prompts: Generate C# code with Aspose.Cells that adds a WordArt shape using the ArchDownCurve preset, sets its curvature adjust value to 0.3, and saves the workbook. | Explain how Geometry.ShapeAdjustValues can be leveraged to control WordArt curvature in Aspose.Cells for .NET. | Show an example that creates several WordArt objects with varying curvature factors in a single Excel file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, inserts a WordArt shape using a preset style, changes its TextEffect to ArchUpCurve, fine‑tunes the curve via Geometry.ShapeAdjustValues, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape with a preset style
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Custom Curve",
            2,   // topRow
            0,   // top offset (pixels)
            2,   // leftColumn
            0,   // left offset (pixels)
            100, // height (pixels)
            400  // width (pixels)
        );

        // Verify the shape is WordArt before applying text‑effect properties
        if (wordArt.IsWordArt)
        {
            // Set a preset shape that provides a basic curvature
            wordArt.TextEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;

            // Access the geometry of the shape to fine‑tune curvature via adjust values
            Geometry geometry = wordArt.Geometry;

            // If the shape already has adjust guides, modify the first one;
            // otherwise, add a new guide named "Adj1" with a custom value.
            if (geometry.ShapeAdjustValues.Count > 0)
            {
                geometry.ShapeAdjustValues[0].Value = 0.5; // custom curvature factor
            }
            else
            {
                geometry.ShapeAdjustValues.Add("Adj1", 0.5);
            }
        }

        // Save the workbook with the customized WordArt
        workbook.Save("WordArtCustomCurvature.xlsx");
    }
}
