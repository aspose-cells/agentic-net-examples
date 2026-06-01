using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class WordArtCustomCurvature
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape using a preset style
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,   // preset style
            "Custom Curved WordArt",            // text
            2, 0,                               // top row and vertical offset
            2, 0,                               // left column and horizontal offset
            100, 400);                          // height and width

        // Ensure the shape is a WordArt and access its TextEffect format
        if (wordArt.IsWordArt)
        {
            TextEffectFormat textEffect = wordArt.TextEffect;

            // Set a preset curvature shape (e.g., ArchUpCurve)
            textEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;

            // Optionally adjust font properties
            textEffect.FontName = "Arial";
            textEffect.FontSize = 24;
            textEffect.FontBold = true;
        }

        // Access the geometry of the shape to modify adjustment values
        // This can fine‑tune the curvature of the WordArt
        Geometry geometry = wordArt.Geometry;

        // Add a custom adjustment guide; the name depends on the shape type.
        // For demonstration, we add a generic guide named "Adj1" with a value of 0.5.
        geometry.ShapeAdjustValues.Add("Adj1", 0.5);

        // Save the workbook with the WordArt shape
        workbook.Save("WordArtCustomCurvature.xlsx");
    }
}