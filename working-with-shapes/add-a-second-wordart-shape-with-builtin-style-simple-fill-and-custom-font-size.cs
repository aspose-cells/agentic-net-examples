// Title: Create a second WordArt shape with Simple Fill style and custom font size using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a second WordArt shape with solid fill and custom font size | How to apply PresetWordArtStyle.WordArtStyle1 for simple fill in Aspose.Cells | Programmatically change WordArt fill color and font size with Aspose.Cells .NET | Add multiple WordArt objects to an Excel worksheet using the Aspose.Cells API | Set custom font size for WordArt shapes in a C# Aspose.Cells example
// Tags: add-wordart-shape-aspocells-csharp | presetwordartstyle-simple-fill | set-wordart-fontsize-aspocells | modify-wordart-fillcolor-aspocells | multiple-wordart-shapes-excel-aspocells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtDemo
{
    // The sample creates a new workbook, adds a first WordArt shape with a preset style and font size 18, then adds a second WordArt shape using the WordArtStyle1 preset (simple solid fill), sets its font size to 28, changes its fill color to DarkGreen, and saves the file as WordArtMultipleShapes.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // First WordArt shape (any preset style, just for demonstration)
            // -----------------------------------------------------------------
            // Parameters: style, text, topRow, top (pixel), leftColumn, left (pixel), height (pixel), width (pixel)
            Shape firstWordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle2,   // Blue fill with shadow
                "First WordArt",
                2,    // topRow
                0,    // top offset in pixels
                2,    // leftColumn
                0,    // left offset in pixels
                100,  // height
                400   // width
            );

            // Optionally set a custom font size for the first shape
            firstWordArt.TextEffect.FontSize = 18;

            // -----------------------------------------------------------------
            // Second WordArt shape with built‑in style "Simple Fill" and custom font size
            // -----------------------------------------------------------------
            // Aspose.Cells does not expose a style named exactly "Simple Fill". 
            // The closest built‑in simple fill style is WordArtStyle1 (black fill, no gradient).
            // You can change the style to any other preset if needed.
            Shape secondWordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,   // Simple fill (solid black)
                "Second WordArt",
                6,    // topRow (placed lower than the first shape)
                0,    // top offset in pixels
                2,    // leftColumn
                0,    // left offset in pixels
                100,  // height
                400   // width
            );

            // Set a custom font size (e.g., 28 points)
            secondWordArt.TextEffect.FontSize = 28;

            // Optionally, change the fill color to demonstrate the "Simple Fill" effect
            // The FillFormat of a WordArt shape controls the background of the text.
            secondWordArt.FillFormat.ForeColor = System.Drawing.Color.DarkGreen;

            // Save the workbook to a file
            string outputPath = "WordArtMultipleShapes.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("First WordArt IsWordArt: " + firstWordArt.IsWordArt);
            Console.WriteLine("Second WordArt IsWordArt: " + secondWordArt.IsWordArt);
        }
    }
}
