// Title: Add a second WordArt shape with Simple Fill style and custom font size in Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, accesses the first worksheet's shape collection, adds a second WordArt using the Simple Fill preset (WordArtStyle1), sets its font size to 28 pt via TextEffectFormat, and saves the file as SecondWordArtDemo.xlsx.
// Keywords: Aspose.Cells WordArt | C# add WordArt shape | Simple Fill WordArt preset | custom WordArt font size | Excel shape collection Aspose | TextEffectFormat font size
// Common Searches: Aspose.Cells add second WordArt | WordArt Simple Fill style C# | change WordArt font size Aspose.Cells | multiple WordArt shapes Excel .NET | how to use TextEffectFormat in Aspose.Cells
// Developer Intent: Insert an additional WordArt with Simple Fill preset and adjust its font size.
// Use Cases: Add a subtitle WordArt beneath a title in automated report generation. | Create visually distinct section headers in dashboards. | Produce marketing spreadsheets with branded WordArt headings.
// AI Prompts: Generate C# code that adds three WordArt shapes with different presets and font sizes using Aspose.Cells. | Show how to change the fill color and outline of a WordArt after it is added in Aspose.Cells. | Write a reusable method that inserts a WordArt at a specified cell address, applies a chosen preset style, and sets a custom font size.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, accesses the first worksheet's shape collection, adds a second WordArt using the Simple Fill preset (WordArtStyle1), sets its font size to 28 pt via TextEffectFormat, and saves the file as SecondWordArtDemo.xlsx.
class AddSecondWordArt
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add the first WordArt (optional, shown for context)
        Shape firstWordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle2, // any preset style
            "First WordArt",
            1,   // top row
            0,   // top offset (pixels)
            1,   // left column
            0,   // left offset (pixels)
            100, // height (pixels)
            400  // width (pixels)
        );

        // Add the second WordArt with a built‑in style (e.g., Simple Fill) and custom font size
        Shape secondWordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // using a preset style as "Simple Fill"
            "Second WordArt",
            5,   // top row
            0,   // top offset (pixels)
            5,   // left column
            0,   // left offset (pixels)
            100, // height (pixels)
            400  // width (pixels)
        );

        // Set a custom font size for the second WordArt
        if (secondWordArt.IsWordArt)
        {
            TextEffectFormat textEffect = secondWordArt.TextEffect;
            textEffect.FontSize = 28; // custom font size in points
        }

        // Save the workbook
        workbook.Save("SecondWordArtDemo.xlsx");
    }
}
