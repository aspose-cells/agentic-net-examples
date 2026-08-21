// Title: Add WordArt with a custom font and 48‑pt size using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new workbook, insert a WordArt shape with AddTextEffect, change its FontName to a custom family (e.g., Comic Sans MS) and set FontSize to 48 points via TextEffectFormat, then save the file as CustomWordArt.xlsx.
// Keywords: Aspose.Cells | C# | .NET | WordArt shape | AddTextEffect | TextEffectFormat | custom font | font size 48 pt | Comic Sans MS | shape formatting | worksheet graphics | example code
// Common Searches: Aspose.Cells add WordArt C# | set custom font for WordArt Aspose.Cells | change WordArt font size .NET | TextEffectFormat font name example | how to use AddTextEffect in Aspose.Cells
// Developer Intent: Insert a WordArt shape, assign a specific font family, and set its size to 48 points in a worksheet.
// Use Cases: Create a decorative title banner with a brand‑specific font for reports. | Add eye‑catching labels to charts or tables using WordArt with precise sizing. | Automate workbook branding by applying the company’s font to WordArt across multiple files.
// AI Prompts: Generate C# code with Aspose.Cells to add a WordArt shape using the "Arial Black" font at 60 pt positioned at row 5, column 3. | Explain how to modify the TextEffectFormat of an existing WordArt shape to change its color, outline, and shadow in Aspose.Cells. | Provide a step‑by‑step tutorial for creating a WordArt shape, setting a custom font, and exporting the workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook, insert a WordArt shape with AddTextEffect, change its FontName to a custom family (e.g., Comic Sans MS) and set FontSize to 48 points via TextEffectFormat, then save the file as CustomWordArt.xlsx.
class WordArtExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a WordArt shape using AddTextEffect
        // Parameters: effect, text, fontName, size, bold, italic,
        // topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = shapes.AddTextEffect(
            MsoPresetTextEffect.TextEffect1,   // preset effect
            "Custom WordArt",                  // displayed text
            "Times New Roman",                 // initial font name (will be overridden below)
            12,                                // initial size (will be overridden below)
            false,                             // not bold
            false,                             // not italic
            2, 0,                              // top row and pixel offset
            2, 0,                              // left column and pixel offset
            200,                               // height in pixels
            400);                              // width in pixels

        // Retrieve the TextEffectFormat to customize font properties
        TextEffectFormat textEffect = wordArt.TextEffect;

        // Set the custom font family
        textEffect.FontName = "Comic Sans MS";

        // Set the font size to 48 points
        textEffect.FontSize = 48;

        // Save the workbook to a file
        workbook.Save("CustomWordArt.xlsx");
    }
}
