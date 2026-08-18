// Title: Apply 30% Fill Transparency to a WordArt Shape in Excel using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a WordArt shape with a preset style, set its Fill.Transparency to 0.3 (30 % opacity), and save the result as WordArtTransparency.xlsx with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | Excel WordArt | shape transparency | fill opacity | Fill.Transparency | preset WordArt style | programmatic Excel graphics | Aspose.Cells shape example
// Common Searches: how to set WordArt transparency in Aspose.Cells | Aspose.Cells C# WordArt fill opacity 30 percent | Excel WordArt shape transparency code sample | set Fill.Transparency for WordArt using Aspose.Cells .NET | make WordArt semi‑transparent in an Excel workbook
// Developer Intent: Programmatically set a WordArt shape's fill transparency to 30 % in an Excel worksheet.
// Use Cases: Add a light watermark behind data by using a semi‑transparent WordArt title. | Create a decorative header where the WordArt blends with cell colors without hiding content. | Generate template files that include faint WordArt branding for consistent report styling.
// AI Prompts: Provide C# code to change a WordArt shape's Fill.Transparency to 0.5 with Aspose.Cells. | How can I modify the transparency of an existing WordArt object in a loaded workbook? | Explain the steps to retrieve a shape from the Shapes collection and adjust its fill opacity.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a WordArt shape with a preset style, set its Fill.Transparency to 0.3 (30 % opacity), and save the result as WordArtTransparency.xlsx with Aspose.Cells for C#.
class WordArtTransparencyDemo
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
            PresetWordArtStyle.WordArtStyle1,   // preset style
            "Subtle Background",                // text
            2, 0,                               // top row and vertical offset (pixels)
            2, 0,                               // left column and horizontal offset (pixels)
            200, 400);                          // height and width (pixels)

        // Set the fill transparency to 30% (0.3)
        wordArt.Fill.Transparency = 0.3;

        // Save the workbook
        workbook.Save("WordArtTransparency.xlsx");
    }
}
