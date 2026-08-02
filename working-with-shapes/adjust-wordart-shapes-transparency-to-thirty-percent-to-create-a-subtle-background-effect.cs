// Title: Aspose.Cells for .NET – Set WordArt Fill Transparency to 30% in C#
// Description: This C# snippet creates a new workbook, inserts a WordArt object, sets its Fill.Transparency property to 0.3 (30 % opacity) for a subtle background effect, and saves the file as WordArtTransparency.xlsx.
// Keywords: Aspose.Cells C# WordArt fill opacity | Excel shape transparency .NET | adjust shape fill transparency Aspose.Cells | WordArt background effect example | C# Excel shape transparency tutorial
// Common Searches: C# Aspose.Cells how to change WordArt opacity | set fill transparency for a shape in Excel using .NET | adjust WordArt transparency programmatically Aspose.Cells | create faint WordArt watermark with 30% opacity in C#
// Developer Intent: I need to make a WordArt object semi‑transparent (30 %) so it can serve as a subtle background in an Excel worksheet.
// Use Cases: Add a low‑opacity WordArt watermark to generated financial reports. | Use semi‑transparent WordArt as a decorative header in Excel templates. | Provide faint visual cues on a sheet without covering data values.
// AI Prompts: Show me how to modify the WordArt fill transparency to 50% after the workbook is saved. | Provide code to read the current transparency of a WordArt shape and adjust it based on a condition. | Explain how to apply a gradient fill with varying opacity to a WordArt object using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# snippet creates a new workbook, inserts a WordArt object, sets its Fill.Transparency property to 0.3 (30 % opacity) for a subtle background effect, and saves the file as WordArtTransparency.xlsx.
class AdjustWordArtTransparency
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape to the worksheet
        // Parameters: style, text, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Background Text",
            2,    // top row index
            10,   // top offset in pixels
            2,    // left column index
            10,   // left offset in pixels
            100,  // height in pixels
            400   // width in pixels
        );

        // Set the fill transparency of the WordArt to 30% (0.3)
        wordArt.Fill.Transparency = 0.3;

        // Save the workbook
        workbook.Save("WordArtTransparency.xlsx");
    }
}
