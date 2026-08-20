// Title: Aspose.Cells for .NET – Apply a Blue‑to‑Transparent Gradient Fill to WordArt (C#)
// Description: This example creates a workbook, inserts a WordArt shape, switches its Fill.FillType to Gradient, configures a GradientFill with SetTwoColorGradient so the color changes from solid blue to fully transparent, and saves the result as an .xlsx file.
// Keywords: Aspose.Cells | C# | WordArt gradient | blue transparent fill | GradientFill | SetTwoColorGradient | horizontal gradient
// Common Searches: Aspose.Cells set gradient fill for WordArt C# | blue to transparent WordArt example Aspose.Cells | how to create fading WordArt in .NET spreadsheet
// Developer Intent: Insert a WordArt object and give it a horizontal gradient that fades from opaque blue to transparent.
// Use Cases: Design a report header where the title text gradually blends into the worksheet background. | Create a visual separator in dashboards that uses a subtle blue fade to draw attention. | Generate marketing spreadsheets with gradient‑filled WordArt to highlight key sections.
// AI Prompts: Show how to change the gradient direction to vertical for the same WordArt shape. | Provide code for a three‑color gradient (blue, white, transparent) on WordArt using Aspose.Cells. | Explain how to modify the gradient variant index to produce a diagonal fade on WordArt.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// This example creates a workbook, inserts a WordArt shape, switches its Fill.FillType to Gradient, configures a GradientFill with SetTwoColorGradient so the color changes from solid blue to fully transparent, and saves the result as an .xlsx file.
class ConfigureWordArtGradient
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape (any preset style works)
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset style
            "Gradient WordArt",                // text
            1, 0,                             // upper left row, top
            1, 0,                             // upper left column, left
            300, 100);                        // height, width

        // Set the fill type to gradient to access gradient fill properties
        wordArt.Fill.FillType = FillType.Gradient;

        // Obtain the GradientFill object from the shape's fill
        GradientFill gradientFill = wordArt.Fill.GradientFill;

        // Configure a two‑color gradient: opaque blue to fully transparent blue
        gradientFill.SetTwoColorGradient(
            Color.Blue,   // first color (opaque)
            0.0,          // transparency for first color (0 = opaque)
            Color.Blue,   // second color (will be transparent)
            1.0,          // transparency for second color (1 = fully transparent)
            GradientStyleType.Horizontal, // gradient direction
            1);           // variant

        // Save the workbook with the configured WordArt
        workbook.Save("WordArtGradient.xlsx");
    }
}
