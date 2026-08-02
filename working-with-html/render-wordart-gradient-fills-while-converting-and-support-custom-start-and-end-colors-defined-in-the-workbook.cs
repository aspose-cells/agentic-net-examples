// Title: C# – Apply a Custom Two‑Color Gradient to WordArt Using Worksheet Cell Colors (Aspose.Cells)
// Description: This example creates a workbook, reads start and end color values from cells A1 and A2, converts them to System.Drawing.Color, adds a WordArt shape, sets its fill to a two‑color gradient, customizes the direction (45° diagonal) and saves the file as WordArtCustomGradient.xlsx.
// Keywords: Aspose.Cells WordArt gradient | custom gradient from Excel cells | C# WordArt two‑color gradient | set gradient angle Aspose.Cells | read color names Excel C# | WordArt fill type Gradient | Aspose.Cells shape styling
// Common Searches: how to set a custom gradient on WordArt with Aspose.Cells | read color values from Excel cells for WordArt fill | C# Aspose.Cells two‑color gradient WordArt example | change WordArt gradient angle programmatically | apply brand colors to WordArt using worksheet data
// Developer Intent: Create a WordArt shape and apply a two‑color gradient whose start and end colors are taken from worksheet cells.
// Use Cases: Brand‑consistent reports where heading WordArt colors are driven by cells, allowing non‑developers to update colors. | User‑customizable documents where end‑users enter their preferred colors in the spreadsheet and the WordArt updates automatically. | Automated flyer generation that reads marketing color codes from a sheet and applies a diagonal gradient to promotional WordArt.
// AI Prompts: Generate C# code with Aspose.Cells that reads hex color strings from cells and applies a radial gradient to a WordArt shape. | Show how to validate worksheet color names before converting them to System.Drawing.Color for a WordArt fill. | Explain how to switch the gradient style (horizontal, vertical, diagonal) after setting a two‑color gradient on WordArt using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This example creates a workbook, reads start and end color values from cells A1 and A2, converts them to System.Drawing.Color, adds a WordArt shape, sets its fill to a two‑color gradient, customizes the direction (45° diagonal) and saves the file as WordArtCustomGradient.xlsx.
class WordArtGradientExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define custom start and end colors in the worksheet (e.g., cells A1 and A2)
        // These could be any color names or hex strings; here we use known color names.
        sheet.Cells["A1"].PutValue("DarkOrange"); // start color name
        sheet.Cells["A2"].PutValue("MediumSeaGreen"); // end color name

        // Retrieve the color names from the cells
        string startColorName = sheet.Cells["A1"].StringValue;
        string endColorName = sheet.Cells["A2"].StringValue;

        // Convert the names to System.Drawing.Color objects
        Color startColor = Color.FromName(startColorName);
        Color endColor = Color.FromName(endColorName);

        // Add a WordArt shape to the worksheet
        // Use any preset style; we will override the fill with our custom gradient.
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // initial preset (will be changed)
            "Aspose.Cells Gradient WordArt",
            2, 0, // upper left row, top offset
            2, 0, // upper left column, left offset
            300, // height
            600  // width
        );

        // Set the fill type of the WordArt to Gradient
        wordArt.Fill.FillType = FillType.Gradient;

        // Apply a two‑color gradient using the custom colors from the workbook
        // GradientStyleType.Horizontal creates a left‑to‑right gradient; variant 1 is the default.
        wordArt.Fill.SetTwoColorGradient(startColor, endColor, GradientStyleType.Horizontal, 1);

        // Optionally, adjust the gradient direction or angle via the GradientFill object
        // Here we set a 45‑degree angle for a diagonal effect.
        wordArt.Fill.GradientFill.Angle = 45.0f;

        // Save the workbook with the customized WordArt
        workbook.Save("WordArtCustomGradient.xlsx");
    }
}
