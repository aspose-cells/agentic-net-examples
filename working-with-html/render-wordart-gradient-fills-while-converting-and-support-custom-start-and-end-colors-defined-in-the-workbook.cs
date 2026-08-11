// Title: Apply Custom Gradient Fill to WordArt Using Cell‑Defined Hex Colors – Aspose.Cells for .NET
// Description: Creates a workbook, stores start and end hex color strings in cells A1 and B1, adds a WordArt shape, sets its fill type to Gradient, converts the hex values to System.Drawing.Color, applies a two‑color horizontal gradient with a 45° linear direction, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells | WordArt gradient fill | custom hex colors | C# | .NET | Excel gradient | two‑color gradient | linear gradient angle | FillType.Gradient | SetTwoColorGradient | GradientFillType.Linear
// Common Searches: Aspose.Cells WordArt custom gradient from Excel cells | C# set start and end colors for WordArt gradient | read hex color value from worksheet and apply to shape | change WordArt gradient direction programmatically | apply two‑color gradient to WordArt using Aspose.Cells
// Developer Intent: Read hex color codes from worksheet cells and use them to define a custom two‑color gradient for a WordArt shape in an Excel file.
// Use Cases: Brand‑consistent reports where WordArt headings use company colors stored in the workbook. | User‑driven styling: end‑users enter palette values in cells and the generated Excel reflects those colors in WordArt. | Dynamic visual emphasis by adjusting gradient direction or angle based on data context.
// AI Prompts: Show how to extend the example to a three‑color gradient using cells C1, D1, and E1. | Generate code that reads RGB triples from a range and applies a radial gradient to a WordArt shape. | Explain how to toggle between linear and path gradient types for WordArt with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, stores start and end hex color strings in cells A1 and B1, adds a WordArt shape, sets its fill type to Gradient, converts the hex values to System.Drawing.Color, applies a two‑color horizontal gradient with a 45° linear direction, and saves the file as an .xlsx document.
class WordArtGradientExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define custom start and end colors in cells (as HTML hex strings)
        sheet.Cells["A1"].PutValue("#FF5733"); // start color
        sheet.Cells["B1"].PutValue("#33C1FF"); // end color

        // Add a WordArt shape (initially with any preset style)
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1, // preset (will be overridden by custom gradient)
            "Custom Gradient WordArt",
            2, 0,   // upper left row, top
            2, 0,   // upper left column, left
            300, 100); // height, width

        // Set the fill type to Gradient to enable gradient operations
        wordArt.Fill.FillType = FillType.Gradient;

        // Retrieve the custom colors from the worksheet
        string startHex = sheet.Cells["A1"].StringValue;
        string endHex = sheet.Cells["B1"].StringValue;

        // Convert HTML hex strings to System.Drawing.Color
        Color startColor = ColorTranslator.FromHtml(startHex);
        Color endColor = ColorTranslator.FromHtml(endHex);

        // Apply a two‑color gradient to the WordArt shape
        // Using Horizontal style and variant 1 (default)
        wordArt.Fill.SetTwoColorGradient(startColor, endColor, GradientStyleType.Horizontal, 1);

        // Optionally, adjust the gradient direction or angle via GradientFill
        // Here we set a linear gradient at 45 degrees
        wordArt.Fill.GradientFill.SetGradient(GradientFillType.Linear, 45, GradientDirectionType.FromUpperLeftCorner);

        // Save the workbook to demonstrate the result
        workbook.Save("WordArtCustomGradient.xlsx");
    }
}
