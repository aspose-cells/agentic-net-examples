// Title: Set WordArt Gradient Angle and Export to HTML with CSS Linear‑Gradient – Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a WordArt shape, applies a linear gradient with a custom angle (e.g., 45°), defines two colors, and saves the sheet as HTML. The generated CSS includes the exact gradient angle using the linear‑gradient syntax.
// Keywords: Aspose.Cells | WordArt gradient angle | C# HTML export | linear-gradient CSS | GradientFill.SetGradient | .NET workbook to HTML | custom gradient direction | PresetWordArtStyle | ExportActiveWorksheetOnly | GradientFill.Angle
// Common Searches: how to set wordart gradient angle in aspose.cells | asp.net export wordart with css gradient | preserve wordart gradient direction when saving to html | c# linear‑gradient angle aspose.cells html | customize wordart fill gradient aspose.cells
// Developer Intent: Apply a specific linear gradient angle to a WordArt shape and have that angle reflected accurately in the HTML/CSS output.
// Use Cases: Programmatically create WordArt, assign a 45° linear gradient, and export the worksheet to HTML for web display. | Update the gradient angle of an existing WordArt shape before re‑saving the workbook as HTML. | Generate HTML reports where WordArt styling must match exact CSS linear‑gradient specifications.
// AI Prompts: Write C# code that sets a 60° linear gradient on a WordArt shape and saves the workbook as HTML using Aspose.Cells. | Explain how Aspose.Cells converts GradientFill.Angle to the CSS linear‑gradient angle in the exported HTML file. | Provide a snippet to load a workbook, change a WordArt shape's gradient angle to 30°, and re‑export it to HTML.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a WordArt shape, applies a linear gradient with a custom angle (e.g., 45°), defines two colors, and saves the sheet as HTML. The generated CSS includes the exact gradient angle using the linear‑gradient syntax.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape that already uses a gradient preset (WordArtStyle7)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2, 0,   // topRow, top (pixels)
            2, 0,   // leftColumn, left (pixels)
            200,    // height (pixels)
            100);   // width (pixels)

        // Ensure the shape's fill type is set to Gradient to access GradientFill properties
        wordArt.Fill.FillType = FillType.Gradient;

        // Retrieve the GradientFill object
        GradientFill gradientFill = wordArt.Fill.GradientFill;

        // Set a linear gradient with a custom angle (e.g., 45 degrees)
        // The SetGradient method applies the type, angle, and direction
        gradientFill.SetGradient(GradientFillType.Linear, 45.0, GradientDirectionType.FromCenter);
        // The Angle property can also be set directly if needed
        gradientFill.Angle = 45.0f;

        // Define the two colors for the gradient (optional but demonstrates full effect)
        gradientFill.SetTwoColorGradient(
            Color.Red,               // First color
            Color.Blue,              // Second color
            GradientStyleType.Horizontal,
            1);                      // Variant (1‑4)

        // Convert the workbook to HTML; the generated CSS will include the gradient angle
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportActiveWorksheetOnly = true
        };
        workbook.Save("WordArtGradient.html", htmlOptions);
    }
}
