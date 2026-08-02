// Title: Set a custom gradient angle for WordArt and preserve it in CSS linear‑gradient during HTML export with Aspose.Cells for .NET
// Description: This example creates a workbook, inserts a WordArt shape, forces a gradient fill, sets a linear gradient at a specific angle (e.g., 45°), defines two colors, and saves the file as HTML. Aspose.Cells automatically generates a CSS linear‑gradient rule that reflects the exact angle, enabling accurate visual rendering in browsers.
// Keywords: Aspose.Cells WordArt gradient | C# gradient angle HTML export | linear-gradient CSS Aspose.Cells | custom WordArt fill .NET | HTML save options gradient angle | Aspose.Cells GradientFill Angle | WordArt CSS linear‑gradient | Aspose.Cells HTML conversion | C# Aspose.Cells example
// Common Searches: how to change WordArt gradient angle in Aspose.Cells | Aspose.Cells HTML export CSS linear‑gradient angle | set custom gradient fill for WordArt .NET | preserve WordArt styling when saving as HTML | C# Aspose.Cells gradient fill angle
// Developer Intent: Apply a specific gradient angle to a WordArt shape and have the exported HTML contain the matching CSS linear‑gradient declaration.
// Use Cases: Design marketing reports where WordArt must display a 45° orange‑to‑purple gradient in the browser. | Programmatically adjust WordArt gradient direction based on user input before generating HTML output. | Create branded dashboards that require consistent gradient angles across Excel and HTML views.
// AI Prompts: Show C# code to set a 30‑degree gradient on a WordArt shape with Aspose.Cells and export it to HTML with the correct linear‑gradient CSS. | Explain how Aspose.Cells maps GradientFill.Angle to the CSS linear‑gradient angle when saving a workbook as HTML. | Generate a snippet that lets a user choose a gradient angle for WordArt, applies it via GradientFill, and saves the workbook to HTML.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, inserts a WordArt shape, forces a gradient fill, sets a linear gradient at a specific angle (e.g., 45°), defines two colors, and saves the file as HTML. Aspose.Cells automatically generates a CSS linear‑gradient rule that reflects the exact angle, enabling accurate visual rendering in browsers.
class GradientWordArtHtmlDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape using a preset style that already has a gradient fill
            // Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
                "Custom Gradient WordArt",
                2,   // topRow
                0,   // top (pixel offset)
                2,   // leftColumn
                0,   // left (pixel offset)
                200, // height (pixels)
                400  // width (pixels)
            );

            // Ensure the shape's fill type is set to Gradient
            wordArt.Fill.FillType = FillType.Gradient;

            // Access the GradientFill object
            GradientFill gradientFill = wordArt.Fill.GradientFill;

            // Set the gradient type to Linear and define a custom angle (e.g., 45 degrees)
            gradientFill.SetGradient(GradientFillType.Linear, 45.0, GradientDirectionType.FromCenter);
            // Alternatively, you can set the Angle property directly
            gradientFill.Angle = 45.0f;

            // Define the two colors for the gradient (optional but makes the effect visible)
            gradientFill.SetTwoColorGradient(
                Color.Orange,          // First color
                Color.Purple,          // Second color
                GradientStyleType.Horizontal,
                1                      // Variant
            );

            // Convert the workbook to HTML.
            // Aspose.Cells will generate CSS that includes the linear-gradient with the specified angle.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save("GradientWordArt.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
