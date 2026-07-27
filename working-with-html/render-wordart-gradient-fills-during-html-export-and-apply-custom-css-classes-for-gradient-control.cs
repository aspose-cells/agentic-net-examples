// Title: Render WordArt Gradient in HTML Export with Custom CSS – Aspose.Cells for .NET
// Description: Creates a workbook, adds a WordArt shape with a two‑color horizontal gradient, assigns a custom name that becomes the HTML element ID, and uses HtmlSaveOptions to embed CSS that replaces the generated image with a CSS linear‑gradient. The result is a single‑file HTML page with precise gradient control.
// Keywords: Aspose.Cells | .NET | WordArt | gradient fill | HTML export | custom CSS | HtmlSaveOptions | shape ID | linear-gradient | single file HTML
// Common Searches: Aspose.Cells export WordArt gradient to HTML | apply custom CSS to WordArt shape in HTML output | replace WordArt image with CSS gradient using Aspose.Cells | set HTML element ID for a shape in Aspose.Cells | render WordArt gradient fill in single‑file HTML
// Developer Intent: Export a workbook that contains a WordArt shape with a gradient fill to HTML and control the visual appearance through a custom CSS selector.
// Use Cases: Generate WordArt with a two‑color horizontal gradient and export it as HTML while preserving the gradient via CSS. | Target the exported shape by its HTML ID to apply a linear‑gradient background that matches the original fill. | Modify gradient colors or direction in the workbook and have the changes automatically reflected in the HTML through embedded CSS. | Produce a single‑file HTML document that combines the workbook data and custom styling for easy deployment.
// AI Prompts: Write C# code that adds a WordArt shape with a custom gradient, assigns a name, and saves the workbook as a single‑file HTML with a CSS rule that replaces the shape image with a linear‑gradient using Aspose.Cells. | Explain how to configure HtmlSaveOptions to embed custom CSS for a specific shape ID so the gradient appears correctly in the exported HTML. | Show how to change the gradient colors or direction of a WordArt shape and ensure the HTML export reflects those changes via CSS.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a WordArt shape with a two‑color horizontal gradient, assigns a custom name that becomes the HTML element ID, and uses HtmlSaveOptions to embed CSS that replaces the generated image with a CSS linear‑gradient. The result is a single‑file HTML page with precise gradient control.
class WordArtGradientHtmlExport
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Add a WordArt shape with a preset gradient style (WordArtStyle7)
            //    Parameters: style, text, topRow, top, leftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7,   // Gradient Fill - Blue, Accent 1, Reflection
                "Gradient WordArt",                 // Text displayed in the WordArt
                2, 0,                               // Upper‑left cell (row 2, column 0)
                0, 0,                               // Pixel offsets within the cell
                100, 400);                          // Height and width in pixels

            // 3. Set a custom name; this becomes the HTML element id when exported
            wordArt.Name = "gradientWordArt";

            // 4. Replace the preset fill with a two‑color linear gradient
            wordArt.Fill.FillType = FillType.Gradient;
            GradientFill gradFill = wordArt.Fill.GradientFill;
            gradFill.SetTwoColorGradient(
                Color.FromArgb(0, 112, 192),   // First color (blue)
                Color.FromArgb(255, 255, 255), // Second color (white)
                GradientStyleType.Horizontal,  // Horizontal gradient
                1);                            // Variant 1

            // 5. Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export shapes as images (default). The gradient will be rendered into the image.
                CssStyles = @"
                    /* Custom CSS for the WordArt shape */
                    #gradientWordArt {
                        /* Example: replace the image background with a CSS gradient */
                        background: linear-gradient(to right, #0070C0, #FFFFFF);
                        /* Ensure the element displays as a block with the same size as the shape */
                        display: inline-block;
                        width: 400px;
                        height: 100px;
                    }",
                // Keep the generated HTML in a single file for simplicity.
                SaveAsSingleFile = true
            };

            // 6. Save the workbook as HTML using the configured options
            workbook.Save("WordArtGradient.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
