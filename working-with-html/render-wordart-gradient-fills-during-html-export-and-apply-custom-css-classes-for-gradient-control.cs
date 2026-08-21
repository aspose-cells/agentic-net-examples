// Title: Export WordArt with Gradient Fill to HTML and Control It via Custom CSS – Aspose.Cells C#
// Description: Demonstrates how to add a WordArt shape, apply a two‑color gradient, assign an ID, and export the workbook to HTML using HtmlSaveOptions.CssStyles to inject a CSS rule that renders the shape with a linear‑gradient background.
// Keywords: Aspose.Cells | C# WordArt gradient | HTML export CSS | HtmlSaveOptions.CssStyles | custom CSS for WordArt | linear gradient background | preserve shape styling in HTML | .NET spreadsheet export
// Common Searches: Aspose.Cells export WordArt gradient to HTML | how to apply custom CSS to WordArt in Aspose.Cells | preserve WordArt fill when saving as HTML | override WordArt image with CSS gradient | HtmlSaveOptions CssStyles example C#
// Developer Intent: Generate HTML from a workbook while keeping a WordArt shape’s gradient appearance under full CSS control.
// Use Cases: Create a WordArt shape, set a horizontal two‑color gradient, give it a unique ID, and export to HTML with a CSS rule that applies a matching linear‑gradient background. | Modify gradient colors or direction directly in the injected CSS without altering the shape’s internal fill settings. | Assign different CSS classes to multiple WordArt objects in the same sheet to produce varied gradient effects in the resulting HTML page.
// AI Prompts: Write C# code using Aspose.Cells to add a WordArt shape with a vertical two‑color gradient, assign a unique name, and export the workbook to HTML with a CssStyles rule that applies a matching linear‑gradient background. | Explain how HtmlSaveOptions.CssStyles can replace the rendered WordArt image with a CSS gradient, including examples of changing direction, colors, and opacity. | Provide a step‑by‑step guide for exporting several WordArt shapes, each with its own CSS gradient class, using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to add a WordArt shape, apply a two‑color gradient, assign an ID, and export the workbook to HTML using HtmlSaveOptions.CssStyles to inject a CSS rule that renders the shape with a linear‑gradient background.
class WordArtGradientHtmlExport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style (WordArtStyle7)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2,          // topRow
            0,          // top (pixels)
            2,          // leftColumn
            0,          // left (pixels)
            100,        // height (pixels)
            400);       // width (pixels)

        // Assign a name to the shape so we can reference it in CSS
        wordArt.Name = "gradientWordArt";

        // Optionally fine‑tune the gradient fill (two‑color horizontal gradient)
        wordArt.Fill.FillType = FillType.Gradient;
        GradientFill gradientFill = wordArt.Fill.GradientFill;
        gradientFill.SetTwoColorGradient(
            Color.Blue,               // first color
            Color.LightBlue,          // second color
            GradientStyleType.Horizontal,
            1);                       // variant

        // Prepare HTML save options and inject a custom CSS class that
        // overrides the background of the exported shape using a CSS gradient.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // The shape is rendered as an <img> element with the assigned name as its id.
        // The CSS below applies a linear gradient background to that element.
        htmlOptions.CssStyles = @"
#gradientWordArt {
    background: linear-gradient(to right, #0000FF, #ADD8E6) !important;
    /* Ensure the gradient covers the whole element */
    background-size: 100% 100% !important;
}";
        // Export the workbook to HTML with the custom CSS applied
        workbook.Save("WordArtGradient.html", htmlOptions);
    }
}
