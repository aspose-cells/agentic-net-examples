// Title: C# – Export Excel WordArt with Gradient to a Single HTML5 File Using Inline SVG (Aspose.Cells)
// Description: This C# example loads an Excel workbook, adds a WordArt shape with a preset gradient, and saves the sheet as a single HTML5 document. HtmlSaveOptions embed all resources as Base64 and render the WordArt as inline SVG, preserving complex gradient fills.
// Keywords: Aspose.Cells | C# Excel to HTML | WordArt gradient export | inline SVG HTML5 | single file HTML | ExportImagesAsBase64 | HtmlSaveOptions | preserve WordArt styling | Aspose.Cells for .NET | Excel WordArt to HTML
// Common Searches: Aspose.Cells export WordArt to HTML5 | C# convert Excel WordArt gradient to inline SVG | Save Excel as single HTML file with embedded images | How to preserve WordArt gradients when exporting to HTML | Inline SVG for Excel shapes Aspose.Cells | Base64 images in HTML export Aspose.Cells
// Developer Intent: Create a self‑contained HTML5 page where Excel WordArt with gradients is rendered as inline SVG.
// Use Cases: Generate web‑ready reports that keep WordArt visual fidelity. | Embed Excel dashboards in web pages without external assets. | Build HTML email templates that include complex WordArt graphics. | Automate documentation of Excel files with styled WordArt for offline viewing.
// AI Prompts: Provide C# code using Aspose.Cells to convert an Excel worksheet containing gradient WordArt into a single HTML5 file with inline SVG and Base64 images. | Explain the impact of HtmlVersion, ExportImagesAsBase64, SaveAsSingleFile, and EnableCssCustomProperties on WordArt gradient rendering in the exported HTML. | Show how to add a WordArt shape with a gradient preset in Aspose.Cells before exporting to HTML. | Describe how to ensure the generated HTML works in major browsers when using inline SVG for WordArt.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# example loads an Excel workbook, adds a WordArt shape with a preset gradient, and saves the sheet as a single HTML5 document. HtmlSaveOptions embed all resources as Base64 and render the WordArt as inline SVG, preserving complex gradient fills.
class ConvertWordArtToHtml
{
    static void Main()
    {
        // Load the existing workbook that contains WordArt (or create a new one if needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient preset (WordArtStyle6) to demonstrate complex gradients
        // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6, // Gradient Fill - Gray
            "Gradient WordArt",
            2,   // topRow
            0,   // top offset in pixels
            2,   // leftColumn
            0,   // left offset in pixels
            100, // height in pixels
            400  // width in pixels
        );

        // Configure HTML save options to embed all resources inline and use HTML5 (supports inline SVG)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            HtmlVersion = HtmlVersion.Html5,          // Use HTML5 for native SVG support
            ExportImagesAsBase64 = true,              // Embed images (including shape renders) as Base64
            SaveAsSingleFile = true,                  // Produce a single HTML file with inline resources
            EnableCssCustomProperties = true          // Optional: reduce duplicate data via CSS custom properties
        };

        // Save the workbook as HTML; the WordArt will be rendered as an inline SVG preserving its gradient
        workbook.Save("output.html", htmlOptions);
    }
}
