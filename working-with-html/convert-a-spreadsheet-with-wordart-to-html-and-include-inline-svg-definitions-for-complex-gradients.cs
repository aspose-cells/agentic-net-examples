// Title: Export WordArt with Gradient and Inline SVG to HTML5 using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a gradient WordArt shape, embed an SVG that defines a multi‑stop linear gradient, configure HtmlSaveOptions for HTML5, and save the file so the SVG appears inline in the generated HTML.
// Keywords: Aspose.Cells | C# | WordArt export | inline SVG | HTML5 conversion | gradient shape | HtmlSaveOptions | Excel to HTML | SVG shape in Excel
// Common Searches: Aspose.Cells export WordArt to HTML5 | inline SVG with Aspose.Cells C# | save Excel workbook as HTML with gradient graphics | add SVG shape to worksheet and convert to HTML | preserve WordArt styling in HTML export
// Developer Intent: Generate an HTML5 file that contains both a gradient WordArt object and an embedded SVG definition without external image files.
// Use Cases: Web‑ready reports that keep Excel‑designed WordArt and custom SVG gradients. | Interactive dashboards where graphics are rendered directly in the browser via inline SVG. | Automated conversion of design‑heavy Excel templates into single‑page HTML for newsletters or intranet portals.
// AI Prompts: Write C# code with Aspose.Cells to add a gradient WordArt shape, embed an SVG with a multi‑stop linear gradient, and export the workbook to HTML5 with the SVG inline. | Explain the effect of HtmlSaveOptions properties HtmlVersion, ExportImagesAsBase64, and EnableCssCustomProperties on the output when exporting WordArt and SVG. | Show how to replace the linear gradient in the SVG with a radial gradient while keeping the shape inline after HTML export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a gradient WordArt shape, embed an SVG that defines a multi‑stop linear gradient, configure HtmlSaveOptions for HTML5, and save the file so the SVG appears inline in the generated HTML.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient preset (WordArtStyle6)
        ShapeCollection shapes = worksheet.Shapes;
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6, // gradient fill style
            "Gradient WordArt",               // text
            2, 0,                             // topRow, top offset (pixels)
            2, 0,                             // leftColumn, left offset (pixels)
            100, 400);                        // height, width (pixels)

        // Define an SVG image that contains a complex linear gradient
        string svgContent = @"
<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100'>
  <defs>
    <linearGradient id='grad1' x1='0%' y1='0%' x2='100%' y2='0%'>
      <stop offset='0%'   style='stop-color:#ff0000;stop-opacity:1' />
      <stop offset='50%'  style='stop-color:#00ff00;stop-opacity:1' />
      <stop offset='100%' style='stop-color:#0000ff;stop-opacity:1' />
    </linearGradient>
  </defs>
  <rect width='200' height='100' fill='url(#grad1)' />
</svg>";
        byte[] svgBytes = System.Text.Encoding.UTF8.GetBytes(svgContent);

        // Add the SVG shape to the worksheet (compatibleImageData is null for modern Excel versions)
        shapes.AddSvg(
            topRow: 5, top: 0,
            leftColumn: 5, left: 0,
            height: -1, width: -1,          // -1 lets Excel auto‑size the shape
            svgData: svgBytes,
            compatibleImageData: null);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HtmlVersion = HtmlVersion.Html5;          // Enable HTML5 for inline SVG support
        htmlOptions.ExportImagesAsBase64 = false;            // Keep images as separate files (SVG will be inline)
        htmlOptions.EnableCssCustomProperties = true;        // Optional: reduce duplicate resources
        htmlOptions.ExportWorksheetCSSSeparately = false;    // Keep CSS in the same file

        // Save the workbook as an HTML file
        string outputHtml = "WordArtWithSvg.html";
        workbook.Save(outputHtml, htmlOptions);

        Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputHtml)}");
    }
}
