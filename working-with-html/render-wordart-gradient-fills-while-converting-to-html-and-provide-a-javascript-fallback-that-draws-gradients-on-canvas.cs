// Title: Render WordArt with gradient fill in HTML using Aspose.Cells for .NET and add a JavaScript canvas fallback
// Common Searches: C# Aspose.Cells export WordArt shape to HTML with gradient rendering | How to add a canvas gradient fallback for WordArt when saving Excel to HTML | Using HtmlSaveOptions PresentationPreference to keep WordArt appearance in HTML output | Inject JavaScript into Aspose.Cells generated HTML to draw WordArt gradients on a canvas | Save workbook as HTML and then modify it to include a canvas gradient for WordArt
// Tags: Aspose.Cells WordArt HTML export | HtmlSaveOptions presentation preference usage | Canvas gradient fallback for WordArt | C# inject script into Aspose.Cells HTML output | Excel workbook to HTML with gradient rendering

using System;
using System.IO;
using System.Text;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtGradientHtmlDemo
{
    // The example creates a workbook, adds a WordArt shape, sets a solid fill, embeds a <canvas> element with color data as the shape's HTML representation, saves the workbook to HTML using HtmlSaveOptions with PresentationPreference, injects JavaScript that draws a linear gradient and renders the WordArt text on the canvas, and finally saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Output file paths
            string workbookPath = "WordArtGradientDemo.xlsx";
            string htmlPath = "WordArtGradientDemo.html";

            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Demo";

                // -------------------------------------------------
                // 2. Add a WordArt shape to the worksheet
                // -------------------------------------------------
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt = sheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle1,
                    "Gradient WordArt",
                    2,      // top row
                    10,     // top offset (pixels)
                    2,      // left column
                    10,     // left offset (pixels)
                    100,    // height (pixels)
                    400);   // width (pixels)

                // -------------------------------------------------
                // 3. Apply a solid fill (gradient will be rendered via JavaScript fallback)
                // -------------------------------------------------
                wordArt.Fill.FillType = FillType.Solid;
                // Note: SolidFillColor property may not be available in some versions;
                // the shape will still render with default color.

                // -------------------------------------------------
                // 4. Embed a canvas placeholder for JavaScript fallback
                // -------------------------------------------------
                string canvasHtml = $@"
<canvas class='wordartCanvas' width='400' height='100'
        data-startcolor='#FF0000' data-endcolor='#0000FF'>
</canvas>";
                wordArt.HtmlText = canvasHtml;

                // -------------------------------------------------
                // 5. Save the workbook as HTML with presentation preference
                // -------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,
                    ExportImagesAsBase64 = true,
                    PresentationPreference = true   // more faithful visual output
                };
                workbook.Save(htmlPath, htmlOptions);

                // -------------------------------------------------
                // 6. Append JavaScript that draws the gradient on the canvas
                // -------------------------------------------------
                string jsFallback = @"
<script type='text/javascript'>
function drawWordArtGradients() {
    var canvases = document.querySelectorAll('canvas.wordartCanvas');
    canvases.forEach(function (c) {
        var ctx = c.getContext('2d');
        var startColor = c.getAttribute('data-startcolor') || '#FF0000';
        var endColor   = c.getAttribute('data-endcolor')   || '#0000FF';
        var grad = ctx.createLinearGradient(0, 0, c.width, 0);
        grad.addColorStop(0, startColor);
        grad.addColorStop(1, endColor);
        ctx.fillStyle = grad;
        ctx.fillRect(0, 0, c.width, c.height);

        // Render the WordArt text (simple approximation)
        ctx.font = '30px Arial';
        ctx.fillStyle = '#FFFFFF';
        ctx.textBaseline = 'middle';
        ctx.textAlign = 'center';
        ctx.fillText('Gradient WordArt', c.width / 2, c.height / 2);
    });
}
if (window.addEventListener) {
    window.addEventListener('load', drawWordArtGradients, false);
} else if (window.attachEvent) {
    window.attachEvent('onload', drawWordArtGradients);
}
</script>";

                // Inject the script before </body> if the HTML file exists
                try
                {
                    if (File.Exists(htmlPath))
                    {
                        string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);
                        int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                        if (bodyCloseIndex >= 0)
                        {
                            htmlContent = htmlContent.Insert(bodyCloseIndex, jsFallback);
                        }
                        else
                        {
                            htmlContent += jsFallback;
                        }
                        File.WriteAllText(htmlPath, htmlContent, Encoding.UTF8);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while injecting JavaScript: " + ex.Message);
                }

                // -------------------------------------------------
                // 7. Save the workbook as an XLSX file
                // -------------------------------------------------
                workbook.Save(workbookPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved as '{workbookPath}'.");
                Console.WriteLine($"HTML with WordArt gradient saved as '{htmlPath}'.");
                Console.WriteLine("Open the HTML file in a browser to see the gradient rendered via canvas fallback.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine("StackTrace: " + ex.StackTrace);
            }
        }
    }
}
