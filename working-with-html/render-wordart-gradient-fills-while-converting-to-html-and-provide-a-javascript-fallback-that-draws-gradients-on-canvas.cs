// Title: C# – Export WordArt Gradient to HTML with Canvas JavaScript Fallback using Aspose.Cells
// Description: Demonstrates how to create a workbook, add a TextBox with PresetWordArtStyle.WordArtStyle7 (gradient fill), save it as a single‑file HTML (images embedded as Base64), and inject a positioned canvas element plus JavaScript that redraws the same gradient and text as a fallback for browsers that cannot render the original shape.
// Keywords: Aspose.Cells | C# | WordArt gradient | HTML export | canvas fallback | JavaScript gradient | ExportImagesAsBase64 | PresetWordArtStyle.WordArtStyle7 | Excel to HTML | single file HTML
// Common Searches: Aspose.Cells export WordArt to HTML | preserve WordArt gradient in HTML | add canvas fallback for Excel WordArt | C# generate HTML with embedded images | draw linear gradient on canvas with JavaScript
// Developer Intent: Create a single‑file HTML version of an Excel workbook that contains gradient‑filled WordArt and add a JavaScript canvas fallback that reproduces the same visual effect.
// Use Cases: Convert Excel worksheets with WordArt to portable HTML for web publishing. | Ensure gradient‑filled WordArt displays correctly in browsers lacking native shape support. | Embed all images and styles as Base64 to simplify deployment and avoid external assets. | Programmatically inject custom JavaScript into the generated HTML for post‑processing. | Provide a lightweight canvas rendering for mobile or low‑power browsers.
// AI Prompts: Write C# code that adds a TextBox shape, applies PresetWordArtStyle.WordArtStyle7, saves the workbook as HTML with ExportImagesAsBase64 = true, then inserts a canvas element with JavaScript that draws the same gradient and centered white text. | Generate JavaScript that creates a linear gradient from light blue to dark blue on a 400 × 100 canvas, fills the rectangle, and renders bold white text centered in the canvas. | Explain how to locate the closing </body> tag in the saved HTML, insert custom canvas markup before it, and fall back to appending at the end if the tag is missing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, add a TextBox with PresetWordArtStyle.WordArtStyle7 (gradient fill), save it as a single‑file HTML (images embedded as Base64), and inject a positioned canvas element plus JavaScript that redraws the same gradient and text as a fallback for browsers that cannot render the original shape.
class WordArtGradientHtmlDemo
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Add a TextBox shape that will hold the WordArt text
            // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, width, height
            Shape textBox = sheet.Shapes.AddTextBox(2, 0, 2, 0, 400, 100);
            textBox.TextBody.Text = "Aspose Cells WordArt";

            // 3. Apply a preset WordArt style that contains a gradient fill (WordArtStyle7)
            textBox.TextBody.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

            // 4. Save the workbook as HTML
            string htmlPath = "WordArtGradient.html";
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Keep the HTML in a single file (styles and images embedded as base64)
                ExportImagesAsBase64 = true
            };
            workbook.Save(htmlPath, saveOptions);

            // 5. Read the generated HTML, inject a canvas fallback and JavaScript that draws the same gradient
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("Generated HTML file not found.", htmlPath);

            string htmlContent = File.ReadAllText(htmlPath);

            // Define a simple canvas element that matches the shape size
            string canvasHtml = @"
<div style='position:relative; width:400px; height:100px;'>
    <canvas id='wordartCanvas' width='400' height='100' style='position:absolute; left:0; top:0;'></canvas>
    <div id='wordartFallback' style='position:absolute; left:0; top:0; width:400px; height:100px;'></div>
</div>
<script>
    // JavaScript fallback: draw the same gradient on the canvas
    (function() {
        var canvas = document.getElementById('wordartCanvas');
        if (!canvas.getContext) return;
        var ctx = canvas.getContext('2d');

        // Create a linear gradient that mimics the WordArtStyle7 (blue to accent1)
        var gradient = ctx.createLinearGradient(0, 0, canvas.width, 0);
        gradient.addColorStop(0, '#ADD8E6'); // LightBlue
        gradient.addColorStop(1, '#00008B'); // DarkBlue

        // Fill the canvas with the gradient
        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        // Optional: add the text on top of the gradient
        ctx.font = 'bold 36px Arial';
        ctx.fillStyle = '#FFFFFF';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('Aspose Cells WordArt', canvas.width / 2, canvas.height / 2);
    })();
</script>";

            // Insert the canvas HTML just before the closing </body> tag
            int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (bodyCloseIndex >= 0)
            {
                htmlContent = htmlContent.Insert(bodyCloseIndex, canvasHtml);
            }
            else
            {
                // Fallback: append at the end
                htmlContent += canvasHtml;
            }

            // 6. Write the modified HTML back to file
            File.WriteAllText(htmlPath, htmlContent);

            Console.WriteLine("HTML with WordArt gradient and JavaScript fallback generated at: " + Path.GetFullPath(htmlPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
