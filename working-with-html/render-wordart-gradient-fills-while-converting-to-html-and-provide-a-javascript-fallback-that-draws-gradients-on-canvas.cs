// Title: Render WordArt Gradient in HTML with Canvas Fallback – Aspose.Cells for .NET
// Description: Adds a WordArt TextBox with preset gradient (WordArtStyle7), saves the workbook as HTML, then injects JavaScript that replaces the generated shape image with a canvas rendering the same linear gradient.
// Keywords: Aspose.Cells | C# | WordArt | gradient fill | HTML export | canvas fallback | JavaScript gradient | PresetWordArtStyle | shape image replacement | Excel to HTML
// Common Searches: export WordArt gradient to HTML using Aspose.Cells | JavaScript canvas fallback for Aspose.Cells WordArt | replace Aspose.Cells shape image with canvas gradient | preserve WordArtStyle7 gradient in HTML output | Aspose.Cells HTML shape rendering issue
// Developer Intent: Generate HTML from an Excel workbook that keeps the WordArt gradient appearance and provide a JavaScript canvas fallback to redraw the gradient when the original shape image cannot be displayed.
// Use Cases: Convert an Excel sheet containing WordArt to HTML while maintaining visual fidelity of gradient fills. | Automatically insert a script into Aspose.Cells HTML output that swaps shape images for canvas elements drawing matching gradients. | Create a reusable helper that adds gradient‑aware fallbacks for any WordArt shape exported by Aspose.Cells.
// AI Prompts: Write a C# method that adds a WordArt TextBox with a preset gradient, saves the workbook as HTML, and injects JavaScript to replace the shape image with a canvas drawing the same gradient. | Generate JavaScript code that locates the <img> tag for a WordArt shape in Aspose.Cells HTML and replaces it with a canvas element rendering a linear gradient matching WordArtStyle7. | Explain how to modify Aspose.Cells‑generated HTML to include a fallback script that draws WordArt gradients on a canvas for browsers that cannot display the original shape image.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace WordArtGradientHtmlDemo
{
    // Adds a WordArt TextBox with preset gradient (WordArtStyle7), saves the workbook as HTML, then injects JavaScript that replaces the generated shape image with a canvas rendering the same linear gradient.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape that will hold the WordArt
            // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape textBox = sheet.Shapes.AddTextBox(2, 0, 2, 0, 120, 400);
            textBox.Name = "MyWordArt";

            // Set the text for the WordArt
            FontSettingCollection fontSettings = textBox.TextBody;
            fontSettings.Text = "Gradient WordArt";

            // Apply a preset WordArt style that includes a gradient fill (WordArtStyle7)
            fontSettings.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

            // Save the workbook as HTML
            string htmlPath = "WordArt.html";
            workbook.Save(htmlPath, SaveFormat.Html);

            // Load the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);

            // JavaScript fallback that draws a matching gradient on a canvas element
            string fallbackScript = @"
<script>
window.addEventListener('load', function () {
    // Locate the first image generated for the shape (Aspose uses <img> tags for shapes)
    var img = document.querySelector('img[alt=""MyWordArt""]');
    if (!img) return;

    // Create a canvas with the same dimensions as the image
    var canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;

    // Draw a linear gradient that approximates WordArtStyle7 (Blue to Accent1)
    var ctx = canvas.getContext('2d');
    var grad = ctx.createLinearGradient(0, 0, canvas.width, 0);
    grad.addColorStop(0, '#00B0F0'); // Approximate Accent1 blue
    grad.addColorStop(1, '#FFFFFF'); // White (reflection effect)
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Replace the image with the canvas
    img.parentNode.replaceChild(canvas, img);
});
</script>";

            // Insert the fallback script before the closing </body> tag
            int bodyCloseIndex = htmlContent.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
            if (bodyCloseIndex >= 0)
            {
                htmlContent = htmlContent.Insert(bodyCloseIndex, fallbackScript);
            }
            else
            {
                // If </body> not found, append at the end
                htmlContent += fallbackScript;
            }

            // Write the modified HTML back to disk
            File.WriteAllText(htmlPath, htmlContent);

            Console.WriteLine("HTML with WordArt gradient and JavaScript fallback has been generated at: " + Path.GetFullPath(htmlPath));
        }
    }
}
