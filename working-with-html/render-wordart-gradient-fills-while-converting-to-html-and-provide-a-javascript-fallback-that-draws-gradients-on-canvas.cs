using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace WordArtGradientHtmlDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Add a TextBox shape that will hold the WordArt text
            //    Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape textBox = sheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 100);

            // 4. Set the text that will be displayed
            textBox.TextBody.Text = "Aspose Cells WordArt";

            // 5. Apply a preset WordArt style that contains a gradient fill
            //    WordArtStyle7 = Gradient Fill - Blue, Accent 1, Reflection
            textBox.TextBody.SetWordArtStyle(PresetWordArtStyle.WordArtStyle7);

            // 6. Save the workbook as HTML
            string htmlPath = "WordArtGradient.html";
            workbook.Save(htmlPath, SaveFormat.Html);

            // 7. Inject a canvas element with JavaScript fallback for browsers that cannot render the WordArt image
            //    The script draws a simple linear gradient that mimics the WordArt colors.
            string canvasHtml = @"
<div style='margin-top:20px;'>
    <canvas id='wordArtCanvas' width='400' height='100'></canvas>
</div>
<script>
    (function(){
        var canvas = document.getElementById('wordArtCanvas');
        if (!canvas.getContext) return;
        var ctx = canvas.getContext('2d');

        // Create linear gradient (left to right)
        var grad = ctx.createLinearGradient(0, 0, canvas.width, 0);
        // Gradient colors approximating WordArtStyle7 (Blue to LightBlue)
        grad.addColorStop(0, '#1F4E79');   // Darker blue
        grad.addColorStop(1, '#6FA8DC');   // Lighter blue

        // Fill the canvas with the gradient
        ctx.fillStyle = grad;
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        // Optional: draw the same text over the gradient
        ctx.font = '30px Arial';
        ctx.fillStyle = '#FFFFFF';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('Aspose Cells WordArt', canvas.width/2, canvas.height/2);
    })();
</script>";

            // 8. Append the canvas HTML to the generated HTML file
            File.AppendAllText(htmlPath, canvasHtml);

            Console.WriteLine($"HTML file with WordArt and canvas fallback saved to: {Path.GetFullPath(htmlPath)}");
        }
    }
}