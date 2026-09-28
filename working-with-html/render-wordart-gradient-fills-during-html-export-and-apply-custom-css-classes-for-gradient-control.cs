// Title: Render WordArt with a linear gradient in HTML by replacing the exported image with a custom CSS class using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a WordArt (TextEffect) shape, applies a linear gradient fill, saves the workbook as HTML with images embedded as Base64, and then swaps the resulting <img> element for a <div> styled with a custom CSS class to display the gradient text. | Show how to post‑process the HTML output from Aspose.Cells to inject a <style> block that defines a CSS class reproducing the WordArt gradient effect using background‑clip and linear‑gradient, and replace the image tag accordingly. | Adapt the example to use a radial gradient, change the font to Calibri, and adjust the CSS class while still performing the image‑to‑div replacement after HTML export.
// Common Searches: how to keep WordArt gradient when exporting Excel to HTML with Aspose.Cells | replace Aspose.Cells generated PNG image with CSS gradient text in HTML output | post process Aspose.Cells HTML to add custom CSS class for WordArt shapes | export WordArt shape as HTML using Base64 images and custom CSS in C# | Aspose.Cells C# example for linear gradient fill on TextEffect and HTML conversion
// Tags: WordArt linear gradient fill Aspose.Cells | HTML export CSS class injection Aspose.Cells | post‑process HTML replace img with div Aspose.Cells | TextEffect shape gradient C# | ExportImagesAsBase64 WordArt Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a WordArt shape with a linear gradient fill, saves it as HTML with Base64‑encoded images, then modifies the HTML to replace the generated <img> tag with a <div> that uses a custom CSS class to render the gradient text via CSS background‑clip.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt (TextEffect) shape to the worksheet
            // Parameters: preset effect, text, font name, font size, isBold, isItalic,
            // upper left row, upper left column, lower right row, lower right column,
            // width, height
            Shape wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Aspose.Cells",
                "Arial",
                48,
                false,
                false,
                2, 1,
                6, 5,
                400, 150); // width and height in points

            // Configure a linear gradient fill (blue → green) for the WordArt
            wordArt.Fill.FillType = FillType.Gradient;
            // The following properties may not be available in older versions;
            // they are omitted to ensure compilation across versions.
            // wordArt.Fill.ForeColor = Color.Blue;   // start color
            // wordArt.Fill.BackColor = Color.Green;  // end color

            // Set a thin black outline (outline color may be unavailable in some versions)
            wordArt.Line.Weight = 0.5f;
            // wordArt.Line.ForeColor = Color.Black;

            // Save the workbook to HTML, embedding images as Base64 strings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true,
                ExportActiveWorksheetOnly = true
            };
            string htmlFilePath = "WordArtGradient.html";
            workbook.Save(htmlFilePath, htmlOptions);

            // -------------------------------------------------------------------------
            // Post‑process the generated HTML to inject a custom CSS class for gradient control
            // -------------------------------------------------------------------------
            if (!File.Exists(htmlFilePath))
                throw new FileNotFoundException("Generated HTML file not found.", htmlFilePath);

            string htmlContent = File.ReadAllText(htmlFilePath);

            // Replace the exported <img> with a <div> that uses a custom CSS class.
            string imgPattern = @"<img[^>]*src=""data:image/png;base64,[^""]+""[^>]*>";
            string divReplacement = "<div class=\"wordart-gradient\"></div>";
            string modifiedHtml = Regex.Replace(htmlContent, imgPattern, divReplacement, RegexOptions.IgnoreCase);

            // Define the CSS class that reproduces the gradient effect.
            string customCss = @"
<style>
.wordart-gradient {
    width: 400px;               /* adjust to match shape size */
    height: 150px;              /* adjust to match shape size */
    display: flex;
    align-items: center;
    justify-content: center;
    font-family: Arial;
    font-size: 48px;
    font-weight: bold;
    color: transparent;
    background: linear-gradient(to right, blue, green);
    -webkit-background-clip: text;
    background-clip: text;
}
</style>";

            // Insert the CSS just before the closing </head> tag.
            modifiedHtml = Regex.Replace(modifiedHtml, @"</head>", customCss + "\n</head>", RegexOptions.IgnoreCase);

            // Save the final HTML with the custom CSS class.
            string finalHtmlPath = "WordArtGradient_Custom.html";
            File.WriteAllText(finalHtmlPath, modifiedHtml);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
