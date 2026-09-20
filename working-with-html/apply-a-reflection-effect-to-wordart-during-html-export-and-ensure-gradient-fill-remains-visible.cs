// Title: Apply a reflection effect to WordArt with a horizontal two‑color gradient and export the worksheet as HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a TextEffect shape, sets a horizontal two‑color gradient fill, adds a reflection with 50% transparency, and saves the active worksheet to an HTML file with images embedded as Base64. | Show how to configure Aspose.Cells HtmlSaveOptions so that a WordArt shape retains its gradient and reflection when the workbook is exported to HTML. | Provide a step‑by‑step example of setting reflection.Transparency and reflection.Size on a WordArt shape while keeping its gradient fill intact during HTML export.
// Common Searches: asp.net how to keep WordArt gradient fill after exporting to HTML with Aspose.Cells | c# Aspose.Cells reflection effect on TextEffect shape not lost in HTML output | set two color gradient and reflection on WordArt using Aspose.Cells for .NET | export worksheet to HTML with base64 images preserving WordArt styling Aspose.Cells
// Tags: Aspose.Cells reflection effect on TextEffect | two‑color gradient fill for WordArt Aspose.Cells | HTML export with base64 images Aspose.Cells | retain WordArt gradient during HTML save Aspose.Cells | C# configure reflection transparency size Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The program creates a workbook, adds a TextEffect (WordArt) shape with a horizontal two‑color gradient, applies a reflection effect (50% transparency, 70% size), and saves only the active worksheet as an HTML file using HtmlSaveOptions that embed images as Base64.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add WordArt (TextEffect) to the worksheet
            // Parameters: preset effect, text, font name, font size, bold, italic,
            // left, top, width, height, rotation, flip
            var wordArt = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Aspose.Cells",
                "Arial",
                48,
                false,
                false,
                50,
                50,
                500,
                100,
                0,   // rotation
                0);  // flip

            // Apply a horizontal two‑color gradient fill to the WordArt
            var fill = wordArt.Fill;
            // Use the SetTwoColorGradient method (compatible with all Aspose.Cells versions)
            fill.SetTwoColorGradient(Color.Blue, Color.LightBlue, GradientStyleType.Horizontal, 0);

            // Apply a reflection effect (basic settings)
            var reflection = wordArt.Reflection;
            // Type property may not be available in older versions; skip if not supported
            // reflection.Type = ReflectionEffectType.Preset0;
            reflection.Transparency = 0.5; // 50% transparent reflection
            reflection.Size = 0.7;         // 70% size of the original shape

            // Export the workbook to HTML
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true,
                ExportImagesAsBase64 = true
            };

            const string outputPath = "WordArtWithReflection.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
