// Title: Export WordArt with Gradient Fill and Solid Fallback to HTML using Aspose.Cells for .NET
// Description: Demonstrates how to add a WordArt shape with a preset gradient (WordArtStyle7), assign a solid LightGray fallback for legacy browsers, and save the workbook as a self‑contained HTML file with images embedded as Base64 using Aspose.Cells for .NET.
// Keywords: Aspose.Cells WordArt HTML export | gradient fill fallback | C# export Excel to HTML | Base64 images Aspose.Cells | WordArtStyle7 gradient | solid color fallback HTML | self‑contained HTML report | legacy browser compatibility Excel | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export WordArt gradient to HTML | add solid fallback color for WordArt in HTML | C# save Excel as HTML with embedded images | how to embed WordArt in self‑contained HTML | gradient WordArt not supported in old browsers
// Developer Intent: Create an HTML file that shows a WordArt gradient in modern browsers while providing a solid color fallback for browsers that lack CSS gradient support.
// Use Cases: Design email templates where WordArt appears with a gradient in recent clients but degrades to a solid color in older email readers. | Generate interactive dashboards that retain visual fidelity across both current and legacy web browsers. | Produce portable HTML reports with embedded graphics that work offline without external image files.
// AI Prompts: Write C# code with Aspose.Cells to insert a WordArt shape using WordArtStyle7, set a LightGray solid fallback, and save the workbook as HTML with Base64‑encoded images. | Explain the HtmlSaveOptions settings needed to embed images, export only the active worksheet, and preserve WordArt styling with a fallback color. | Provide a testing checklist to confirm that the generated HTML displays the gradient in Chrome/Edge and the solid fallback in browsers that do not support CSS gradients.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to add a WordArt shape with a preset gradient (WordArtStyle7), assign a solid LightGray fallback for legacy browsers, and save the workbook as a self‑contained HTML file with images embedded as Base64 using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a WordArt shape that uses a preset style containing a gradient fill.
            // WordArtStyle7 = Gradient Fill - Blue, Accent 1, Reflection
            Shape wordArt = worksheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7, // preset gradient style
                "Gradient WordArt",               // displayed text
                2, 0,                            // upper left row, top offset
                2, 0,                            // upper left column, left offset
                200, 400);                       // height, width

            // OPTIONAL: Define a solid fallback color for very old browsers that cannot render gradients.
            // The fallback is added as a CSS style attribute; Aspose.Cells will embed it when exporting to HTML.
            wordArt.Fill.FillType = FillType.Solid;
            wordArt.Fill.SolidFill.Color = Color.LightGray; // fallback solid color

            // Prepare HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export images (including VML resources) as Base64 strings so the HTML file is self‑contained.
                ExportImagesAsBase64 = true,

                // Export only the active worksheet to keep the output simple.
                ExportActiveWorksheetOnly = true
            };

            // Define output file path
            string outputPath = "WordArtGradient.html";

            // Ensure the directory exists (in case a relative path with folders is used)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
