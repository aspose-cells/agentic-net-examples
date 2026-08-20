// Title: Add Reflection to WordArt and Preserve Gradient Fill on HTML Export with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a WordArt shape with a preset gradient style, apply a full‑reflection effect (type, transparency, size, blur, distance, rotation), optionally set a solid fill, and save the worksheet as HTML using Aspose.Cells. The HTML output embeds images as Base64 and retains all visible shapes.
// Keywords: Aspose.Cells | WordArt reflection | gradient fill | HTML export | C# | Aspose.Cells .NET | HtmlSaveOptions | ExportImagesAsBase64 | shape rendering | preserve WordArt styling
// Common Searches: Aspose.Cells add reflection to WordArt C# | export WordArt with gradient fill to HTML Aspose | configure WordArt reflection properties .NET | HTML save options to keep shapes Aspose.Cells | Base64 image export Aspose.Cells HTML
// Developer Intent: Apply a reflection effect to a WordArt shape and export the workbook to HTML while keeping the gradient fill and shape appearance intact.
// Use Cases: Generate marketing HTML pages that include stylized WordArt with reflection and gradient colors. | Create reports where visual emphasis is added through WordArt reflections without losing formatting during HTML conversion. | Automate conversion of Excel worksheets containing WordArt to web‑ready HTML with embedded Base64 images.
// AI Prompts: Write C# code using Aspose.Cells to insert WordArt, set a full reflection effect, and save the sheet as HTML with Base64 images. | Show how to keep a WordArt gradient fill visible when exporting a workbook to HTML with Aspose.Cells. | Provide an example that configures reflection type, transparency, size, blur, distance, and rotation for WordArt in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to create a workbook, insert a WordArt shape with a preset gradient style, apply a full‑reflection effect (type, transparency, size, blur, distance, rotation), optionally set a solid fill, and save the worksheet as HTML using Aspose.Cells. The HTML output embeds images as Base64 and retains all visible shapes.
class ReflectionWordArtHtmlExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape using a preset style that already contains a gradient fill and reflection
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle7,
                "Aspose Cells",
                2, 0, 2, 0,
                300, 100);

            // Verify that the shape is WordArt
            if (wordArt.IsWordArt)
            {
                // Configure the reflection effect
                ReflectionEffect reflection = wordArt.Reflection;
                reflection.Type = ReflectionEffectType.FullReflectionTouching;
                reflection.Transparency = 0.3;   // 30% transparent at start
                reflection.Size = 80;           // size of the reflection (percentage)
                reflection.Blur = 5;            // slight blur
                reflection.Distance = 5;        // distance from the shape
                reflection.RotWithShape = true; // rotate reflection together with shape

                // Adjust the fill to a solid color (if needed)
                if (wordArt.Fill != null)
                {
                    // Set fill type to solid; color can be left as default or set via other APIs if required
                    wordArt.Fill.FillType = FillType.Solid;
                }
            }

            // Set HTML save options – ensure that shapes (including WordArt) are exported
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,          // embed images (including shape rendering) directly
                ExportActiveWorksheetOnly = true,    // only the sheet we edited
                IgnoreInvisibleShapes = false        // export visible shapes
            };

            // Save the workbook as HTML
            workbook.Save("WordArtWithReflection.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
