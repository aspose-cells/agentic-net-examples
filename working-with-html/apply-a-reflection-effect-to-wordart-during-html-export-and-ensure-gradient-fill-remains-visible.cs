// Title: Add Custom Reflection to WordArt While Preserving Gradient Fill in HTML Export with Aspose.Cells for .NET
// Description: This example creates a workbook, inserts a WordArt shape using preset style 7 (gradient fill with reflection), customizes reflection properties (type, transparency, size, blur, distance, rotation), and saves the sheet as HTML. The HTML export uses HtmlSaveOptions to embed the WordArt as a Base64 image and keeps all visible shapes and their gradient fill intact.
// Keywords: Aspose.Cells | C# | .NET | WordArt | reflection effect | gradient fill | HTML export | HtmlSaveOptions | Base64 images | shape export | preset WordArt style 7
// Common Searches: Aspose.Cells add reflection to WordArt HTML export | preserve WordArt gradient fill when saving as HTML | customize WordArt reflection properties .NET | export WordArt as Base64 image with Aspose.Cells | HtmlSaveOptions keep shapes visible Aspose.Cells
// Developer Intent: Apply a reflection to a WordArt shape and ensure its gradient fill remains visible when exporting the worksheet to HTML using Aspose.Cells for .NET.
// Use Cases: Design a marketing dashboard where section titles are WordArt with reflection and gradient colors that render correctly in browsers. | Generate an HTML newsletter that includes stylized WordArt headings, preserving both reflection and gradient effects. | Export a spreadsheet containing diagram labels as WordArt, keeping the visual styling intact in the resulting HTML page.
// AI Prompts: Show how to modify the reflection size and blur of a WordArt shape in Aspose.Cells before HTML export. | Provide code to verify that a WordArt shape's gradient fill is retained after saving the workbook as HTML with Base64 images. | Explain which HtmlSaveOptions settings prevent WordArt with reflection from being omitted during HTML export.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This example creates a workbook, inserts a WordArt shape using preset style 7 (gradient fill with reflection), customizes reflection properties (type, transparency, size, blur, distance, rotation), and saves the sheet as HTML. The HTML export uses HtmlSaveOptions to embed the WordArt as a Base64 image and keeps all visible shapes and their gradient fill intact.
class WordArtReflectionHtmlExport
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with preset style 7 (Gradient Fill - Blue, Accent 1, Reflection)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Aspose.Cells Reflection",
            2, 0, 2, 0, 200, 100);

        // Access the reflection effect of the WordArt shape
        ReflectionEffect reflection = wordArt.Reflection;

        // Set custom reflection properties to make the effect more visible
        reflection.Type = ReflectionEffectType.FullReflection4PtOffset; // full reflection with offset
        reflection.Transparency = 0.2;   // low transparency (more opaque)
        reflection.Size = 80;           // larger reflection size
        reflection.Blur = 5;            // slight blur for smoothness
        reflection.Distance = 5;        // distance from the shape
        reflection.RotWithShape = true; // rotate reflection together with shape

        // Ensure the gradient fill remains visible (preset already includes it)
        // If further customization is needed, you can modify the fill format:
        // wordArt.FillFormat.GradientFill.Colors[0] = System.Drawing.Color.Blue;
        // wordArt.FillFormat.GradientFill.Colors[1] = System.Drawing.Color.LightBlue;

        // Prepare HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true,          // embed images (including WordArt) as base64
            IgnoreInvisibleShapes = false,        // export all visible shapes
            ExportHiddenWorksheet = true,         // include hidden worksheets if any
            ExportActiveWorksheetOnly = true,     // export only the active sheet
            IsExportComments = false,             // no comments needed
            ExportGridLines = true,               // show grid lines for better context
            ExportWorkbookProperties = true
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("WordArtReflection.html", htmlOptions);
    }
}
