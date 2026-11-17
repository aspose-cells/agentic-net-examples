// Title: Export WordArt with vertical gradient fill to HTML using Aspose.Cells for .NET (reflection effect unavailable)
// Common Searches: how to keep gradient fill on WordArt when saving a workbook as HTML with Aspose.Cells C# | Aspose.Cells C# export WordArt shape with reflection effect to HTML | preserve WordArt formatting (gradient and reflection) during HTML conversion using Aspose.Cells | C# Aspose.Cells HTML export shape gradient fill not lost | add reflection effect to a shape in Aspose.Cells before exporting to HTML
// Tags: Aspose.Cells HTML export gradient fill | WordArt shape reflection Aspose.Cells | C# shape FillType Gradient Aspose.Cells | preserve shape formatting HTML conversion | Aspose.Cells shape effect support

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape used as WordArt, sets its text and a vertical two‑color gradient fill, notes that reflection effects are not currently supported, and saves the workbook as an HTML file while retaining the gradient appearance.
class WordArtReflectionHtmlExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (placeholder for WordArt)
            // Parameters: shape type, upper left row, column, top, left, height, width (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                0, 0,          // upper left cell
                5, 5,          // top and left offset (pixels)
                100, 200);     // height and width (pixels)

            // Set shape text (acts as WordArt text)
            shape.Text = "Aspose.Cells WordArt";
            shape.Font.Size = 14;
            shape.Font.Color = Color.Black;

            // Apply a two‑color vertical gradient fill (LightBlue → DarkBlue)
            shape.Fill.FillType = FillType.Gradient;
            shape.Fill.GradientFill.SetTwoColorGradient(
                Color.LightBlue,
                Color.DarkBlue,
                GradientStyleType.Vertical,
                0);

            // Note: Reflection effect is not supported in older Aspose.Cells versions.
            // If needed, it can be added when the API provides EffectFormat support.

            // Ensure output directory exists
            string outputPath = "WordArtWithReflection.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, SaveFormat.Html);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
