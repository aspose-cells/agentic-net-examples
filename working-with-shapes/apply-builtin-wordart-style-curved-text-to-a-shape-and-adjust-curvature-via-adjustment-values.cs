using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            ShapeCollection shapes = worksheet.Shapes;

            // Add a WordArt (text effect) shape
            // Parameters: effect, text, font name, size, bold, italic,
            // topRow, top offset, leftColumn, left offset, height, width
            Shape wordArt = shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Curved Text",
                "Arial",
                36,
                false,
                false,
                2, 0,   // topRow, top offset
                2, 0,   // leftColumn, left offset
                200,    // height (pixels)
                400);   // width (pixels)

            // Set the preset shape to a curved WordArt style (ArchUpCurve)
            TextEffectFormat textEffect = wordArt.TextEffect;
            textEffect.PresetShape = MsoPresetTextEffectShape.ArchUpCurve;

            // NOTE: Adjustments property is not available in all Aspose.Cells versions.
            // If needed, curvature can be modified via Adjustments when supported.
            // The following block is omitted to maintain compatibility.

            // Define output file path
            string outputPath = "CurvedWordArt.xlsx";

            // Ensure the directory exists (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the WordArt shape
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}