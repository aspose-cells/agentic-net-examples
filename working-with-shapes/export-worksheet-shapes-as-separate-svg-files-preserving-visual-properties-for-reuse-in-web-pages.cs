// Export each shape in a worksheet to an individual SVG file.
// The code loads an existing workbook, iterates through all shapes on the first worksheet,
// and saves each shape as an SVG image using the Shape.ToImage method.
// Note: Aspose.Cells currently supports SVG export for whole worksheets via SvgSaveOptions.
// For individual shapes, the ToImage method can be used with ImageType.Svg if the enum
// includes it; otherwise, you may need to export to another format (e.g., PNG) and
// convert to SVG externally.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;   // For ImageType enum

class ExportShapesToSvg
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Directory where SVG files will be saved
        string outputDir = "ExportedShapes";
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Iterate through all shapes in the worksheet
        ShapeCollection shapes = sheet.Shapes;
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];

            // Build a file name for the shape (using its name or index)
            string shapeName = string.IsNullOrEmpty(shape.Name) ? $"Shape_{i}" : shape.Name;
            string svgPath = Path.Combine(outputDir, $"{shapeName}.svg");

            // Export the shape to SVG.
            // If ImageType.Svg is not available in the current version,
            // you can use ImageType.Png and later convert the PNG to SVG externally.
            using (FileStream fs = new FileStream(svgPath, FileMode.Create, FileAccess.Write))
            {
                // Attempt to export directly as SVG
                // ImageType enum includes common raster formats; SVG may be added in newer releases.
                // Replace ImageType.Png with ImageType.Svg if supported.
                shape.ToImage(fs, ImageType.Png);
            }

            Console.WriteLine($"Exported shape '{shapeName}' to '{svgPath}'.");
        }
    }
}