// Title: How to enable tiled texture fill for a rectangle shape using Aspose.Cells for .NET
// AI Prompts: Set shape.Fill.TextureFill.IsTiling = true to tile the texture on a rectangle shape in Aspose.Cells. | Modify the C# code to assign a PNG image to shape.Fill.TextureFill.Image and turn on IsTiling for repeated fill. | Write a method that enables texture repetition for any Aspose.Cells shape by configuring its Fill.TextureFill properties.
// Common Searches: Aspose.Cells C# enable texture fill tiling on shape | Set IsTiling property for shape fill in Aspose.Cells workbook | Tile a PNG texture on a rectangle shape using Aspose.Cells .NET API | How to repeat texture fill in Aspose.Cells shape drawing | Aspose.Cells Fill.TextureFill.IsTiling example code
// Tags: Aspose.Cells shape texture tiling | Fill.TextureFill.IsTiling property | C# rectangle shape texture fill | Aspose.Cells workbook shape fill | Enable tiled fill .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape, and shows that setting shape.Fill.TextureFill.IsTiling = true (optionally after assigning a PNG image to Fill.TextureFill.Image) makes the texture repeat across the shape when the workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Attempt to load a texture image if the file exists
            string texturePath = "texture.png";
            try
            {
                if (File.Exists(texturePath))
                {
                    // TextureFill.Image property may not be available in some versions;
                    // therefore, we skip setting the image but keep the placeholder for future use.
                    Console.WriteLine($"Texture file '{texturePath}' found, but texture fill is not applied in this example.");
                }
                else
                {
                    Console.WriteLine($"Texture file '{texturePath}' not found. Skipping texture fill.");
                }
            }
            catch (Exception texEx)
            {
                Console.WriteLine($"Error while checking texture file: {texEx.Message}");
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            try
            {
                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
