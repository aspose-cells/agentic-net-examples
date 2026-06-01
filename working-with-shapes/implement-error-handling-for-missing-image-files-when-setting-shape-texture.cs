using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeTextureErrorHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape that will receive the texture fill
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
                shape.Fill.FillType = FillType.Texture; // Enable texture fill

                // Path to the image that should be used as texture
                string imagePath = Path.Combine("Images", "texture.png");

                // Load image data; use placeholder if file is missing
                if (File.Exists(imagePath))
                {
                    shape.Fill.TextureFill.ImageData = File.ReadAllBytes(imagePath);
                    Console.WriteLine($"Texture image loaded from '{imagePath}'.");
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{imagePath}' not found. Using a default 1x1 pixel PNG as texture.");
                    const string base64Placeholder = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lXcV6wAAAABJRU5ErkJggg==";
                    shape.Fill.TextureFill.ImageData = Convert.FromBase64String(base64Placeholder);
                }

                // Optional: adjust additional texture properties
                shape.Fill.TextureFill.IsTiling = true;
                shape.Fill.TextureFill.Scale = 0.5;

                // Save the workbook
                string outputPath = "ShapeTextureErrorHandlingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeTextureErrorHandlingDemo.Run();
        }
    }
}