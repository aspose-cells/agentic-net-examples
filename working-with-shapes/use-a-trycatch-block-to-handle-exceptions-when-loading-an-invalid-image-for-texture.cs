using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class TextureFillInvalidImageDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate texture fill
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);

            // Set the fill type to texture
            shape.Fill.FillType = FillType.Texture;
            TextureFill textureFill = shape.Fill.TextureFill;

            // Path to an image that does not exist or is invalid
            string invalidImagePath = "nonexistent_image.png";

            // Load image data only if the file exists
            if (File.Exists(invalidImagePath))
            {
                try
                {
                    byte[] imageData = File.ReadAllBytes(invalidImagePath);
                    textureFill.ImageData = imageData;
                    Console.WriteLine("Image data loaded successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading image for texture: {ex.Message}");
                    textureFill.Type = TextureType.Unknown;
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {invalidImagePath}");
                textureFill.Type = TextureType.Unknown;
            }

            // Save the workbook
            try
            {
                workbook.Save("TextureFillInvalidImageDemo.xlsx");
                Console.WriteLine("Workbook saved as TextureFillInvalidImageDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}