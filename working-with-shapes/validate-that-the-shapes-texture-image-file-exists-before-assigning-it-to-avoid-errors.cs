using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape rectangle = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 100);

            // Set the fill type to Texture so we can assign an image
            rectangle.Fill.FillType = FillType.Texture;

            // Path to the texture image file
            string imagePath = Path.Combine(Environment.CurrentDirectory, "texture.png");

            // Validate that the image file exists before assigning it
            if (File.Exists(imagePath))
            {
                // Read the image bytes and assign to the shape's texture fill
                byte[] imageData = File.ReadAllBytes(imagePath);
                rectangle.Fill.TextureFill.ImageData = imageData;

                // Optional: configure additional texture fill properties
                rectangle.Fill.TextureFill.IsTiling = true;
                rectangle.Fill.TextureFill.Scale = 0.5;
            }
            else
            {
                Console.WriteLine($"Texture image file not found: {imagePath}");
                // Handle the missing file scenario as needed (e.g., use a default texture or skip assignment)
            }

            // Save the workbook
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TextureShapeDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}