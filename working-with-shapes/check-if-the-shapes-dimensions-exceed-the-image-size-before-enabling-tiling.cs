// Title: Check shape size against image size before enabling texture tiling in Aspose.Cells for .NET
// Description: This example creates a workbook, loads an image, adds a rectangle shape, applies the image as a texture fill, reads the image dimensions via a temporary picture, compares the shape's width and height with the image's size, sets the Fill.TextureFill.IsTiling flag accordingly, and saves the file.
// Keywords: Aspose.Cells texture tiling | shape size comparison | IsTiling property C# | retrieve picture dimensions | .NET Excel shape fill | image size check Aspose.Cells
// Common Searches: Aspose.Cells enable tiling only when shape larger than image | how to get picture width height in Aspose.Cells .NET | compare shape dimensions with image before setting IsTiling | temporary picture to read image size Aspose.Cells
// Developer Intent: Determine whether a shape's texture fill should repeat by evaluating the shape's dimensions relative to the source image.
// Use Cases: Prevent unnecessary tiling when the shape fits within the image. | Automatically repeat the texture for shapes larger than the source graphic. | Validate image dimensions to avoid runtime errors in texture fills.
// AI Prompts: Write C# code using Aspose.Cells that compares a shape's width and height with an image's dimensions and toggles Fill.TextureFill.IsTiling. | Suggest a way to obtain image size without inserting a temporary picture in Aspose.Cells. | Explain how to scale a texture instead of tiling when the shape is smaller than the image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, loads an image, adds a rectangle shape, applies the image as a texture fill, reads the image dimensions via a temporary picture, compares the shape's width and height with the image's size, sets the Fill.TextureFill.IsTiling flag accordingly, and saves the file.
    public class ShapeTilingCheckDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file to be used as texture
                string imagePath = "sample.png"; // replace with your image file path

                // Verify that the image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Read image bytes
                byte[] imageData = File.ReadAllBytes(imagePath);

                // Add a rectangle shape that will receive the texture fill
                // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 300, 200); // width=300, height=200 (points)

                // Apply texture fill and assign the image data
                shape.Fill.FillType = FillType.Texture;
                shape.Fill.TextureFill.ImageData = imageData;

                // Retrieve the shape's size (in points)
                double shapeWidth = shape.Width;
                double shapeHeight = shape.Height;

                // Add the image to the worksheet temporarily to obtain its dimensions (in points)
                int picIdx;
                using (MemoryStream ms = new MemoryStream(imageData))
                {
                    picIdx = worksheet.Pictures.Add(0, 0, ms);
                }

                Picture tempPic = worksheet.Pictures[picIdx];
                double imageWidth = tempPic.Width;
                double imageHeight = tempPic.Height;

                // Remove the temporary picture as it is no longer needed
                worksheet.Pictures.RemoveAt(picIdx);

                // Determine whether tiling should be enabled
                // If the shape is larger than the image, enable tiling; otherwise, disable it
                bool enableTiling = shapeWidth > imageWidth || shapeHeight > imageHeight;

                // Set the IsTiling property accordingly
                shape.Fill.TextureFill.IsTiling = enableTiling;

                // Output the decision for verification
                Console.WriteLine($"Shape size: {shapeWidth}x{shapeHeight} (points)");
                Console.WriteLine($"Image size: {imageWidth}x{imageHeight} (points)");
                Console.WriteLine($"IsTiling set to: {enableTiling}");

                // Save the workbook
                string outputPath = "ShapeTilingCheckDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
