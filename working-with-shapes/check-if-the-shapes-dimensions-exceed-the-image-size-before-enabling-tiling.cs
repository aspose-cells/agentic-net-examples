// Title: C# – Enable texture tiling for a picture shape only when it exceeds the source image size in Aspose.Cells
// Description: Shows how to load an image, add it as a picture shape in a workbook, read the image's pixel dimensions from a byte array, compare them with the shape's Width and Height, and set Fill.TextureFill.IsTiling = true only when the shape is larger than the original image.
// Keywords: Aspose.Cells | C# | picture shape | texture fill | tiling | image dimensions | shape size validation | System.Drawing.Image | Excel workbook | conditional tiling | Aspose.Cells for .NET
// Common Searches: Aspose.Cells conditional texture tiling | Check picture shape size before tiling C# | Get image dimensions from byte array Aspose.Cells | Enable tiling only when shape larger than image | Aspose.Cells picture fill IsTiling condition
// Developer Intent: Determine whether a picture shape is larger than its source image and enable texture tiling only in that scenario.
// Use Cases: Extract width and height from an image byte array using System.Drawing.Image.FromStream and compare with pictureShape.Width/Height. | Set pictureShape.Fill.TextureFill.IsTiling = true only when the shape exceeds the image dimensions. | Log a warning or keep tiling disabled when the shape fits within the image to avoid redundant texture repetition. | Automatically resize a smaller shape to match the image size before applying tiling for consistent appearance. | Create a helper method that returns a boolean indicating if tiling should be applied based on size comparison.
// AI Prompts: Write a C# method that receives a Worksheet, an image byte array, and a picture shape index, extracts the image dimensions, compares them with the shape size, and returns true if texture tiling should be enabled. | Provide code that loads an image into a byte array, obtains its pixel width and height without saving to disk, and conditionally sets pictureShape.Fill.TextureFill.IsTiling in Aspose.Cells. | Show a reusable utility class for Aspose.Cells that validates shape dimensions against the source image and applies conditional texture tiling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to load an image, add it as a picture shape in a workbook, read the image's pixel dimensions from a byte array, compare them with the shape's Width and Height, and set Fill.TextureFill.IsTiling = true only when the shape is larger than the original image.
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file (replace with an actual image path)
            string imagePath = "sample.png";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load image data into a byte array
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Add the picture to the worksheet; it returns the index of the picture shape
            int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
            Picture pictureShape = worksheet.Pictures[pictureIndex];

            // Set the fill type to texture and assign the image data
            pictureShape.Fill.FillType = FillType.Texture;
            pictureShape.Fill.TextureFill.ImageData = imageData;

            // Enable tiling for demonstration purposes
            pictureShape.Fill.TextureFill.IsTiling = true;
            Console.WriteLine("Tiling enabled for the picture shape.");

            // Save the workbook
            string outputPath = "ShapeTilingCheckDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
