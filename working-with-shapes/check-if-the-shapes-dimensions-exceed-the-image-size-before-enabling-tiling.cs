// Title: Check shape dimensions against texture image size before enabling tiling in Aspose.Cells (C#)
// Common Searches: how to prevent texture stretching by verifying shape size versus image size in Aspose.Cells C# | Aspose.Cells C# enable texture fill tiling only when shape fits the source image | determine if a rectangle shape can be tiled with a PNG texture using Aspose.Cells | C# compare shape dimensions with image dimensions before setting IsTiling in Aspose.Cells | using reflection to read PNG dimensions for texture fill tiling decision in Aspose.Cells
// Tags: shape dimension validation for texture fill | texture tiling based on image size | pixel-to-point conversion for shape comparison | reflection-based image size retrieval in C# | Aspose.Cells rectangle texture fill example

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeTextureTilingCheck
{
    // The sample loads a PNG via reflection to obtain its pixel dimensions, converts them to points, creates a rectangle shape in a new workbook, assigns the image as a texture fill, and enables tiling only when the shape's width and height are less than or equal to the image's point dimensions, then saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // Configuration – adjust these paths as needed.
            // -----------------------------------------------------------------
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "texture.png");
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ShapeTextureTilingDemo.xlsx");

            // Validate image existence.
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file not found at '{imagePath}'.");
                return;
            }

            try
            {
                // Load image dimensions using reflection to avoid a direct System.Drawing dependency.
                int imgPixelWidth, imgPixelHeight;
                Assembly drawingAssembly = Assembly.Load("System.Drawing.Common");
                Type imageType = drawingAssembly.GetType("System.Drawing.Image");
                MethodInfo fromFile = imageType.GetMethod("FromFile", new[] { typeof(string) });
                PropertyInfo widthProp = imageType.GetProperty("Width");
                PropertyInfo heightProp = imageType.GetProperty("Height");

                using (IDisposable img = (IDisposable)fromFile.Invoke(null, new object[] { imagePath }))
                {
                    imgPixelWidth = (int)widthProp.GetValue(img);
                    imgPixelHeight = (int)heightProp.GetValue(img);
                }

                // Convert pixel dimensions to points (1 pixel = 0.75 point).
                const double pixelToPoint = 0.75;
                double imgPointWidth = imgPixelWidth * pixelToPoint;
                double imgPointHeight = imgPixelHeight * pixelToPoint;

                Console.WriteLine($"Image dimensions: {imgPixelWidth}x{imgPixelHeight} pixels ({imgPointWidth:F2}x{imgPointHeight:F2} points).");

                // -----------------------------------------------------------------
                // Create a new workbook and add a shape.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define shape position and size (in points). Adjust as needed.
                int upperRow = 2;
                int upperColumn = 1;
                int upperRowOffset = 0;
                int upperColumnOffset = 0;
                int shapeHeight = 200; // points (int required by API)
                int shapeWidth = 300;  // points (int required by API)

                Shape shape = sheet.Shapes.AddRectangle(upperRow, upperColumn, upperRowOffset, upperColumnOffset, shapeHeight, shapeWidth);
                shape.Name = "TextureRectangle";

                // Set the shape's fill type to texture.
                shape.Fill.FillType = FillType.Texture;
                TextureFill textureFill = shape.Fill.TextureFill;

                // Load image bytes and assign to the texture fill.
                byte[] imageBytes = File.ReadAllBytes(imagePath);
                textureFill.ImageData = imageBytes;

                // -----------------------------------------------------------------
                // Determine whether tiling should be enabled.
                // -----------------------------------------------------------------
                bool canTile = shape.Height <= imgPointHeight && shape.Width <= imgPointWidth;

                if (canTile)
                {
                    textureFill.IsTiling = true;
                    Console.WriteLine("Shape fits within the image. Tiling enabled.");
                }
                else
                {
                    textureFill.IsTiling = false;
                    Console.WriteLine("Shape exceeds image dimensions. Tiling disabled to prevent stretching.");
                    Console.WriteLine($"Shape size: {shape.Width:F2}x{shape.Height:F2} points.");
                }

                // Optional: configure additional texture options.
                textureFill.Scale = 1.0;        // No scaling.
                textureFill.Transparency = 0.0; // Fully opaque.

                // Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
