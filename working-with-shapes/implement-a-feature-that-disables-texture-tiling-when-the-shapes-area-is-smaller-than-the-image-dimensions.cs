using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingDemo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                DisableTilingIfShapeSmaller.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    public class DisableTilingIfShapeSmaller
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (example dimensions)
                // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);

                // Set the fill type to texture
                shape.Fill.FillType = FillType.Texture;

                // Load texture image data (replace with your actual image path)
                string imagePath = "texture.png";

                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Texture image not found: {imagePath}");

                byte[] imageData = File.ReadAllBytes(imagePath);

                // Assign image data to the texture fill
                TextureFill textureFill = shape.Fill.TextureFill;
                textureFill.ImageData = imageData;

                // Determine image dimensions (PNG only) without System.Drawing
                GetImageDimensions(imageData, out int imageWidth, out int imageHeight);

                // Determine shape dimensions (Aspose.Cells uses points; convert to pixels assuming 96 DPI)
                // 1 point = 1/72 inch, 1 inch = 96 pixels => 1 point ≈ 1.3333 pixels
                const double pointsToPixels = 96.0 / 72.0;
                double shapeWidthPx = shape.Width * pointsToPixels;
                double shapeHeightPx = shape.Height * pointsToPixels;

                // Compare areas
                double shapeArea = shapeWidthPx * shapeHeightPx;
                double imageArea = imageWidth * imageHeight;

                // Disable tiling if shape area is smaller than image area
                textureFill.IsTiling = shapeArea >= imageArea;

                // Save the workbook (lifecycle rule: save)
                string outputPath = "TextureTilingAdjusted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in Run(): {ex.Message}");
                throw;
            }
        }

        // Simple PNG header parser to obtain width and height.
        private static void GetImageDimensions(byte[] imageData, out int width, out int height)
        {
            width = height = 0;
            try
            {
                // PNG signature is 8 bytes; IHDR chunk starts at offset 8
                if (imageData.Length < 24)
                    throw new InvalidDataException("Image data is too short to contain PNG header.");

                // Verify PNG signature
                byte[] pngSignature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
                for (int i = 0; i < pngSignature.Length; i++)
                {
                    if (imageData[i] != pngSignature[i])
                        throw new InvalidDataException("Image is not a valid PNG file.");
                }

                // IHDR chunk length (4 bytes) + chunk type (4 bytes) = offset 8
                // Width and height are 4-byte big‑endian integers starting at offset 16
                width = ReadInt32BigEndian(imageData, 16);
                height = ReadInt32BigEndian(imageData, 20);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read image dimensions: {ex.Message}");
                // Fallback to zero dimensions which will force tiling to be disabled
                width = height = 0;
            }
        }

        private static int ReadInt32BigEndian(byte[] data, int startIndex)
        {
            return (data[startIndex] << 24) |
                   (data[startIndex + 1] << 16) |
                   (data[startIndex + 2] << 8) |
                   data[startIndex + 3];
        }
    }
}