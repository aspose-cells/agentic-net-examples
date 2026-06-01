using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPerformanceDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                TextureTilingPerformance.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }
    }

    public class TextureTilingPerformance
    {
        public static void Run()
        {
            // Path to the texture image on the desktop
            string imagePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "texture.png");
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            byte[] textureData;
            try
            {
                textureData = File.ReadAllBytes(imagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read image file: {ex.Message}");
                return;
            }

            const int shapeCount = 2000;

            // -------------------------------------------------
            // Test 1: Texture fill without tiling (IsTiling = false)
            // -------------------------------------------------
            Workbook wbNoTile = new Workbook();
            Worksheet wsNoTile = wbNoTile.Worksheets[0];
            Stopwatch swNoTile = Stopwatch.StartNew();

            for (int i = 0; i < shapeCount; i++)
            {
                Shape shape = wsNoTile.Shapes.AddRectangle(0, 0, i % 100, i % 100, 50, 30);
                shape.Fill.FillType = FillType.Texture;
                shape.Fill.TextureFill.ImageData = textureData;
                shape.Fill.TextureFill.IsTiling = false;
            }

            swNoTile.Stop();
            Console.WriteLine($"Time without tiling: {swNoTile.ElapsedMilliseconds} ms");

            try
            {
                wbNoTile.Save("TextureNoTiling.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook (no tiling): {ex.Message}");
            }

            // -------------------------------------------------
            // Test 2: Texture fill with tiling enabled (IsTiling = true)
            // -------------------------------------------------
            Workbook wbTile = new Workbook();
            Worksheet wsTile = wbTile.Worksheets[0];
            Stopwatch swTile = Stopwatch.StartNew();

            for (int i = 0; i < shapeCount; i++)
            {
                Shape shape = wsTile.Shapes.AddRectangle(0, 0, i % 100, i % 100, 50, 30);
                shape.Fill.FillType = FillType.Texture;
                shape.Fill.TextureFill.ImageData = textureData;
                shape.Fill.TextureFill.IsTiling = true;
                shape.Fill.TextureFill.TilePicOption = new TilePicOption
                {
                    ScaleX = 0.5,
                    ScaleY = 0.5,
                    OffsetX = 5,
                    OffsetY = 5
                };
            }

            swTile.Stop();
            Console.WriteLine($"Time with tiling: {swTile.ElapsedMilliseconds} ms");

            try
            {
                wbTile.Save("TextureWithTiling.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook (tiling): {ex.Message}");
            }
        }
    }
}