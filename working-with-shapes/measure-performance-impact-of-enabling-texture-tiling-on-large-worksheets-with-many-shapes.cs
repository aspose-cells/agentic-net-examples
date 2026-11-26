// Title: Benchmarking texture fill tiling performance when adding thousands of shapes to an Excel worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells benchmark texture fill tiling with 2000 shapes in C# | measure performance impact of TextureFill.IsTiling in large Excel worksheets | how to test shape insertion speed with texture tiling using Aspose.Cells | C# code to compare tiling enabled vs disabled for shape textures in Excel | performance test for adding rectangle shapes with texture fill in Aspose.Cells
// Tags: texture fill tiling performance Aspose.Cells | shape insertion benchmark C# Aspose.Cells | TextureFill.IsTiling impact Excel workbook | large worksheet shape rendering Aspose.Cells | grid layout shape addition Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TextureTilingPerformanceTest
{
    // The sample creates two workbooks, each with 2,000 rectangle shapes arranged in a grid. Each shape receives a texture fill from a 1x1 PNG, with the TextureFill.IsTiling property toggled on or off. Execution time for building and saving the workbook is measured, allowing developers to compare the performance impact of enabling texture tiling on large worksheets.
    class Program
    {
        // Number of shapes to add for the performance test.
        private const int ShapeCount = 2000;

        // Size of each shape (in pixels).
        private const int ShapeWidth = 100;
        private const int ShapeHeight = 60;

        static void Main(string[] args)
        {
            try
            {
                // Generate a simple texture image in memory (1x1 PNG).
                byte[] textureBytes = CreateTextureBytes();

                Console.WriteLine("Starting performance test with {0} shapes per workbook.", ShapeCount);
                Console.WriteLine();

                // Test without tiling.
                RunTest(textureBytes, tilingEnabled: false);

                // Test with tiling.
                RunTest(textureBytes, tilingEnabled: true);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex);
            }
        }

        /// <param name="textureData">Byte array containing the texture image.</param>
        /// <param name="tilingEnabled">Whether the TextureFill.IsTiling property should be set to true.</param>
        private static void RunTest(byte[] textureData, bool tilingEnabled)
        {
            string mode = tilingEnabled ? "TilingEnabled" : "TilingDisabled";
            string outputFile = $"TextureTiling_{mode}.xlsx";

            Stopwatch sw = new Stopwatch();
            sw.Start();

            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add shapes in a grid layout to avoid overlapping.
                int cols = (int)Math.Ceiling(Math.Sqrt(ShapeCount));
                int rows = (int)Math.Ceiling((double)ShapeCount / cols);
                int startRow = 0;
                int startColumn = 0;

                for (int i = 0; i < ShapeCount; i++)
                {
                    int row = startRow + (i / cols) * (ShapeHeight + 5);
                    int col = startColumn + (i % cols) * (ShapeWidth + 5);

                    // Add a rectangle shape.
                    Shape shape = sheet.Shapes.AddRectangle(row, col, 0, 0, ShapeWidth, ShapeHeight);

                    // Configure texture fill.
                    shape.Fill.FillType = FillType.Texture;
                    TextureFill texFill = shape.Fill.TextureFill;
                    texFill.ImageData = textureData;
                    texFill.IsTiling = tilingEnabled;

                    // Optional: set a small scale to make tiling effect visible.
                    texFill.Scale = 0.5;
                }

                // Ensure workbook is fully built.
                workbook.CalculateFormula();

                // Save the workbook.
                workbook.Save(outputFile, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during {mode} test: {ex.Message}");
                return;
            }
            finally
            {
                sw.Stop();
            }

            Console.WriteLine($"{mode}:");
            Console.WriteLine($"  Shapes added: {ShapeCount}");
            Console.WriteLine($"  Time elapsed: {sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"  Output file: {outputFile}");
            Console.WriteLine();
        }

        /// <returns>Byte array containing PNG image data.</returns>
        private static byte[] CreateTextureBytes()
        {
            // Base64-encoded 1x1 transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}
