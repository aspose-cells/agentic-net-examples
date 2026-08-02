// Title: Benchmark Texture Tiling Performance for Thousands of Shapes in Aspose.Cells (.NET)
// Description: C# sample that loads a texture image once, creates a new workbook, and adds 1,000 rectangle shapes in a grid on the first worksheet. Each shape uses a texture fill with the IsTiling flag toggled to compare tiled vs. non‑tiled scenarios. The code measures elapsed milliseconds, saves two workbooks (tiling enabled/disabled), and provides a clear performance baseline for shape‑heavy Excel files using Aspose.Cells for .NET.
// Keywords: Aspose.Cells texture fill benchmark | C# texture tiling performance | shape performance Aspose.Cells | large worksheet shape test | IsTiling property speed | Aspose.Cells .NET performance | Excel texture fill tiling | measure shape creation time | Aspose.Cells shape grid | texture fill tiling comparison
// Common Searches: Aspose.Cells texture tiling benchmark | measure performance of texture fill in Excel using C# | how fast is shape creation with tiling in Aspose.Cells | compare tiled vs non‑tiled texture fill speed | performance test for thousands of shapes Aspose.Cells
// Developer Intent: Evaluate the impact of enabling the IsTiling property on execution time when adding a large number of textured shapes to an Excel worksheet with Aspose.Cells.
// Use Cases: Determine whether texture tiling meets performance requirements for high‑volume reporting. | Create a baseline to decide if tiling should be used in dashboard or visualization workbooks. | Validate that adding many textured shapes does not exceed acceptable generation time before production deployment.
// AI Prompts: Generate a memory‑optimized version of this texture tiling benchmark for .NET. | Show how to parallelize shape creation in Aspose.Cells while keeping thread safety. | Explain how to interpret the timing results and suggest performance thresholds for enabling texture tiling in production spreadsheets.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TextureTilingPerformanceDemo
{
    // C# sample that loads a texture image once, creates a new workbook, and adds 1,000 rectangle shapes in a grid on the first worksheet. Each shape uses a texture fill with the IsTiling flag toggled to compare tiled vs. non‑tiled scenarios. The code measures elapsed milliseconds, saves two workbooks (tiling enabled/disabled), and provides a clear performance baseline for shape‑heavy Excel files using Aspose.Cells for .NET.
    class Program
    {
        // Path to a sample texture image (replace with an actual file path on your system)
        private const string TextureImagePath = @"C:\Images\texture.png";

        // Number of shapes to add for the performance test
        private const int ShapeCount = 1000;

        static void Main()
        {
            try
            {
                // Verify that the texture image exists
                if (!File.Exists(TextureImagePath))
                {
                    Console.WriteLine($"Error: Texture image not found at '{TextureImagePath}'.");
                    return;
                }

                // Load image data once to reuse for all shapes
                byte[] textureData = File.ReadAllBytes(TextureImagePath);

                // Measure performance without tiling
                long timeWithoutTiling = MeasureTextureTilingPerformance(textureData, isTiling: false);
                Console.WriteLine($"Time without tiling: {timeWithoutTiling} ms");

                // Measure performance with tiling enabled
                long timeWithTiling = MeasureTextureTilingPerformance(textureData, isTiling: true);
                Console.WriteLine($"Time with tiling: {timeWithTiling} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <param name="imageData">Byte array of the texture image.</param>
        /// <param name="isTiling">Whether to enable texture tiling.</param>
        /// <returns>Elapsed time in milliseconds.</returns>
        private static long MeasureTextureTilingPerformance(byte[] imageData, bool isTiling)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Start timing
                Stopwatch sw = Stopwatch.StartNew();

                // Add many rectangle shapes
                for (int i = 0; i < ShapeCount; i++)
                {
                    // Position each shape in a grid layout
                    int row = i / 50;          // 50 shapes per row
                    int col = i % 50;          // 50 shapes per column
                    int upperRow = row * 5;    // each shape occupies 5 rows
                    int leftCol = col * 5;     // each shape occupies 5 columns
                    int lowerRow = upperRow + 4;
                    int rightCol = leftCol + 4;

                    // Height and width of the shape (in points). Adjust as needed.
                    const int shapeHeight = 100;
                    const int shapeWidth = 100;

                    // Add rectangle shape with required parameters
                    Shape rect = sheet.Shapes.AddRectangle(upperRow, leftCol, lowerRow, rightCol, shapeHeight, shapeWidth);

                    // Set fill type to texture
                    rect.Fill.FillType = FillType.Texture;

                    // Configure texture fill
                    TextureFill texFill = rect.Fill.TextureFill;
                    texFill.ImageData = imageData;   // use the same image for all shapes
                    texFill.IsTiling = isTiling;    // enable or disable tiling

                    // Optional: configure tile options for better visual effect
                    if (isTiling)
                    {
                        TilePicOption tileOption = new TilePicOption
                        {
                            ScaleX = 0.5,   // 50% horizontal scaling
                            ScaleY = 0.5,   // 50% vertical scaling
                            OffsetX = 5,
                            OffsetY = 5
                        };
                        texFill.TilePicOption = tileOption;
                    }
                }

                // Stop timing
                sw.Stop();

                // Save the workbook to verify the result (lifecycle rule: save)
                string fileName = isTiling ? "Workbook_TilingEnabled.xlsx" : "Workbook_TilingDisabled.xlsx";
                workbook.Save(fileName);

                // Return elapsed milliseconds
                return sw.ElapsedMilliseconds;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during performance measurement: {ex.Message}");
                return -1;
            }
        }
    }
}
