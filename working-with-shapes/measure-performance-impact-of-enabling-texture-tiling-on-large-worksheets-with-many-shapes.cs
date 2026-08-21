// Title: C# Benchmark: Texture Fill Tiling Performance on 2,000 Shapes with Aspose.Cells
// Description: A C# example that creates 2,000 rectangle shapes, applies a built‑in texture fill with and without tiling, measures execution time using Stopwatch, clears shapes between runs, and saves the workbook to evaluate the performance impact of texture tiling in Aspose.Cells.
// Keywords: Aspose.Cells texture tiling | C# shape performance benchmark | IsTiling execution time | large worksheet shape rendering | TilePicOption impact | Aspose.Cells shape fill speed
// Common Searches: Aspose.Cells texture tiling performance test | measure shape creation time with texture fill C# | benchmark IsTiling on many shapes | how fast is texture tiling in Aspose.Cells | performance of TilePicOption in .NET
// Developer Intent: Compare the runtime of adding textured shapes with tiling disabled versus enabled to determine the overhead introduced by texture tiling.
// Use Cases: Assess whether texture tiling is viable for reports containing thousands of shapes. | Identify performance bottlenecks when configuring TilePicOption properties. | Validate that saving a workbook with tiled textures does not cause unacceptable delays.
// AI Prompts: Create a parameterized method that runs the tiling benchmark for any shape count and returns both elapsed times. | Suggest code optimizations to minimize the cost of texture tiling when populating large worksheets. | Provide a snippet that logs memory consumption and CPU usage during the tiling and non‑tiling loops.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# example that creates 2,000 rectangle shapes, applies a built‑in texture fill with and without tiling, measures execution time using Stopwatch, clears shapes between runs, and saves the workbook to evaluate the performance impact of texture tiling in Aspose.Cells.
class TextureTilingPerformance
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        const int shapeCount = 2000; // Number of shapes to test with

        // -------------------------------------------------
        // Test 1: Apply texture fill without tiling
        // -------------------------------------------------
        Stopwatch swNoTile = Stopwatch.StartNew();

        for (int i = 0; i < shapeCount; i++)
        {
            // Add a rectangle shape
            Shape shape = sheet.Shapes.AddRectangle(0, 0, i, i, 100, 30);

            // Set texture fill (built‑in texture)
            shape.Fill.FillType = FillType.Texture;
            shape.Fill.TextureFill.Type = TextureType.BlueTissuePaper;

            // Disable tiling
            shape.Fill.TextureFill.IsTiling = false;
        }

        swNoTile.Stop();

        // Clear all shapes before the next test
        sheet.Shapes.Clear();

        // -------------------------------------------------
        // Test 2: Apply texture fill with tiling enabled
        // -------------------------------------------------
        Stopwatch swTile = Stopwatch.StartNew();

        for (int i = 0; i < shapeCount; i++)
        {
            // Add a rectangle shape
            Shape shape = sheet.Shapes.AddRectangle(0, 0, i, i, 100, 30);

            // Set texture fill (built‑in texture)
            shape.Fill.FillType = FillType.Texture;
            shape.Fill.TextureFill.Type = TextureType.BlueTissuePaper;

            // Enable tiling
            shape.Fill.TextureFill.IsTiling = true;

            // Optional: configure tile options (scale and offset)
            shape.Fill.TextureFill.TilePicOption = new TilePicOption
            {
                ScaleX = 0.5,
                ScaleY = 0.5,
                OffsetX = 5,
                OffsetY = 5
            };
        }

        swTile.Stop();

        // -------------------------------------------------
        // Output performance results
        // -------------------------------------------------
        Console.WriteLine($"Time without tiling: {swNoTile.ElapsedMilliseconds} ms");
        Console.WriteLine($"Time with tiling   : {swTile.ElapsedMilliseconds} ms");

        // Save the workbook (contains the tiled shapes)
        workbook.Save("TextureTilingPerformance.xlsx");
    }
}
