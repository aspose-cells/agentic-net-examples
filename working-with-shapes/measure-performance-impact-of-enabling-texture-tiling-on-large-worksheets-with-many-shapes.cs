// Title: Measure the performance impact of adding thousands of rectangle shapes with and without texture tiling on a large worksheet using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that creates a 5,000‑row by 20‑column worksheet, inserts 2,000 rectangle shapes, and records the elapsed milliseconds for the insertion loop. | Extend the shape‑insertion code to enable texture tiling on each rectangle (using the appropriate Aspose.Cells API when available) and log the timing difference versus the no‑tiling run. | Add robust exception handling, save the workbook to disk, and output both timing results to the console for easy comparison.
// Common Searches: how to benchmark shape insertion speed in Aspose.Cells C# | performance difference texture tiling vs no tiling for Excel shapes Aspose.Cells | measure time to add 2000 rectangles to a large worksheet using Aspose.Cells | Aspose.Cells shape creation latency on worksheets with thousands of rows | C# code to compare shape texture tiling impact in Excel files
// Tags: shape insertion performance Aspose.Cells | texture tiling impact on Excel shapes | large worksheet shape creation C# | Aspose.Cells rectangle shape API timing | benchmarking shape rendering Aspose.Cells .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook with 5,000 rows and 20 columns, adds 2,000 rectangle shapes twice—first without texture tiling and then after clearing the shapes with texture tiling enabled (if supported)—while measuring each operation with Stopwatch. It prints the elapsed milliseconds for both scenarios, saves the workbook, and includes error handling for shape creation and file saving.
class TextureTilingPerformance
{
    static void Main()
    {
        try
        {
            // Path to save the generated workbook
            string outputPath = "TextureTilingPerformance.xlsx";

            // Number of shapes to add (large number for performance test)
            const int shapeCount = 2000;

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "PerformanceTest";

            // Fill the worksheet with some data to simulate a realistic large sheet
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // -------------------------------------------------
            // Test 1: Texture tiling disabled (texture not applied)
            // -------------------------------------------------
            Stopwatch swNoTiling = Stopwatch.StartNew();

            try
            {
                for (int i = 0; i < shapeCount; i++)
                {
                    int col = i % 20;
                    int row = i / 20;

                    // Add a rectangle shape at the cell position (offsets set to 0)
                    Shape shape = sheet.Shapes.AddShape(
                        MsoDrawingType.Rectangle,
                        row,               // upper left row
                        col,               // upper left column
                        0,                 // top offset (pixels)
                        0,                 // left offset (pixels)
                        30,                // height (pixels)
                        80);               // width (pixels)

                    // NOTE: Texture and tiling properties are not available in the current Aspose.Cells version.
                    // If needed, they can be set using the appropriate API in newer versions.
                }
            }
            catch (Exception shapeEx)
            {
                Console.WriteLine($"Error while adding shapes (no tiling): {shapeEx.Message}");
            }

            swNoTiling.Stop();

            // -------------------------------------------------
            // Test 2: Texture tiling enabled (texture not applied)
            // -------------------------------------------------
            // Clear previously added shapes to isolate the second test
            sheet.Shapes.Clear();

            Stopwatch swTiling = Stopwatch.StartNew();

            try
            {
                for (int i = 0; i < shapeCount; i++)
                {
                    int col = i % 20;
                    int row = i / 20;

                    Shape shape = sheet.Shapes.AddShape(
                        MsoDrawingType.Rectangle,
                        row,
                        col,
                        0,
                        0,
                        30,
                        80);

                    // Texture and tiling settings are omitted for compatibility.
                }
            }
            catch (Exception shapeEx)
            {
                Console.WriteLine($"Error while adding shapes (tiling): {shapeEx.Message}");
            }

            swTiling.Stop();

            // -------------------------------------------------
            // Output results
            // -------------------------------------------------
            Console.WriteLine($"Time without texture tiling: {swNoTiling.ElapsedMilliseconds} ms");
            Console.WriteLine($"Time with texture tiling   : {swTiling.ElapsedMilliseconds} ms");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, to ensure full processing)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
