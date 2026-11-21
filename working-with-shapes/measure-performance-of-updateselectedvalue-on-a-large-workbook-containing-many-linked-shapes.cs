// Title: Measuring execution time of Shape.UpdateSelectedValue for linked shapes in a large Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to profile Shape.UpdateSelectedValue execution time with Aspose.Cells in C# | performance testing of linked shape updates in a big Excel file using Aspose.Cells | benchmark Shape.UpdateSelectedValue across thousands of shapes in a workbook .NET | C# example to time shape linked cell value synchronization with Aspose.Cells
// Tags: Shape.UpdateSelectedValue execution timing | Aspose.Cells linked shape performance | large workbook shape processing .NET | benchmarking Aspose.Cells shape operations | C# profiling of shape value updates

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads a large Excel workbook with Aspose.Cells, retrieves the first worksheet's ShapeCollection, and iterates through each shape calling UpdateSelectedValue inside a try‑catch to skip non‑linked shapes. A Stopwatch records the total elapsed time, which is printed to the console, and the workbook is saved to a new file.
class UpdateSelectedValuePerformance
{
    static void Main()
    {
        try
        {
            const string inputPath = "LargeWorkbook.xlsx";
            const string outputPath = "LargeWorkbook_Updated.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the large workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of shapes in the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Prepare a stopwatch to measure performance
            Stopwatch sw = new Stopwatch();

            // Start timing
            sw.Start();

            // Iterate through all shapes and update selected values for linked shapes
            foreach (Shape shape in shapes)
            {
                try
                {
                    // UpdateSelectedValue writes the shape's value back to the linked cell.
                    // If the shape is not linked to a cell, an exception may be thrown,
                    // which we safely ignore.
                    shape.UpdateSelectedValue();
                }
                catch
                {
                    // Ignored: shape is not linked to a cell or operation not supported.
                }
            }

            // Stop timing
            sw.Stop();

            // Output the elapsed time
            Console.WriteLine($"UpdateSelectedValue executed on {shapes.Count} shapes in {sw.Elapsed.TotalSeconds} seconds.");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, replace with desired output path)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
