// Title: Measuring the execution time of Shape.UpdateSelectedValue across thousands of linked shapes in a large Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create C# code that iterates through all shapes in a worksheet, calls the method that syncs each shape with its linked cell, and records the total duration with a Stopwatch. | Add try‑catch blocks around each shape update to capture exceptions and log the shape name together with the error message. | After processing, save the workbook to a new file and output the full path and the number of shapes processed. | Include a console summary that prints the elapsed milliseconds and the average time per shape.
// Common Searches: how to time Shape.UpdateSelectedValue calls in Aspose.Cells .NET | performance of linked shape refresh in large Excel files using Aspose | measure execution speed of shape value synchronization with Aspose.Cells | C# benchmark for updating thousands of worksheet shapes in Excel | Aspose.Cells shape update latency testing on big workbooks
// Tags: shape value synchronization timing Aspose.Cells | large workbook linked shape processing .NET | measure shape refresh latency Excel | optimize shape update throughput Aspose | profiling worksheet shape operations C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads a large Excel workbook, iterates over every shape on the first worksheet, calls UpdateSelectedValue for each shape while timing the operation with Stopwatch, logs any errors per shape, prints total shapes processed and elapsed milliseconds, then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Path to the input workbook.
        string inputPath = "LargeWorkbook.xlsx";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook.
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Work with the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Get all shapes on the worksheet.
            ShapeCollection shapes = worksheet.Shapes;

            // Measure time taken to update each shape.
            Stopwatch timer = Stopwatch.StartNew();

            foreach (Shape shape in shapes)
            {
                try
                {
                    // Update the shape's displayed value based on its linked cell.
                    shape.UpdateSelectedValue();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error updating shape '{shape.Name}': {ex.Message}");
                }
            }

            timer.Stop();

            Console.WriteLine($"UpdateSelectedValue called on {shapes.Count} shapes in {timer.ElapsedMilliseconds} ms.");

            // Save the updated workbook.
            string outputPath = "LargeWorkbook_Updated.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Processing error: {ex.Message}");
        }
    }
}
