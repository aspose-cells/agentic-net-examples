// Title: Benchmark SmartArt Detection Across Worksheets with Aspose.Cells for .NET
// Description: Loads an Excel file, uses a Stopwatch to time the enumeration of every worksheet and its Shapes collection, counts shapes where IsSmartArt is true, outputs the SmartArt total and elapsed milliseconds, and saves the workbook unchanged. Ideal for measuring and optimizing SmartArt detection performance.
// Keywords: Aspose.Cells | C# | .NET | SmartArt detection | shape enumeration | performance benchmark | execution time measurement | Stopwatch | Excel workbook processing | count SmartArt shapes
// Common Searches: how to benchmark SmartArt detection with Aspose.Cells | measure execution time of shape enumeration in C# | count SmartArt objects in Excel using Aspose.Cells | performance test for SmartArt detection .NET | timing shape.IsSmartArt loop Aspose
// Developer Intent: Find out how long it takes to locate and count SmartArt objects across all worksheets in an Excel workbook using Aspose.Cells.
// Use Cases: Evaluate the latency of SmartArt detection before adding further processing in a server‑side workflow. | Compare detection speed between small and large workbooks to decide if optimization is required. | Verify that SmartArt enumeration meets a specific performance SLA in an automated Excel pipeline.
// AI Prompts: Generate a C# method that returns both the SmartArt count and elapsed milliseconds for a given workbook using Aspose.Cells. | Suggest performance improvements for the SmartArt detection loop, such as parallel processing, early exit, or shape type filtering. | Create unit tests that assert the SmartArt count is correct and that the detection time stays under a defined threshold for sample files.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file, uses a Stopwatch to time the enumeration of every worksheet and its Shapes collection, counts shapes where IsSmartArt is true, outputs the SmartArt total and elapsed milliseconds, and saves the workbook unchanged. Ideal for measuring and optimizing SmartArt detection performance.
class SmartArtDetectionPerformance
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Start measuring time
        Stopwatch timer = Stopwatch.StartNew();

        int smartArtCount = 0;

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    smartArtCount++;

                    // Optional: convert SmartArt to grouped shapes if further processing is needed
                    // GroupShape group = shape.GetResultOfSmartArt();
                }
            }
        }

        // Stop the timer
        timer.Stop();

        // Output the results
        Console.WriteLine($"SmartArt shapes detected: {smartArtCount}");
        Console.WriteLine($"Execution time: {timer.ElapsedMilliseconds} ms");

        // Save the workbook (unchanged) to demonstrate lifecycle compliance
        workbook.Save("output.xlsx");
    }
}
