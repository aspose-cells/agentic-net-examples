// Title: Benchmark SmartArt Shape Detection Across Worksheets with Aspose.Cells for .NET
// Description: Loads an Excel workbook, starts a Stopwatch, iterates every worksheet and its shapes, counts shapes where IsSmartArt is true, stops the timer, prints the SmartArt total and elapsed milliseconds, then saves the file.
// Keywords: Aspose.Cells | SmartArt detection | performance benchmark | C# Stopwatch | shape enumeration | Excel workbook | .NET | count SmartArt shapes | execution time measurement | large workbook optimization
// Common Searches: Aspose.Cells measure time to find SmartArt | C# benchmark shape iteration in Excel | How long does SmartArt detection take with Aspose.Cells | Performance test for SmartArt enumeration .NET | Count SmartArt objects in workbook using Aspose.Cells
// Developer Intent: The developer wants to time how long it takes to detect and count SmartArt shapes across all worksheets in a workbook.
// Use Cases: Identify performance bottlenecks when processing workbooks that contain many SmartArt objects. | Compare detection speed before and after optimizing shape traversal logic. | Log execution time for SmartArt detection as part of an automated Excel processing pipeline. | Include timing metrics in CI/CD reports to ensure scalability. | Generate performance dashboards for large‑scale Excel automation.
// AI Prompts: Suggest code changes to reduce SmartArt detection time in large workbooks using Aspose.Cells. | Show how to log Stopwatch results to a file with robust exception handling for missing input files. | Explain a safe way to parallelize worksheet shape iteration while accurately counting SmartArt objects. | Provide an example of sending execution metrics to Azure Application Insights or another monitoring service. | Recommend memory‑efficient techniques for enumerating shapes in very large Excel files.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, starts a Stopwatch, iterates every worksheet and its shapes, counts shapes where IsSmartArt is true, stops the timer, prints the SmartArt total and elapsed milliseconds, then saves the file.
class SmartArtDetectionPerformance
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Start measuring time
        Stopwatch stopwatch = Stopwatch.StartNew();

        int smartArtCount = 0;

        // Iterate through all worksheets and their shapes
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            foreach (Shape shape in worksheet.Shapes)
            {
                // Detect if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    smartArtCount++;
                }
            }
        }

        // Stop the timer
        stopwatch.Stop();

        // Output results
        Console.WriteLine($"Total SmartArt shapes found: {smartArtCount}");
        Console.WriteLine($"Execution time: {stopwatch.ElapsedMilliseconds} ms");

        // Save the workbook (no modifications made, just to follow lifecycle)
        workbook.Save("output.xlsx");
    }
}
