using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtDetectionTiming
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Start timing the SmartArt detection across all worksheets
            Stopwatch sw = Stopwatch.StartNew();

            int totalSmartArtCount = 0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check if the shape is a SmartArt object
                    if (shape.IsSmartArt)
                    {
                        totalSmartArtCount++;

                        // Optional: Convert SmartArt to grouped shapes if further processing is needed
                        // GroupShape group = shape.GetResultOfSmartArt();
                        // Perform additional operations on 'group' here
                    }
                }
            }

            // Stop the timer
            sw.Stop();

            // Output the results
            Console.WriteLine($"Total SmartArt shapes found: {totalSmartArtCount}");
            Console.WriteLine($"SmartArt detection time: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (no changes made, just to follow the lifecycle)
            workbook.Save("output.xlsx");
        }
    }
}