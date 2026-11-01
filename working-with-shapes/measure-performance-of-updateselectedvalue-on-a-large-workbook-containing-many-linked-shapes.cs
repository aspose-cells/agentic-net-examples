// Title: Benchmark the UpdateSelectedValue method on linked shapes in a large Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook, iterates all worksheets and shapes, calls UpdateSelectedValue only on supported shapes, and records the total execution time with Stopwatch. | Write a performance test that measures how long UpdateSelectedValue takes for every linked shape in a multi‑sheet .xlsx file, handling exceptions for non‑linked shapes. | Create a utility that skips shapes that throw when UpdateSelectedValue is invoked, logs the number of processed shapes, and outputs the elapsed milliseconds.
// Common Searches: how to benchmark UpdateSelectedValue for linked shapes using Aspose.Cells .NET | measure time of shape value updates in a large Excel file C# Aspose.Cells | iterate all shapes in workbook and call UpdateSelectedValue with error handling | performance testing of linked shape updates in Aspose.Cells large workbook
// Tags: Aspose.Cells UpdateSelectedValue benchmark | linked shape value update performance .NET | measure shape processing time Aspose.Cells | large workbook shape iteration Aspose.Cells | handle unsupported shapes Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a large Excel workbook, walks through each worksheet and its shapes, attempts to invoke UpdateSelectedValue on linked shapes while safely ignoring unsupported ones, measures the total elapsed time with a Stopwatch, prints the duration, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the large workbook (load rule)
        Workbook workbook = new Workbook("LargeWorkbook.xlsx");

        // Prepare a stopwatch to measure performance
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Attempt to update the selected value for linked shapes
                // (UpdateSelectedValue is applicable only to linked shapes;
                //  unsupported shapes will throw, so we catch and ignore)
                try
                {
                    shape.UpdateSelectedValue();
                }
                catch
                {
                    // Ignore shapes that do not support UpdateSelectedValue
                }
            }
        }

        stopwatch.Stop();

        // Output the elapsed time
        Console.WriteLine($"UpdateSelectedValue executed in {stopwatch.ElapsedMilliseconds} ms.");

        // Save the workbook after updates (save rule)
        workbook.Save("LargeWorkbook_Updated.xlsx");
    }
}
