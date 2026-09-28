// Title: Measure how Workbook.RemoveUnusedStyles scales with increasing numbers of unused styles in C# using Aspose.Cells
// AI Prompts: Generate a C# console application that creates a workbook, adds a configurable number of unused styles, applies a few used styles, and records the elapsed milliseconds of Workbook.RemoveUnusedStyles(). | Write a script that iterates over style counts (e.g., 1000, 5000, 10000, 20000), builds a workbook for each count, invokes RemoveUnusedStyles, and prints the performance results to the console. | Create a performance benchmark that varies the number of styles in an Aspose.Cells workbook, calls Workbook.RemoveUnusedStyles, and logs the timing data for each iteration.
// Common Searches: c# benchmark Aspose.Cells RemoveUnusedStyles performance with thousands of styles | how does RemoveUnusedStyles execution time grow with workbook size in Aspose.Cells | measure time to clean up unused styles in large Excel workbook using Aspose.Cells C# | performance testing of Aspose.Cells style removal for 10k and 20k styles
// Tags: Aspose.Cells remove unused styles performance benchmark | C# workbook style cleanup timing | measure RemoveUnusedStyles scaling behavior | large workbook unused style removal Aspose.Cells | performance test for Aspose.Cells style count

using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Cells;

// The example creates workbooks with 1,000, 5,000, 10,000 and 20,000 unused styles, applies a few default styles to a 10x10 cell range, then measures and prints the milliseconds required for Workbook.RemoveUnusedStyles() for each style count, illustrating the scaling behavior of the cleanup operation.
class RemoveUnusedStylesBenchmark
{
    static void Main()
    {
        try
        {
            // Different numbers of styles to add (simulating varying workbook sizes)
            int[] styleCounts = new int[] { 1000, 5000, 10000, 20000 };

            foreach (int count in styleCounts)
            {
                try
                {
                    // Create a new workbook
                    Workbook workbook = new Workbook();

                    // Add a large number of unused styles to the workbook
                    for (int i = 0; i < count; i++)
                    {
                        // Create a new style based on the default style
                        Style newStyle = workbook.CreateStyle();
                        newStyle.Font.Color = (i % 2 == 0) ? Color.Red : Color.Blue;
                        newStyle.Font.Size = 10 + (i % 5);
                        // No need to explicitly add to the collection; CreateStyle registers it.
                    }

                    // Apply a few styles to cells so that some styles are used
                    Worksheet sheet = workbook.Worksheets[0];
                    for (int row = 0; row < 10; row++)
                    {
                        for (int col = 0; col < 10; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            cell.PutValue($"R{row}C{col}");
                            // Apply the default style for simplicity
                            cell.SetStyle(workbook.DefaultStyle);
                        }
                    }

                    // Measure the time required to remove unused styles
                    Stopwatch sw = Stopwatch.StartNew();
                    workbook.RemoveUnusedStyles();
                    sw.Stop();

                    // Output the result
                    Console.WriteLine($"Styles added: {count}, Time to remove unused styles: {sw.ElapsedMilliseconds} ms");
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Error processing count {count}: {innerEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
