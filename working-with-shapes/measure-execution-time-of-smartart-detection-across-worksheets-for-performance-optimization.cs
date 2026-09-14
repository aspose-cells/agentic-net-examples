// Title: Benchmark SmartArt shape detection time across all worksheets with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel workbook, iterates through every worksheet and its shapes, counts shapes where IsSmartArt is true, and records the elapsed time with System.Diagnostics.Stopwatch. | Add file‑existence verification, robust workbook‑loading exception handling, and console logging of both the total SmartArt count and the detection duration in milliseconds to the SmartArt counting sample.
// Common Searches: how to benchmark smartart shape enumeration in a large Excel workbook using Aspose.Cells C# | c# measure execution time of shape.IsSmartArt check across multiple worksheets | performance testing for Aspose.Cells shape iteration and smartart counting
// Tags: smartart detection performance Aspose.Cells | measure shape iteration time .xlsx C# | benchmark Aspose.Cells smartart counting | stopwatch timing worksheet shapes Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file, validates its presence, iterates through each worksheet and every shape, increments a counter for shapes with the IsSmartArt property, and outputs both the total SmartArt count and the detection time measured in milliseconds using Stopwatch, with comprehensive error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Start measuring execution time
        Stopwatch stopwatch = Stopwatch.StartNew();

        int smartArtCount = 0;

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Detect SmartArt objects. In Aspose.Cells, SmartArt shapes expose the IsSmartArt property.
                    if (shape.IsSmartArt)
                    {
                        smartArtCount++;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing shapes: {ex.Message}");
            return;
        }

        // Stop the timer
        stopwatch.Stop();

        // Output results
        Console.WriteLine($"Total SmartArt objects found: {smartArtCount}");
        Console.WriteLine($"Detection time: {stopwatch.ElapsedMilliseconds} ms");
    }
}
