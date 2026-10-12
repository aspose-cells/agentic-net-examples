// Title: Benchmark SmartArt shape detection time per worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates all worksheets, checks each shape's Type string for the word "SmartArt" (case‑insensitive), and records the detection time for each worksheet using Stopwatch. | Add a global Stopwatch to measure the total time taken to scan every worksheet for SmartArt shapes and output both per‑sheet and overall elapsed milliseconds. | Include error handling for missing input files and shape‑processing exceptions, and format console output to show worksheet name, SmartArt presence, and timing results.
// Common Searches: how to measure smartart detection performance with aspose.cells in c# | c# benchmark shape scanning time across multiple excel worksheets | using stopwatch to profile aspose.cells shape enumeration per sheet | detect smartart objects in workbook and log execution time .net | performance testing of smartart detection in large excel files using aspose
// Tags: smartart detection timing aspnet | worksheet shape enumeration performance | aspose.cells stopwatch profiling | excel smartart shape identification c# | total execution measurement workbook scanning

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, examines every shape, detects SmartArt by performing a case‑insensitive string search on the shape type, and uses Stopwatch to record detection time per worksheet and overall. It prints the worksheet name, whether SmartArt was found, per‑sheet elapsed milliseconds, and total elapsed time, while handling missing file and shape‑processing errors.
class SmartArtDetectionPerformance
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Stopwatch for total execution time
        Stopwatch totalStopwatch = Stopwatch.StartNew();

        try
        {
            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Stopwatch for per‑worksheet detection time
                Stopwatch wsStopwatch = Stopwatch.StartNew();

                bool containsSmartArt = false;

                try
                {
                    // Examine each shape in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Detect SmartArt by checking the shape type name (avoids direct enum dependency)
                        if (shape.Type.ToString().Contains("SmartArt", StringComparison.OrdinalIgnoreCase))
                        {
                            containsSmartArt = true;
                            break; // No need to continue once a SmartArt is found
                        }
                    }
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Error processing shapes in worksheet \"{sheet.Name}\": {exShape.Message}");
                }

                wsStopwatch.Stop();

                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine($"  SmartArt detected: {containsSmartArt}");
                Console.WriteLine($"  Detection time: {wsStopwatch.ElapsedMilliseconds} ms");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
        finally
        {
            totalStopwatch.Stop();
            Console.WriteLine($"Total detection time for all worksheets: {totalStopwatch.ElapsedMilliseconds} ms");
        }

        // Optional: save a copy of the workbook (e.g., after processing)
        // string outputPath = "output.xlsx";
        // workbook.Save(outputPath);
    }
}
