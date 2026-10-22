// Title: Copy primary chart axis scaling (min, max, major unit) to the secondary axis using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, accesses the first chart, and sets the secondary value axis MinValue, MaxValue, and MajorUnit to match the primary value axis. | Show how to use reflection in C# to obtain the SecondaryValueAxis property of an Aspose.Cells chart when it is not directly exposed, then synchronize its scaling with the primary axis. | Provide a complete example that saves the workbook after updating the secondary axis limits, including error handling for missing files or charts.
// Common Searches: Aspose.Cells C# copy primary axis limits to secondary axis in a chart | How to set secondary value axis minimum and maximum in Aspose.Cells chart using C# | C# reflection to access SecondaryValueAxis property Aspose.Cells | Synchronize major unit of secondary axis with primary axis Aspose.Cells chart | Update secondary axis scaling in existing Excel file with Aspose.Cells
// Tags: Aspose.Cells set secondary axis scaling C# | apply primary axis scaling to secondary axis Aspose.Cells | chart secondary value axis reflection Aspose.Cells | synchronize chart axis limits C# | modify Excel chart axis range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, retrieves the first chart, uses reflection to get the secondary value axis (if present), copies the MinValue, MaxValue, and MajorUnit from the primary axis, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the sheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                    return;
                }

                // Assume the chart to modify is the first chart on the sheet
                var chart = worksheet.Charts[0];
                var primaryAxis = chart.ValueAxis;

                // Example: set explicit scaling values if needed
                // primaryAxis.MinValue = 0;
                // primaryAxis.MaxValue = 100;
                // primaryAxis.MajorUnit = 10;

                // Attempt to handle secondary axis via reflection (may not be supported)
                try
                {
                    var secondaryProp = chart.GetType().GetProperty("SecondaryValueAxis");
                    if (secondaryProp != null)
                    {
                        var secondaryAxis = secondaryProp.GetValue(chart) as Axis;
                        if (secondaryAxis != null)
                        {
                            // Copy scaling settings from primary axis
                            secondaryAxis.MinValue = primaryAxis.MinValue;
                            secondaryAxis.MaxValue = primaryAxis.MaxValue;
                            secondaryAxis.MajorUnit = primaryAxis.MajorUnit;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Secondary axis handling skipped: {ex.Message}");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
