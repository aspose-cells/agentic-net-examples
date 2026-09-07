// Title: Add an exponential trendline to the first series of an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx workbook, locate the first chart, and programmatically insert an exponential trendline into its first data series with Aspose.Cells. | When the Trendlines property is not exposed, use .NET reflection to invoke Trendlines.Add passing TrendlineType.Exponential. | Assign a custom name to the created trendline and save the modified workbook to a new file.
// Common Searches: how to add an exponential trendline to a chart series with Aspose.Cells C# | Aspose.Cells .NET reflection example for adding trendlines | set trendline type to exponential programmatically in Excel using Aspose.Cells | C# code to rename a trendline added to an Aspose.Cells chart | check for chart and series before adding a trendline Aspose.Cells
// Tags: exponential trendline insertion Aspose.Cells | reflection based trendline addition .NET | custom trendline naming Aspose.Cells | chart series manipulation Aspose.Cells | fallback handling for missing Trendlines API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, verifies a chart and its first series, uses reflection to add an exponential trendline with a custom name when the API is unavailable, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found on the first worksheet.");
                    return;
                }

                // Get the first chart on the worksheet
                Chart chart = sheet.Charts[0];

                // Ensure the chart has at least one data series
                if (chart.NSeries.Count == 0)
                {
                    Console.WriteLine("Error: The chart does not contain any data series.");
                    return;
                }

                // Attempt to add an exponential trendline using reflection (covers versions without direct API)
                try
                {
                    Series series = chart.NSeries[0];
                    var trendlinesProp = series.GetType().GetProperty("Trendlines");
                    if (trendlinesProp != null)
                    {
                        object trendlines = trendlinesProp.GetValue(series);
                        var addMethod = trendlines.GetType().GetMethod("Add", new[] { typeof(TrendlineType) });
                        if (addMethod != null)
                        {
                            object trendline = addMethod.Invoke(trendlines, new object[] { TrendlineType.Exponential });
                            var nameProp = trendline.GetType().GetProperty("Name");
                            nameProp?.SetValue(trendline, "Exponential Trendline");
                            Console.WriteLine("Exponential trendline added successfully.");
                        }
                        else
                        {
                            Console.WriteLine("Add method for trendlines not found.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Trendlines property not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add trendline via reflection: {ex.Message}");
                }

                // Save the workbook with the (potential) new trendline
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
