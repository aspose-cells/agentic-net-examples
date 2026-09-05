// Title: Add a linear trendline to the first series of an Excel chart with Aspose.Cells for .NET using reflection fallback
// AI Prompts: Write a C# program that opens an existing .xlsx file with Aspose.Cells, verifies that a chart and at least one series exist, and uses reflection to add a linear trendline to the first series of the first chart, then prints a confirmation to the console. | Generate C# code that loads a workbook, checks for chart presence, accesses the Series.Trendlines property via reflection for compatibility across Aspose.Cells versions, adds a linear trendline, and includes robust error handling.
// Common Searches: Aspose.Cells C# add linear trendline to chart series when Trendlines property not available | how to use reflection to add chart trendline in Aspose.Cells .NET | C# example loading Excel workbook and programmatically adding trendline with Aspose.Cells
// Tags: add linear trendline Aspose.Cells .NET | reflection access chart series Trendlines Aspose.Cells | load workbook modify chart Aspose.Cells | handle missing Trendlines API version Aspose.Cells | programmatic chart series manipulation Excel C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, ensures a chart and series are present, then uses reflection to access the series' Trendlines property and adds a linear trendline to the first series, outputting a success message to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Ensure the chart contains at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Add a linear trendline to the first series using reflection (covers versions where Trendlines may not be directly exposed)
            Series series = chart.NSeries[0];
            var trendlinesProp = series.GetType().GetProperty("Trendlines");
            if (trendlinesProp != null)
            {
                var trendlinesObj = trendlinesProp.GetValue(series);
                var addMethod = trendlinesObj?.GetType().GetMethod("Add", new[] { typeof(TrendlineType) });
                if (addMethod != null)
                {
                    var trendline = addMethod.Invoke(trendlinesObj, new object[] { TrendlineType.Linear });
                    Console.WriteLine("Linear trendline added to the first series.");
                }
                else
                {
                    Console.WriteLine("Unable to add trendline: Add method not found.");
                }
            }
            else
            {
                Console.WriteLine("Trendlines property not available in this Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
