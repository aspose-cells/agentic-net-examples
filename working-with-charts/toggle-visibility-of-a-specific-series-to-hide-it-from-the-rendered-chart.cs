// Title: Toggle visibility of a specific series in an Excel chart using Aspose.Cells for .NET
// AI Prompts: Generate C# code that hides or shows a chosen series in an Aspose.Cells chart, using reflection when the IsVisible property is unavailable. | Create a reusable method that accepts a workbook path, chart index, and series index, then toggles that series' visibility with proper error handling. | Show an example that validates the series index, flips the IsVisible flag via reflection, and saves the updated workbook.
// Common Searches: c# hide specific series in an Aspose.Cells generated Excel chart | how to programmatically toggle chart series visibility with Aspose.Cells .NET | using reflection to access IsVisible property of a chart series in Aspose.Cells | validate series index before changing visibility in Aspose.Cells chart example
// Tags: Aspose.Cells chart series IsVisible reflection | C# hide Excel chart series using Aspose.Cells | validate chart series index Aspose.Cells | modify chart series visibility programmatically Aspose.Cells | save workbook after chart changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, accesses the first chart, validates the requested series index, uses reflection to toggle the series' IsVisible flag when supported, and saves the workbook.
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
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or modify as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found on the first worksheet.");
                    return;
                }

                // Get the first chart on the worksheet
                Chart chart = sheet.Charts[0];

                // Index of the series to toggle (0‑based)
                int seriesIndex = 1; // example: second series

                // Validate the series index
                if (seriesIndex < 0 || seriesIndex >= chart.NSeries.Count)
                {
                    Console.WriteLine($"Error: Series index {seriesIndex} is out of range. Chart contains {chart.NSeries.Count} series.");
                    return;
                }

                // Retrieve the series
                Series series = chart.NSeries[seriesIndex];

                // Attempt to toggle visibility using reflection (covers versions without IsVisible property)
                var visibilityProp = series.GetType().GetProperty("IsVisible");
                if (visibilityProp != null && visibilityProp.PropertyType == typeof(bool))
                {
                    bool current = (bool)visibilityProp.GetValue(series);
                    visibilityProp.SetValue(series, !current);
                    Console.WriteLine($"Series visibility toggled to {!current}.");
                }
                else
                {
                    Console.WriteLine("Series visibility property not available in this Aspose.Cells version.");
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
