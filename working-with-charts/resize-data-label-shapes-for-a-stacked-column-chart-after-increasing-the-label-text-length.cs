// Title: Enable and style data labels on a stacked column chart with Aspose.Cells for .NET
// Common Searches: asp.net how to enable data labels on a stacked column chart using Aspose.Cells | increase data label font size and make it bold in Aspose.Cells chart series C# | show values in chart data labels programmatically with Aspose.Cells | use reflection to set HasDataLabels property for compatibility with older Aspose.Cells versions | apply formatting to chart data labels across all series in a workbook with Aspose.Cells
// Tags: enable data labels Aspose.Cells chart | set data label font size C# Aspose.Cells | stacked column chart label styling .NET | reflection workaround HasDataLabels Aspose.Cells | format chart series data labels programmatically | apply bold font to chart data labels Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The code loads an existing workbook, accesses the first worksheet’s stacked column chart, iterates through each series, enables data labels (using reflection for older Aspose.Cells versions), configures the labels to show values, sets the label font to 12‑point bold, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook that contains a stacked column chart
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found on the first worksheet.");
                    return;
                }

                // Get the first chart on the sheet (assumed to be a stacked column chart)
                Chart chart = sheet.Charts[0];

                // Ensure the chart has at least one series
                if (chart.NSeries.Count == 0)
                {
                    Console.WriteLine("Error: The chart does not contain any series.");
                    return;
                }

                // Iterate through each series in the chart
                for (int s = 0; s < chart.NSeries.Count; s++)
                {
                    Series series = chart.NSeries[s];

                    // Enable data labels if the API supports it (guarded by null checks)
                    if (series.DataLabels != null)
                    {
                        // Show values on data labels
                        // Note: Some older Aspose.Cells versions may not expose HasDataLabels or IsValueShown.
                        // The following properties are set only when available.
                        try
                        {
                            // Attempt to enable data labels via reflection for compatibility with older versions
                            var hasDataLabelsProp = series.GetType().GetProperty("HasDataLabels");
                            if (hasDataLabelsProp != null && hasDataLabelsProp.CanWrite)
                            {
                                hasDataLabelsProp.SetValue(series, true);
                            }

                            var isValueShownProp = series.DataLabels.GetType().GetProperty("IsValueShown");
                            if (isValueShownProp != null && isValueShownProp.CanWrite)
                            {
                                isValueShownProp.SetValue(series.DataLabels, true);
                            }
                        }
                        catch
                        {
                            // Ignore reflection errors; continue with formatting only
                        }

                        // Apply formatting to the data labels for better visibility
                        series.DataLabels.Font.Size = 12;
                        series.DataLabels.Font.IsBold = true;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
