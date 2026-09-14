// Title: How to programmatically hide or show specific series in a Waterfall chart using Aspose.Cells for C#
// AI Prompts: Write C# code that loads an Excel workbook, finds a Waterfall chart, and sets the IsVisible flag of each series according to a Dictionary<int, bool> of user preferences. | Demonstrate using reflection in C# to assign the Series.IsVisible property when the property is not directly exposed by the Aspose.Cells API. | Provide a complete example that reads a workbook, modifies series visibility in a Waterfall chart, and saves the updated file to a new location.
// Common Searches: C# Aspose.Cells hide specific series in a waterfall chart | programmatically set series visibility in Aspose.Cells chart | use reflection to change chart series IsVisible property Aspose.Cells | toggle waterfall chart series based on user settings in .NET | Aspose.Cells chart series visibility dictionary example
// Tags: Aspose.Cells set waterfall series visibility C# | chart series IsVisible property reflection Aspose.Cells | toggle chart series visibility programmatically | user preference driven chart series display Aspose.Cells | load and save workbook with modified chart Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook, locates the first Waterfall chart, and iterates through its series. For each series it reads a user‑defined dictionary that maps series indexes to a visibility flag and applies the flag using the Series.IsVisible property via reflection when available. The workbook is then saved with the updated chart visibility settings.
class WaterfallSeriesVisibilityToggle
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the Waterfall chart
            Workbook workbook = new Workbook(inputPath);

            // Assume the Waterfall chart is on the first worksheet and is the first chart object
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Verify that the chart type is Waterfall (optional safety check)
            if (chart.Type != ChartType.Waterfall)
            {
                Console.WriteLine("The first chart is not a Waterfall chart.");
                return;
            }

            // User preferences: key = series index (0‑based), value = desired visibility (true = visible)
            // Example: hide series 1 and show series 0 and 2
            Dictionary<int, bool> userPreferences = new Dictionary<int, bool>()
            {
                { 0, true },
                { 1, false },
                { 2, true }
            };

            // Iterate through all series in the Waterfall chart and set visibility
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                bool makeVisible = true; // default to visible
                if (userPreferences.ContainsKey(i))
                    makeVisible = userPreferences[i];

                Series series = chart.NSeries[i];

                try
                {
                    // Attempt to set the IsVisible property via reflection (available in newer versions)
                    PropertyInfo visibleProp = series.GetType().GetProperty("IsVisible", BindingFlags.Public | BindingFlags.Instance);
                    if (visibleProp != null && visibleProp.CanWrite)
                    {
                        visibleProp.SetValue(series, makeVisible);
                    }
                    else
                    {
                        // If IsVisible is not available, no direct API exists in older versions.
                        // As a best‑effort, we simply leave the series as is.
                        // Optionally, you could manipulate data labels or values here if needed.
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set visibility for series {i}: {ex.Message}");
                }
            }

            // Save the workbook with the updated chart visibility settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
