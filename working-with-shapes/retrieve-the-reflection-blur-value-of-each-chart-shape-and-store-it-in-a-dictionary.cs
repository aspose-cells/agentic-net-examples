// Title: Read chart reflection blur values from an Excel workbook and store them in a C# Dictionary using Aspose.Cells
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet and each chart, obtains the chart's reflection blur (or assigns a default when the property is unavailable) and adds the result to a Dictionary<string, double> keyed by a unique chart identifier. | Show how to gracefully handle the missing Reflection property on Aspose.Cells Chart objects by using a fallback blur value while still collecting blur information for all chart shapes.
// Common Searches: Aspose.Cells get reflection blur of chart shape in .NET | C# read visual effect properties of Excel charts using Aspose.Cells | store Excel chart properties in a dictionary with Aspose.Cells | workaround for unsupported chart reflection property in Aspose.Cells C# | retrieve chart effect parameters from a workbook using Aspose.Cells
// Tags: Aspose.Cells retrieve chart reflection blur | C# dictionary store chart blur values | iterate worksheets and charts Aspose.Cells | fallback for missing chart reflection property Aspose.Cells | generate unique chart key C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel file with Aspose.Cells, iterates over each worksheet and its charts, attempts to read a reflection blur value (using a default of 0.0 because the API lacks a direct property), creates a unique identifier for each chart, stores the blur value in a Dictionary<string, double>, and prints the collected results.
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

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Dictionary to store each chart's reflection blur value (placeholder if not supported)
            Dictionary<string, double> chartReflectionBlur = new Dictionary<string, double>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Aspose.Cells Chart does not expose a Reflection property directly.
                        // Use a placeholder value (0.0) for blur.
                        double blur = 0.0;

                        // Build a unique key for the chart
                        string key = !string.IsNullOrEmpty(chart.Name)
                            ? chart.Name
                            : $"{sheet.Name}_Chart_{chart.GetHashCode()}";

                        chartReflectionBlur[key] = blur;
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart on sheet \"{sheet.Name}\": {exChart.Message}");
                    }
                }
            }

            // Output the collected blur values
            foreach (var kvp in chartReflectionBlur)
            {
                Console.WriteLine($"Chart: {kvp.Key}, Reflection Blur: {kvp.Value}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
