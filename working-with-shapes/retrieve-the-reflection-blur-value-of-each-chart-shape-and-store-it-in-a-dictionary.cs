// Title: How to read the reflection blur setting of each chart shape and store it in a C# dictionary using Aspose.Cells
// Common Searches: Aspose.Cells get chart shape reflection blur value in C# | C# read reflection blur property of Excel chart shapes with Aspose.Cells | store chart reflection blur amounts in a dictionary using Aspose.Cells .NET | iterate worksheets and charts to extract shape reflection settings Aspose.Cells | retrieve Excel chart reflection blur using Aspose.Cells API
// Tags: Aspose.Cells chart shape reflection extraction | C# dictionary of chart blur values | iterate worksheets and charts Aspose.Cells | read shape reflection settings .NET | Excel chart reflection property Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, loops through every worksheet and its charts, obtains the Shape object for each chart, reads the Reflection.Blur property when present, and saves the blur amount in a Dictionary keyed by worksheet name and chart index before printing the results.
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

        // Dictionary to store each chart's reflection blur value
        // Key: worksheet name + chart index, Value: blur amount
        var chartReflectionBlur = new Dictionary<string, double>();

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartIdx = 0; // local index for charts in this sheet

                // Iterate through all charts in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Attempt to get the shape that represents the chart by name
                    Shape chartShape = null;
                    try
                    {
                        chartShape = sheet.Shapes[chart.Name];
                    }
                    catch
                    {
                        // Shape not found; continue without throwing
                    }

                    // Ensure the shape and its reflection settings exist
                    if (chartShape != null && chartShape.Reflection != null)
                    {
                        double blurValue = chartShape.Reflection.Blur;

                        // Build a unique key for the chart
                        string key = $"{sheet.Name}_Chart{chartIdx}";

                        // Store the blur value in the dictionary
                        chartReflectionBlur[key] = blurValue;
                    }

                    chartIdx++;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing charts: {ex.Message}");
            return;
        }

        // Output the collected values
        foreach (var kvp in chartReflectionBlur)
        {
            Console.WriteLine($"Chart: {kvp.Key}, Reflection Blur: {kvp.Value}");
        }
    }
}
