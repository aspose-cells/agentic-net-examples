// Title: Set chart legend entry font to transparent for series that contain negative values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide legend entry for series with negative numbers | How to make Excel chart legend transparent for negative data points using Aspose.Cells | C# code to change legend font color based on series values in Aspose.Cells chart | Remove legend entry for negative series in an Excel workbook with Aspose.Cells .NET | Conditional legend formatting in Aspose.Cells chart based on data sign
// Tags: transparent legend font Aspose.Cells chart | negative value detection in chart series C# | modify legend entry color Aspose.Cells .NET | conditional legend formatting Aspose.Cells | iterate chart series Aspose.Cells C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through each chart series, detects if a series contains any negative values, and sets the corresponding legend entry's font color to transparent before saving the workbook.
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet and its first chart
                Worksheet worksheet = workbook.Worksheets[0];
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the first worksheet.");
                    return;
                }

                Chart chart = worksheet.Charts[0];

                // Iterate through each series (legend entry) in the chart
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];

                    // Determine if the series contains any negative values
                    bool containsNegative = false;
                    foreach (object value in series.Values)
                    {
                        if (value is double d && d < 0)
                        {
                            containsNegative = true;
                            break;
                        }
                        if (value is int intVal && intVal < 0)
                        {
                            containsNegative = true;
                            break;
                        }
                    }

                    // If the series contains negative values, make its legend entry transparent
                    if (containsNegative && i < chart.Legend.LegendEntries.Count)
                    {
                        LegendEntry legendEntry = chart.Legend.LegendEntries[i];
                        // Set legend entry font color to transparent
                        legendEntry.Font.Color = Color.Transparent;
                    }
                }

                // Save the modified workbook
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
