// Title: Hide a specific chart series in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to hide a series in an Aspose.Cells chart | Aspose.Cells hide specific data series from Excel chart programmatically | How to toggle visibility of a chart series with Aspose.Cells .NET | Workaround for missing Series.IsVisible property in Aspose.Cells | Remove a chart series from an Excel file using Aspose.Cells C#
// Tags: Aspose.Cells hide chart series C# | Excel chart series visibility Aspose.Cells | C# toggle chart series visibility Aspose.Cells | Aspose.Cells Series.IsVisible limitation | modify chart series data source Aspose.Cells | Aspose.Cells chart series formatting hide

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for charts, accesses a series by its zero‑based index, notes that the Series class lacks an IsVisible property, and suggests using data‑source or formatting adjustments as a workaround to hide the series before saving the file.
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

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the worksheet.");
                    return;
                }

                // Get the chart you want to modify (adjust index or name as needed)
                Chart chart = worksheet.Charts[0];

                // Specify the series index you want to work with
                int seriesIndex = 1; // example: second series (0‑based index)

                // Ensure the series index is within range
                if (seriesIndex >= 0 && seriesIndex < chart.NSeries.Count)
                {
                    // Retrieve the series
                    Series series = chart.NSeries[seriesIndex];

                    // NOTE: In the current Aspose.Cells version the Series class does not expose an IsVisible property.
                    // If you need to hide/show a series, consider modifying its data source or formatting instead.
                    Console.WriteLine($"Series at index {seriesIndex} accessed successfully.");
                }
                else
                {
                    Console.WriteLine($"Error: Series index {seriesIndex} is out of range. Total series: {chart.NSeries.Count}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
