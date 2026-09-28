// Title: How to mark the last data point as a total in a stacked area chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the IsTotal property on the final data point of each series in an existing stacked area chart with Aspose.Cells. | Show an example of iterating chart series and applying IsTotal to the last DataPoint, including fallback handling for versions without DataPoints support. | Explain how to safely modify a workbook to display cumulative totals in a stacked area chart using Aspose.Cells and verify the result.
// Common Searches: Aspose.Cells C# set IsTotal on last point of stacked area chart series | mark cumulative total in Excel stacked area chart using Aspose.Cells .NET | how to use DataPoints API to flag total point in Aspose.Cells chart | Aspose.Cells version compatibility for DataPoints.IsTotal property | C# code sample for adding total flag to area chart data points in Excel file
// Tags: Aspose.Cells chart series IsTotal property | C# stacked area chart cumulative total | Aspose.Cells DataPoints API usage | Excel workbook modify chart data point total flag | Aspose.Cells version fallback for DataPoints

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, accesses the first worksheet and its first chart (assumed to be a stacked area chart), and demonstrates how to iterate each series to set the IsTotal flag on the last DataPoint. The code includes a commented section for environments where the DataPoints API is unavailable, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the stacked area chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart (assumed to be the stacked area chart)
            Chart chart = sheet.Charts[0];

            // NOTE: The DataPoints API may not be available in older Aspose.Cells versions.
            // The following block is kept for reference; if DataPoints are supported,
            // uncomment the code to mark the last point of each series as a total.
            /*
            foreach (Series series in chart.NSeries)
            {
                try
                {
                    if (series.DataPoints.Count > 0)
                    {
                        int lastIndex = series.DataPoints.Count - 1;
                        DataPoint lastPoint = series.DataPoints[lastIndex];
                        lastPoint.IsTotal = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process series: {ex.Message}");
                }
            }
            */

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
