// Title: How to filter out Excel chart series that never exceed a numeric threshold using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, scans each chart series, and removes any series whose data values are all ≤ 50. | Show how to retrieve a series' data range, evaluate each cell against a threshold, and delete the series from the chart using Aspose.Cells. | Provide a complete Aspose.Cells example that filters chart series by a numeric limit and saves the modified workbook.
// Common Searches: Aspose.Cells C# filter chart series by value threshold | remove Excel chart series that do not exceed a certain number using Aspose.Cells | iterate over NSeries in Aspose.Cells and delete series based on cell values | C# code to hide chart series with low values in an Excel file with Aspose.Cells | how to programmatically prune chart data in Aspose.Cells .NET
// Tags: Aspose.Cells chart series filtering C# | delete low-value series Excel Aspose.Cells | NSeries iteration Aspose.Cells | threshold-based chart pruning .NET | save workbook after chart changes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads Input.xlsx, checks the first worksheet for a chart, iterates each series to see if any numeric cell exceeds 50, removes series that never exceed the threshold, and saves the result as Output.xlsx.
class ChartDataFilter
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart (adjust index if needed)
            Chart chart = sheet.Charts[0];

            double threshold = 50.0;
            List<int> seriesToRemove = new List<int>();

            // Iterate through all series in the chart
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];

                // Address of the data range for this series (e.g., "B2:B10")
                string valuesRangeAddress = series.Values;

                // Obtain the range object from the worksheet
                Aspose.Cells.Range valuesRange = sheet.Cells.CreateRange(valuesRangeAddress);

                bool exceedsThreshold = false;

                // Scan each cell in the range
                foreach (Cell cell in valuesRange)
                {
                    if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue > threshold)
                    {
                        exceedsThreshold = true;
                        break; // No need to continue scanning this series
                    }
                }

                // Mark series for removal if it does not exceed the threshold
                if (!exceedsThreshold)
                {
                    seriesToRemove.Add(i);
                }
            }

            // Remove unwanted series starting from the highest index
            for (int i = seriesToRemove.Count - 1; i >= 0; i--)
            {
                chart.NSeries.RemoveAt(seriesToRemove[i]);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
